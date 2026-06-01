// tools/meetings.zig — Microsoft Teams online-meeting transcript tools.
//
// Two tools live here:
//
//   list-meeting-transcripts — enumerate transcripts available for a meeting
//     (GET /me/onlineMeetings/{meetingId}/transcripts).
//   get-meeting-transcript   — download a transcript's VTT content
//     (GET /me/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content).
//
// Identifying a meeting is the awkward part: the onlineMeeting `id` is NOT
// the same as the calendar event id, and the LLM almost never has it. So
// every handler accepts one of three inputs and resolves to the meetingId:
//
//   - `meetingId` — already the onlineMeeting id (no lookup needed).
//   - `joinUrl`   — Teams join URL; resolved via
//                   GET /me/onlineMeetings?$filter=JoinWebUrl eq '<url>'.
//   - `eventId`   — calendar event id; we GET the event, pull
//                   onlineMeeting.joinUrl, then resolve as above.
//
// Reading guide for TS/Python readers:
//   - We use std.json.parseFromSlice + an arena that dies on parsed.deinit().
//     Any string we want past deinit must be allocator.dupe()'d first.
//   - Single quotes inside an OData string literal must be doubled (`''`).
//     We then URL-encode the whole filter so the HTTP layer sees a clean
//     query value. Without the doubling step, a join URL with a `'` would
//     break the filter (and could be an injection vector).

const std = @import("std");
const graph = @import("../graph.zig");
const json_rpc = @import("../json_rpc.zig");
const url_util = @import("../url.zig");
const ToolContext = @import("context.zig").ToolContext;

const Allocator = std.mem.Allocator;
const Value = std.json.Value;
const ObjectMap = std.json.ObjectMap;

/// Resolve the onlineMeeting id from whichever identifier the caller passed.
///
/// Exactly one of `meetingId`, `joinUrl`, `eventId` must be present on
/// `args`. Returns an allocated meeting id the caller must free. On any
/// failure (missing/conflicting args, event without a Teams meeting,
/// Graph error, etc.) sends an error response and returns null.
fn resolveMeetingId(ctx: ToolContext, args: ObjectMap, token: []const u8) ?[]u8 {
    const meeting_id = json_rpc.getStringArg(args, "meetingId");
    const join_url = json_rpc.getStringArg(args, "joinUrl");
    const event_id = json_rpc.getStringArg(args, "eventId");

    // Count provided inputs. Zig has no boolean-to-int coercion, so we
    // compare to null and sum 0/1 explicitly.
    var provided: u8 = 0;
    if (meeting_id != null) provided += 1;
    if (join_url != null) provided += 1;
    if (event_id != null) provided += 1;

    if (provided == 0) {
        ctx.sendResult("Provide exactly one of: meetingId, joinUrl, or eventId.");
        return null;
    }
    if (provided > 1) {
        ctx.sendResult("Provide only one of meetingId, joinUrl, or eventId — not multiple.");
        return null;
    }

    // Easy path: caller already has the meeting id.
    if (meeting_id) |id| return ctx.allocator.dupe(u8, id) catch null;

    // Otherwise we need a join URL — either supplied directly or pulled
    // from a calendar event.
    const resolved_join_url: []u8 = if (join_url) |u|
        (ctx.allocator.dupe(u8, u) catch return null)
    else
        (joinUrlFromEvent(ctx, token, event_id.?) orelse return null);
    defer ctx.allocator.free(resolved_join_url);

    return meetingIdFromJoinUrl(ctx, token, resolved_join_url);
}

/// Fetch a calendar event and pull `onlineMeeting.joinUrl` from it.
/// Returns an allocated join URL the caller must free, or null with an
/// error already written to the client.
fn joinUrlFromEvent(ctx: ToolContext, token: []const u8, event_id: []const u8) ?[]u8 {
    // Reject ids that could escape the path segment — same rule as
    // ctx.getPathArg uses on direct args.
    for (event_id) |c| switch (c) {
        '/', '?', '&', '#' => {
            ctx.sendResult("Invalid 'eventId' — contains URL-reserved characters.");
            return null;
        },
        else => {},
    };

    const path = std.fmt.allocPrint(
        ctx.allocator,
        "/me/events/{s}?$select=id,subject,isOnlineMeeting,onlineMeeting",
        .{event_id},
    ) catch return null;
    defer ctx.allocator.free(path);

    const response = graph.get(ctx.allocator, ctx.io, token, path) catch |err| {
        ctx.sendGraphError(err);
        return null;
    };
    defer ctx.allocator.free(response);

    // Parse, walk to onlineMeeting.joinUrl. Each parsed-arena string must
    // be dupe()'d before the arena's defer fires.
    const parsed = std.json.parseFromSlice(Value, ctx.allocator, response, .{}) catch {
        ctx.sendResult("Unexpected response shape from /me/events.");
        return null;
    };
    defer parsed.deinit();

    const root = switch (parsed.value) {
        .object => |o| o,
        else => {
            ctx.sendResult("Unexpected response shape from /me/events.");
            return null;
        },
    };

    const om_val = root.get("onlineMeeting") orelse {
        ctx.sendResult("This event has no Teams meeting attached — list-meeting-transcripts needs an event where isOnlineMeeting=true. Try passing the joinUrl directly.");
        return null;
    };
    const om = switch (om_val) {
        .object => |o| o,
        else => {
            ctx.sendResult("Event's onlineMeeting field was not an object.");
            return null;
        },
    };

    const join_val = om.get("joinUrl") orelse {
        ctx.sendResult("Event's onlineMeeting has no joinUrl field.");
        return null;
    };
    const join_str = switch (join_val) {
        .string => |s| s,
        else => {
            ctx.sendResult("Event's onlineMeeting.joinUrl was not a string.");
            return null;
        },
    };

    return ctx.allocator.dupe(u8, join_str) catch null;
}

/// Look up an onlineMeeting by its join URL, returning the meeting id.
/// Uses `GET /me/onlineMeetings?$filter=JoinWebUrl eq '<url>'`. The user
/// must be the organizer; otherwise Graph returns an empty result set.
fn meetingIdFromJoinUrl(ctx: ToolContext, token: []const u8, join_url: []const u8) ?[]u8 {
    // OData string literals are delimited by single quotes. Any single
    // quote within the literal must be doubled (`''`). Join URLs from
    // Teams shouldn't contain `'`, but we double them anyway so a crafted
    // URL can't break out of the literal and inject filter syntax.
    const doubled = doubleSingleQuotes(ctx.allocator, join_url) catch return null;
    defer ctx.allocator.free(doubled);

    // Build the OData filter expression, then URL-encode it whole so
    // characters like `/`, `:`, `?` in the URL don't terminate the query.
    const filter_literal = std.fmt.allocPrint(
        ctx.allocator,
        "JoinWebUrl eq '{s}'",
        .{doubled},
    ) catch return null;
    defer ctx.allocator.free(filter_literal);

    const filter_encoded = url_util.encode(ctx.allocator, filter_literal) catch return null;
    defer ctx.allocator.free(filter_encoded);

    const path = std.fmt.allocPrint(
        ctx.allocator,
        "/me/onlineMeetings?$filter={s}",
        .{filter_encoded},
    ) catch return null;
    defer ctx.allocator.free(path);

    const response = graph.get(ctx.allocator, ctx.io, token, path) catch |err| {
        ctx.sendGraphError(err);
        return null;
    };
    defer ctx.allocator.free(response);

    const parsed = std.json.parseFromSlice(Value, ctx.allocator, response, .{}) catch {
        ctx.sendResult("Unexpected response shape from /me/onlineMeetings.");
        return null;
    };
    defer parsed.deinit();

    const root = switch (parsed.value) {
        .object => |o| o,
        else => {
            ctx.sendResult("Unexpected response shape from /me/onlineMeetings.");
            return null;
        },
    };
    const items = switch (root.get("value") orelse .null) {
        .array => |a| a.items,
        else => {
            ctx.sendResult("No onlineMeetings.value array in response.");
            return null;
        },
    };
    if (items.len == 0) {
        ctx.sendResult("No onlineMeeting found for that join URL. Note: /me/onlineMeetings only returns meetings you organize.");
        return null;
    }

    const first_obj = switch (items[0]) {
        .object => |o| o,
        else => {
            ctx.sendResult("First onlineMeeting was not an object.");
            return null;
        },
    };
    const id_val = first_obj.get("id") orelse {
        ctx.sendResult("onlineMeeting result missing 'id' field.");
        return null;
    };
    const id_str = switch (id_val) {
        .string => |s| s,
        else => {
            ctx.sendResult("onlineMeeting 'id' was not a string.");
            return null;
        },
    };
    return ctx.allocator.dupe(u8, id_str) catch null;
}

/// Replace every `'` in `input` with `''`. Returned slice is allocated
/// and owned by the caller.
fn doubleSingleQuotes(allocator: Allocator, input: []const u8) ![]u8 {
    var quote_count: usize = 0;
    for (input) |c| {
        if (c == '\'') quote_count += 1;
    }
    if (quote_count == 0) return allocator.dupe(u8, input);

    const out = try allocator.alloc(u8, input.len + quote_count);
    var i: usize = 0;
    for (input) |c| {
        out[i] = c;
        i += 1;
        if (c == '\'') {
            out[i] = '\'';
            i += 1;
        }
    }
    return out;
}

/// List transcripts available for an online meeting.
/// Graph: GET /me/onlineMeetings/{meetingId}/transcripts
pub fn handleListMeetingTranscripts(ctx: ToolContext) void {
    const token = ctx.requireAuth() orelse return;
    const args = ctx.getArgs("Missing arguments. Provide one of: meetingId, joinUrl, eventId.") orelse return;

    const meeting_id = resolveMeetingId(ctx, args, token) orelse return;
    defer ctx.allocator.free(meeting_id);

    // The meeting id Graph returned is path-safe by construction — it's a
    // base64-ish opaque string with no '/', '?', '&', '#'. But validate
    // defensively so we never paste an attacker-controlled segment into a URL.
    for (meeting_id) |c| switch (c) {
        '/', '?', '&', '#' => {
            ctx.sendResult("Resolved meetingId contained URL-reserved characters — refusing to use.");
            return;
        },
        else => {},
    };

    const path = std.fmt.allocPrint(
        ctx.allocator,
        "/me/onlineMeetings/{s}/transcripts",
        .{meeting_id},
    ) catch return;
    defer ctx.allocator.free(path);

    const response = graph.get(ctx.allocator, ctx.io, token, path) catch |err| {
        ctx.sendGraphError(err);
        return;
    };
    defer ctx.allocator.free(response);

    // Parse and emit one line per transcript. The list endpoint returns:
    //   { "value": [{ "id": "...", "createdDateTime": "...",
    //                 "transcriptContentUrl": "..." }] }
    // We surface only id and createdDateTime — transcriptContentUrl is a
    // pre-auth Graph URL that requires the same token anyway, so the LLM
    // gets nothing by seeing it and we save the context bytes.
    const parsed = std.json.parseFromSlice(Value, ctx.allocator, response, .{}) catch {
        ctx.sendResult("Unexpected response shape from transcripts endpoint.");
        return;
    };
    defer parsed.deinit();

    const root = switch (parsed.value) {
        .object => |o| o,
        else => {
            ctx.sendResult("Unexpected response shape from transcripts endpoint.");
            return;
        },
    };
    const items = switch (root.get("value") orelse .null) {
        .array => |a| a.items,
        else => {
            ctx.sendResult("No transcripts available for this meeting.");
            return;
        },
    };
    if (items.len == 0) {
        ctx.sendResult("No transcripts available for this meeting.");
        return;
    }

    var buf: std.Io.Writer.Allocating = .init(ctx.allocator);
    defer buf.deinit();
    const w = &buf.writer;

    for (items) |item| {
        const obj = switch (item) {
            .object => |o| o,
            else => continue,
        };
        const tid = stringField(obj, "id");
        const created = stringField(obj, "createdDateTime");
        w.writeAll("createdDateTime: ") catch continue;
        w.writeAll(created) catch continue;
        w.writeAll(" | id: ") catch continue;
        w.writeAll(tid) catch continue;
        w.writeAll("\n") catch continue;
    }

    const out = buf.toOwnedSlice() catch return;
    defer ctx.allocator.free(out);
    ctx.sendResult(out);
}

/// Download a meeting transcript as WebVTT text.
/// Graph: GET /me/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content?$format=text/vtt
pub fn handleGetMeetingTranscript(ctx: ToolContext) void {
    const token = ctx.requireAuth() orelse return;
    const args = ctx.getArgs("Missing arguments. Provide transcriptId plus one of: meetingId, joinUrl, eventId.") orelse return;

    const transcript_id = ctx.getPathArg(args, "transcriptId", "Missing 'transcriptId' argument (from list-meeting-transcripts).") orelse return;

    const meeting_id = resolveMeetingId(ctx, args, token) orelse return;
    defer ctx.allocator.free(meeting_id);

    for (meeting_id) |c| switch (c) {
        '/', '?', '&', '#' => {
            ctx.sendResult("Resolved meetingId contained URL-reserved characters — refusing to use.");
            return;
        },
        else => {},
    };

    // text/vtt is the default for this endpoint with delegated permissions;
    // pinning it explicitly future-proofs against Graph changing defaults.
    const path = std.fmt.allocPrint(
        ctx.allocator,
        "/me/onlineMeetings/{s}/transcripts/{s}/content?$format=text/vtt",
        .{ meeting_id, transcript_id },
    ) catch return;
    defer ctx.allocator.free(path);

    const response = graph.get(ctx.allocator, ctx.io, token, path) catch |err| {
        ctx.sendGraphError(err);
        return;
    };
    defer ctx.allocator.free(response);

    if (response.len == 0) {
        ctx.sendResult("Transcript content was empty.");
        return;
    }

    // VTT is text — return inline. Large transcripts (multi-hour meetings)
    // may exceed comfortable context size; if that becomes a problem we'll
    // add a `save: true` flag that routes through binary_download.zig.
    ctx.sendResult(response);
}

/// Pull a string field from a JSON object, returning "" if missing/non-string.
fn stringField(obj: ObjectMap, key: []const u8) []const u8 {
    const v = obj.get(key) orelse return "";
    return switch (v) {
        .string => |s| s,
        else => "",
    };
}

// --- Tests ---

const testing = std.testing;

test "doubleSingleQuotes: no quotes passes through" {
    const out = try doubleSingleQuotes(testing.allocator, "hello world");
    defer testing.allocator.free(out);
    try testing.expectEqualStrings("hello world", out);
}

test "doubleSingleQuotes: single quote becomes two" {
    const out = try doubleSingleQuotes(testing.allocator, "it's");
    defer testing.allocator.free(out);
    try testing.expectEqualStrings("it''s", out);
}

test "doubleSingleQuotes: multiple quotes" {
    const out = try doubleSingleQuotes(testing.allocator, "'a'b'");
    defer testing.allocator.free(out);
    try testing.expectEqualStrings("''a''b''", out);
}

test "doubleSingleQuotes: empty string" {
    const out = try doubleSingleQuotes(testing.allocator, "");
    defer testing.allocator.free(out);
    try testing.expectEqualStrings("", out);
}
