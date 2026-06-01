// e2e/cases/meetings.zig — Smoke tests for the transcript tools.
//
// We can't write a true round-trip test for transcripts: real ones only
// exist when someone has actually attended a Teams meeting that recorded a
// transcript, and there's no Graph API to synthesize one. So these tests
// exercise the validation paths, which is what catches the bulk of
// registration/wiring regressions (missing dispatch entry, broken
// resolveMeetingId branching, schema mismatch).

const std = @import("std");

const client_mod = @import("../client.zig");
const runner = @import("../runner.zig");

const McpClient = client_mod.McpClient;
const pass = runner.pass;
const fail = runner.fail;

/// list-meeting-transcripts with no identifier args must explain that one is
/// required — not crash, not silently return an empty list.
pub fn testListMeetingTranscriptsValidation(client: *McpClient) !void {
    const parsed = try client.callTool("list-meeting-transcripts", "{}");
    defer parsed.deinit();

    const text = McpClient.getResultText(parsed) orelse {
        fail("list-meeting-transcripts (no args)", "no text in response");
        return;
    };

    if (std.mem.indexOf(u8, text, "meetingId") != null and
        std.mem.indexOf(u8, text, "joinUrl") != null and
        std.mem.indexOf(u8, text, "eventId") != null)
    {
        pass("list-meeting-transcripts (no args validation)");
    } else {
        fail("list-meeting-transcripts (no args)", text);
    }
}

/// list-meeting-transcripts called with two identifier args must reject the
/// ambiguity instead of guessing which to use.
pub fn testListMeetingTranscriptsConflict(client: *McpClient) !void {
    const parsed = try client.callTool(
        "list-meeting-transcripts",
        "{\"meetingId\":\"x\",\"eventId\":\"y\"}",
    );
    defer parsed.deinit();

    const text = McpClient.getResultText(parsed) orelse {
        fail("list-meeting-transcripts (conflict)", "no text in response");
        return;
    };

    if (std.mem.indexOf(u8, text, "only one") != null) {
        pass("list-meeting-transcripts (conflict rejected)");
    } else {
        fail("list-meeting-transcripts (conflict)", text);
    }
}

/// get-meeting-transcript without transcriptId must surface the missing-arg
/// message rather than 404 against Graph.
pub fn testGetMeetingTranscriptMissingId(client: *McpClient) !void {
    const parsed = try client.callTool(
        "get-meeting-transcript",
        "{\"meetingId\":\"x\"}",
    );
    defer parsed.deinit();

    const text = McpClient.getResultText(parsed) orelse {
        fail("get-meeting-transcript (no transcriptId)", "no text in response");
        return;
    };

    if (std.mem.indexOf(u8, text, "transcriptId") != null) {
        pass("get-meeting-transcript (missing transcriptId)");
    } else {
        fail("get-meeting-transcript (no transcriptId)", text);
    }
}
