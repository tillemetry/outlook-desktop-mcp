import unittest

from outlook_desktop_mcp.server_mac import _ui_list_messages
from outlook_desktop_mcp.utils.applescript_helpers import RECORD_DELIM


class FallbackBridge:
    """Fake bridge: fails the legacy layout, serves one canned row otherwise."""

    def __init__(self):
        self.scripts = []

    async def run(self, script):
        self.scripts.append(script)
        if 'starts with "Inbox"' in script:
            return "1"  # window-index probe
        if "tell splitter group 1" in script:
            raise RuntimeError("legacy Outlook hierarchy is unavailable")
        return (
            "Jamie Lee, Q3 Planning Notes,     "
            f"7/24/26,{RECORD_DELIM}"
        )


class RawBridge:
    """Fake bridge that serves a fixed scrape payload for the layout script."""

    def __init__(self, raw):
        self._raw = raw

    async def run(self, script):
        if 'starts with "Inbox"' in script:
            return "1"
        if "tell splitter group 1" in script:
            raise RuntimeError("legacy layout unavailable")
        return self._raw


# Row descriptions modeled on New Outlook 16.111 accessibility output. Content
# is synthetic; the structure (field spacing, importance/count prefixes, meeting
# metadata, multi-sender lists) mirrors the real rows the parser must handle.
_STANDARD = (
    " Unread,    Jane Doe, Weekly Product Update,     "
    "8:17 AM,        Body preview text here"
)
_HEADER = "Today, Expanded"
_CELL = "cell"
_OTHER = (
    "Other Emails from Marketing List, Newsletter Weekly, Events Team, "
    "Community Forum"
)
# Meeting thread: the list time ("Yesterday") sits behind the embedded meeting
# metadata (only 3 spaces), so the standard split lands preview text as the time
# and the row must be salvaged.
_THREAD_MEETING = (
    "High priority, 1 unread message,   4 messages, Alex Kim, "
    "Project Sync Notes,  Meeting message,Mon 7/27/26, 8:30 PM "
    "(1 hour) No Conflicts   Yesterday,        Body preview text here"
)
# Thread whose time parses normally, but the importance/count prefix and the
# multi-sender participant list leak into sender/subject unless stripped.
_THREAD_TIMED = (
    "High priority, 1 unread message,   4 messages, Team Alpha, "
    "Jordan Blake, Riley Poe, Fall Schedule Update,     "
    "9:46 AM,        Body preview text here"
)


class UiMessageListTests(unittest.IsolatedAsyncioTestCase):
    async def test_uses_current_outlook_layout_when_legacy_layout_is_unavailable(self):
        bridge = FallbackBridge()

        messages = await _ui_list_messages(bridge, 1)

        self.assertEqual(messages[0]["sender_name"], "Jamie Lee")
        # window-index probe + legacy attempt + current layout
        self.assertEqual(len(bridge.scripts), 3)

    async def test_drops_non_message_rows(self):
        raw = RECORD_DELIM.join([_HEADER, _CELL, _OTHER, _STANDARD]) + RECORD_DELIM

        messages = await _ui_list_messages(RawBridge(raw), 10)

        self.assertEqual(len(messages), 1)
        self.assertEqual(messages[0]["sender_name"], "Jane Doe")
        self.assertEqual(messages[0]["subject"], "Weekly Product Update")
        self.assertEqual(messages[0]["received_time"], "8:17 AM")

    async def test_salvages_thread_and_meeting_rows(self):
        raw = RECORD_DELIM.join([_STANDARD, _THREAD_MEETING]) + RECORD_DELIM

        messages = await _ui_list_messages(RawBridge(raw), 10)

        self.assertEqual(len(messages), 2)
        thread = messages[1]
        self.assertEqual(thread["sender_name"], "Alex Kim")
        self.assertEqual(thread["subject"], "Project Sync Notes")
        self.assertEqual(thread["received_time"], "Yesterday")

    async def test_strips_importance_and_counts_from_timed_thread(self):
        raw = RECORD_DELIM.join([_THREAD_TIMED]) + RECORD_DELIM

        messages = await _ui_list_messages(RawBridge(raw), 10)

        self.assertEqual(len(messages), 1)
        msg = messages[0]
        self.assertEqual(msg["sender_name"], "Team Alpha")
        self.assertIn("Fall Schedule Update", msg["subject"])
        self.assertEqual(msg["received_time"], "9:46 AM")
        self.assertNotIn("priority", msg["sender_name"].lower())
        self.assertNotIn("message", msg["sender_name"].lower())

    async def test_count_limits_real_messages(self):
        raw = RECORD_DELIM.join([_HEADER, _STANDARD, _STANDARD, _STANDARD]) + RECORD_DELIM

        messages = await _ui_list_messages(RawBridge(raw), 2)

        self.assertEqual(len(messages), 2)
        self.assertTrue(all(m["received_time"] == "8:17 AM" for m in messages))


if __name__ == "__main__":
    unittest.main()
