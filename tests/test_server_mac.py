import unittest

from outlook_desktop_mcp.server_mac import _ui_list_messages
from outlook_desktop_mcp.utils.applescript_helpers import RECORD_DELIM


class FallbackBridge:
    def __init__(self):
        self.scripts = []

    async def run(self, script):
        self.scripts.append(script)
        if "tell splitter group 1" in script:
            raise RuntimeError("legacy Outlook hierarchy is unavailable")
        return (
            "Jamie Lee, Q3 Planning Notes,     "
            f"7/24/26,{RECORD_DELIM}"
        )


class UiMessageListTests(unittest.IsolatedAsyncioTestCase):
    async def test_uses_current_outlook_layout_when_legacy_layout_is_unavailable(self):
        bridge = FallbackBridge()

        messages = await _ui_list_messages(bridge, 1)

        self.assertEqual(messages[0]["sender_name"], "Jamie Lee")
        self.assertEqual(len(bridge.scripts), 2)
