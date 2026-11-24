import unittest

from wex_viewer import JetDatabase


class JetTextDecodingTests(unittest.TestCase):
    def test_compressed_text_switches_between_ascii_and_chinese(self):
        raw = bytes.fromhex("fffe 312d34 00 507d5f67 00 28383029")
        self.assertEqual(JetDatabase.decode_text_value(raw), "1-4結束(80)")

    def test_compressed_text_supports_multiple_mode_switches(self):
        raw = bytes.fromhex("fffe 41 00 b065 00 42 00 8765 00 43")
        self.assertEqual(JetDatabase.decode_text_value(raw), "A新B文C")

    def test_compressed_and_uncompressed_characters_can_contain_high_or_zero_bytes(self):
        raw = bytes.fromhex("fffe e9 00 e900 00 41")
        self.assertEqual(JetDatabase.decode_text_value(raw), "ééA")

    def test_plain_utf16_still_decodes(self):
        self.assertEqual(JetDatabase.decode_text_value("第一節".encode("utf-16-le")), "第一節")

    def test_truncated_uncompressed_character_is_reported(self):
        with self.assertRaises(UnicodeDecodeError):
            JetDatabase.decode_text_value(bytes.fromhex("fffe 41 00 50"))


if __name__ == "__main__":
    unittest.main()
