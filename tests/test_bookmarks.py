import unittest

import multipane_explorer as explorer


class BookmarkLayoutTests(unittest.TestCase):
    def test_top_row_is_filled_before_bottom_row(self):
        visible, top = explorer._two_row_bookmark_fit([40] * 4, 200, 4, 30)

        self.assertEqual((visible, top), (4, 4))

    def test_overflow_button_is_reserved_on_second_row(self):
        visible, top = explorer._two_row_bookmark_fit([40] * 8, 150, 4, 30)

        self.assertEqual((visible, top), (5, 3))

    def test_empty_bookmark_toolbar_has_no_visible_rows(self):
        self.assertEqual(explorer._two_row_bookmark_fit([], 150, 4, 30), (0, 0))

    def test_long_bookmark_names_have_a_roomier_width_limit(self):
        self.assertGreaterEqual(explorer.QUICK_BOOKMARK_MAX_W, 150)


if __name__ == "__main__":
    unittest.main()
