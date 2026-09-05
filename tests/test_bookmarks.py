import unittest

import multipane_explorer as explorer


class BookmarkLayoutTests(unittest.TestCase):
    def test_bookmarks_are_balanced_across_two_rows(self):
        visible, top = explorer._two_row_bookmark_fit([40] * 6, 150, 4, 30)

        self.assertEqual((visible, top), (6, 3))

    def test_overflow_button_is_reserved_on_second_row(self):
        visible, top = explorer._two_row_bookmark_fit([40] * 8, 150, 4, 30)

        self.assertEqual((visible, top), (5, 3))

    def test_empty_bookmark_toolbar_has_no_visible_rows(self):
        self.assertEqual(explorer._two_row_bookmark_fit([], 150, 4, 30), (0, 0))


if __name__ == "__main__":
    unittest.main()
