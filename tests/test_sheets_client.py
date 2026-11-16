import unittest
from unittest.mock import Mock

from importer import ImportValidationError
from sheets_client import write_header


class HeaderTest(unittest.TestCase):
    def test_empty_worksheet_gets_one_header(self):
        worksheet = Mock()
        worksheet.row_values.return_value = []
        write_header(worksheet, ["A", "B"])
        worksheet.append_row.assert_called_once_with(["A", "B"], value_input_option="RAW")

    def test_header_only_worksheet_does_not_get_duplicate_header(self):
        worksheet = Mock()
        worksheet.row_values.return_value = ["A", "B"]
        write_header(worksheet, ["A", "B"])
        worksheet.append_row.assert_not_called()

    def test_changed_header_blocks_writing(self):
        worksheet = Mock()
        worksheet.row_values.return_value = ["B", "A"]
        with self.assertRaises(ImportValidationError):
            write_header(worksheet, ["A", "B"])
        worksheet.append_row.assert_not_called()
