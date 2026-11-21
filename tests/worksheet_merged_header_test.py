import unittest
from unittest.mock import Mock

from gspread.exceptions import InvalidInputValue, MergedHeaderError
from gspread.http_client import HTTPClient
from gspread.worksheet import MergedHeaderRecords, Worksheet


class GetMergedHeaderRecordsTest(unittest.TestCase):
    """Tests for Worksheet.get_merged_header_records with a mocked HTTP client."""

    def _make_worksheet(self):
        client = HTTPClient(auth=None, session=Mock())
        client.values_get = Mock()
        client.values_batch_get = Mock()
        client.fetch_sheet_metadata = Mock()
        properties = {
            "sheetId": 1234,
            "title": "Sheet1",
            "index": 0,
            "gridProperties": {"rowCount": 100, "columnCount": 26},
        }
        worksheet = Worksheet(
            spreadsheet=None,
            properties=properties,
            spreadsheet_id="abc123",
            client=client,
        )
        return worksheet, client

    def _set_metadata(self, client, merges):
        client.fetch_sheet_metadata.return_value = {
            "sheets": [
                {
                    "properties": {"sheetId": 1234, "title": "Sheet1"},
                    "merges": merges,
                }
            ]
        }

    def test_records_with_merged_header(self):
        """Full flow: window offset, ragged API values, out-of-window anchor,
        zero/false preservation and blank padding."""
        worksheet, client = self._make_worksheet()

        # read window B2:E6, 2 header rows (rows 2-3), data rows 4-6
        # merges: A1:B3 (anchor A1 outside the window), C2:D2
        merges = [
            {
                "startRowIndex": 0,
                "endRowIndex": 3,
                "startColumnIndex": 0,
                "endColumnIndex": 2,
            },
            {
                "startRowIndex": 1,
                "endRowIndex": 2,
                "startColumnIndex": 2,
                "endColumnIndex": 4,
            },
        ]
        self._set_metadata(client, merges)

        # the API omits trailing blank cells and trailing blank rows
        client.values_get.return_value = {
            "range": "'Sheet1'!B2:E6",
            "majorDimension": "ROWS",
            "values": [
                ["", "price", "", "qty"],
                ["", "unit", "total"],
                ["apple", 3, 30, 10],
                ["pear", 0, 0, False],
                # row 6 fully blank: omitted by the API
            ],
        }
        # the true top-left cell of A1:B3 lives outside the read window
        client.values_batch_get.return_value = {
            "valueRanges": [
                {
                    "range": "'Sheet1'!A1",
                    "majorDimension": "ROWS",
                    "values": [["product"]],
                }
            ]
        }

        records = worksheet.get_merged_header_records("B2:E6", header_rows=2)

        self.assertIsInstance(records, MergedHeaderRecords)
        self.assertEqual(records.range, "B2:E6")
        self.assertEqual(records.header_rows, 2)
        self.assertEqual(records.first_data_row, 4)

        # the out-of-window anchor was fetched with a single extra request
        client.values_batch_get.assert_called_once()
        _, kwargs = client.values_batch_get.call_args
        self.assertEqual(kwargs["ranges"], ["'Sheet1'!A1"])

        # header paths: vertical merge recorded once, level positions kept
        paths = {c.column: c.path for c in records.columns}
        self.assertEqual(paths["B"], ("product", None))
        self.assertEqual(paths["C"], ("price", "unit"))
        self.assertEqual(paths["D"], ("price", "total"))
        self.assertEqual(paths["E"], ("qty", ""))

        # label sources use absolute coordinates
        label_b = records.columns[0].labels[0]
        self.assertEqual(label_b.address, "A1")
        self.assertEqual((label_b.row, label_b.column), (1, 1))
        self.assertEqual((label_b.level, label_b.level_span), (1, 2))

        # data aligned on the original rows/columns, blanks padded without
        # shifting, zero and false preserved
        self.assertEqual(
            list(records),
            [
                {"B": "apple", "C": 3, "D": 30, "E": 10},
                {"B": "pear", "C": 0, "D": 0, "E": False},
                {"B": "", "C": "", "D": "", "E": ""},
            ],
        )

    def test_same_path_columns_output_independently(self):
        """A 2D merge gives identical paths; columns still output by address."""
        worksheet, client = self._make_worksheet()

        merges = [
            {
                "startRowIndex": 0,
                "endRowIndex": 2,
                "startColumnIndex": 0,
                "endColumnIndex": 2,
            }
        ]
        self._set_metadata(client, merges)
        client.values_get.return_value = {
            "range": "'Sheet1'!A1:B3",
            "majorDimension": "ROWS",
            "values": [["merged"], [""], ["a", "b"]],
        }

        records = worksheet.get_merged_header_records("A1:B3", header_rows=2)

        self.assertEqual([c.path for c in records.columns], [("merged", None)] * 2)
        self.assertEqual([c.column for c in records.columns], ["A", "B"])
        self.assertEqual(list(records), [{"A": "a", "B": "b"}])
        # all merges start inside the window: no extra request
        client.values_batch_get.assert_not_called()

    def test_merge_crossing_header_boundary_raises(self):
        worksheet, client = self._make_worksheet()

        # merge A2:A3 crosses the boundary between 2 header rows and data
        merges = [
            {
                "startRowIndex": 1,
                "endRowIndex": 3,
                "startColumnIndex": 0,
                "endColumnIndex": 1,
            }
        ]
        self._set_metadata(client, merges)
        client.values_get.return_value = {
            "range": "'Sheet1'!A1:A3",
            "majorDimension": "ROWS",
            "values": [["h"], ["h"], ["d"]],
        }

        with self.assertRaises(MergedHeaderError):
            worksheet.get_merged_header_records("A1:A3", header_rows=2)

    def test_invalid_range_and_header_rows(self):
        worksheet, client = self._make_worksheet()
        self._set_metadata(client, [])

        with self.assertRaises(InvalidInputValue):
            worksheet.get_merged_header_records("A:B", header_rows=1)
        with self.assertRaises(InvalidInputValue):
            worksheet.get_merged_header_records("A1", header_rows=1)
        with self.assertRaises(InvalidInputValue):
            worksheet.get_merged_header_records("named_range", header_rows=1)
        with self.assertRaises(InvalidInputValue):
            worksheet.get_merged_header_records("A1:B2", header_rows=0)
        with self.assertRaises(InvalidInputValue):
            worksheet.get_merged_header_records("A1:B2", header_rows=3)

        # invalid input is rejected before any API request
        client.values_get.assert_not_called()

    def test_default_blank_used_for_padding(self):
        worksheet, client = self._make_worksheet()
        self._set_metadata(client, [])
        client.values_get.return_value = {
            "range": "'Sheet1'!A1:B3",
            "majorDimension": "ROWS",
            "values": [["h1", "h2"], ["only-a"]],
        }

        records = worksheet.get_merged_header_records(
            "A1:B3", header_rows=1, default_blank=None
        )

        self.assertEqual(
            list(records),
            [
                {"A": "only-a", "B": None},
                {"A": None, "B": None},
            ],
        )

    def test_empty_range_values(self):
        """The API returns no values at all: records are padded blanks."""
        worksheet, client = self._make_worksheet()
        self._set_metadata(client, [])
        client.values_get.return_value = {
            "range": "'Sheet1'!A1:B2",
            "majorDimension": "ROWS",
        }

        records = worksheet.get_merged_header_records("A1:B2", header_rows=1)

        self.assertEqual([c.path for c in records.columns], [("",), ("",)])
        self.assertEqual(list(records), [{"A": "", "B": ""}])
