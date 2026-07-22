import unittest
from datetime import date

from gst_reconciliation import reconcile_records


class ReconciliationTests(unittest.TestCase):
    def test_exact_match_within_tolerance_is_marked_match(self):
        b2b = [{
            "gstin": "27ABCDE1234F1Z5",
            "invoice_number": "INV-1",
            "invoice_date": date(2026, 6, 10),
            "igst": 100.0,
            "cgst": 50.0,
            "sgst": 50.0,
        }]
        zoho = [{
            "gstin": "27ABCDE1234F1Z5",
            "invoice_number": "INV-1",
            "invoice_date": date(2026, 6, 10),
            "igst": 100.4,
            "cgst": 49.5,
            "sgst": 50.0,
        }]

        result = reconcile_records(b2b, zoho, tolerance=1.0)

        self.assertEqual(result.b2b_rows[0].status, "MATCH")
        self.assertEqual(result.b2b_rows[0].reason, "All tax amounts matched")

    def test_matching_record_with_blank_zoho_tax_is_red(self):
        b2b = [{
            "gstin": "27ABCDE1234F1Z5",
            "invoice_number": "INV-2",
            "invoice_date": date(2026, 6, 11),
            "igst": 100.0,
            "cgst": 50.0,
            "sgst": 50.0,
        }]
        zoho = [{
            "gstin": "27ABCDE1234F1Z5",
            "invoice_number": "INV-2",
            "invoice_date": date(2026, 6, 11),
            "igst": None,
            "cgst": None,
            "sgst": None,
        }]

        result = reconcile_records(b2b, zoho, tolerance=1.0)

        self.assertEqual(result.b2b_rows[0].status, "ZOHO_TAX_BLANK")
        self.assertEqual(result.b2b_rows[0].reason, "Zoho tax blank")

    def test_unmatched_previous_month_b2b_row_is_green_status(self):
        b2b = [
            {
                "gstin": "27ABCDE1234F1Z5",
                "invoice_number": "INV-OLD",
                "invoice_date": date(2026, 5, 20),
                "igst": 10.0,
                "cgst": 5.0,
                "sgst": 5.0,
            },
            {
                "gstin": "27ABCDE1234F1Z5",
                "invoice_number": "INV-NEW",
                "invoice_date": date(2026, 6, 20),
                "igst": 10.0,
                "cgst": 5.0,
                "sgst": 5.0,
            },
        ]
        zoho = [{
            "gstin": "27ABCDE1234F1Z5",
            "invoice_number": "INV-NEW",
            "invoice_date": date(2026, 6, 20),
            "igst": 10.0,
            "cgst": 5.0,
            "sgst": 5.0,
        }]

        result = reconcile_records(b2b, zoho, tolerance=1.0)

        old_invoice = next(r for r in result.b2b_rows if r.cleaned["invoice_number"] == "INV-OLD")
        self.assertEqual(old_invoice.status, "PREVIOUS_MONTH")
        self.assertEqual(old_invoice.reason, "Previous month's invoice")


if __name__ == "__main__":
    unittest.main()
