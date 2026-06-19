import sys
import unittest
from datetime import date, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from generate_billable_hours_from_pdf import (  # noqa: E402
    AttendanceDay,
    compute_billable_for_known_days,
    parse_rain_dates,
)


class RainDatesTests(unittest.TestCase):
    def test_parses_both_supported_formats(self):
        self.assertEqual(
            parse_rain_dates("2026-05-04, 12/05/2026"),
            {date(2026, 5, 4), date(2026, 5, 12)},
        )

    def test_rejects_invalid_date_with_clear_message(self):
        with self.assertRaisesRegex(ValueError, "Fechas de lluvia inválidas: 31/02/2026"):
            parse_rain_dates("31/02/2026")

    def test_normal_and_rain_thresholds_preserve_billable_hours(self):
        days = [
            AttendanceDay(date(2026, 5, 4), time(8, 30), time(19, 0), "Real", ""),
            AttendanceDay(date(2026, 5, 5), time(8, 45), time(19, 0), "Real", ""),
            AttendanceDay(date(2026, 5, 6), time(8, 46), time(19, 0), "Real", ""),
        ]
        billable, events = compute_billable_for_known_days(
            days, 8, 3, 8, 25, time(8, 0), time(8, 16), time(16, 0), 3,
            {date(2026, 5, 5), date(2026, 5, 6)},
        )
        self.assertEqual(events[date(2026, 5, 4)], 1)
        self.assertEqual(events[date(2026, 5, 5)], 0)
        self.assertEqual(events[date(2026, 5, 6)], 1)
        self.assertGreater(billable[date(2026, 5, 5)], 0)

    def test_rain_arrival_does_not_consume_quota_or_trigger_fourth_late_exclusion(self):
        days = [
            AttendanceDay(date(2026, 5, day), time(8, 20), time(19, 0), "Real", "")
            for day in (4, 5, 6)
        ]
        days.extend(
            [
                AttendanceDay(date(2026, 5, 7), time(8, 45), time(19, 0), "Real", ""),
                AttendanceDay(date(2026, 5, 8), time(8, 20), time(19, 0), "Real", ""),
            ]
        )
        billable, events = compute_billable_for_known_days(
            days, 8, 3, 20, 25, time(8, 0), time(8, 16), time(16, 0), 3,
            {date(2026, 5, 7)},
        )
        self.assertEqual(events[date(2026, 5, 7)], 0)
        self.assertGreater(billable[date(2026, 5, 7)], 0)
        self.assertEqual(billable[date(2026, 5, 8)], 0)


if __name__ == "__main__":
    unittest.main()
