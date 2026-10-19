using System;
using System.Globalization;

namespace KSC.Excel.Engine
{
    /// <summary>Persian (Jalali) calendar helpers used by expressions, filters and parameters. Dates in the file stay Gregorian.</summary>
    public static class Jalali
    {
        static readonly PersianCalendar Pc = new PersianCalendar();

        public static readonly string[] MonthNames =
            { "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند" };

        public static readonly string[] DayNames = { "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه" };

        public static bool TryParts(DateTime dt, out int y, out int m, out int d)
        {
            y = m = d = 0;
            try { y = Pc.GetYear(dt); m = Pc.GetMonth(dt); d = Pc.GetDayOfMonth(dt); return true; }
            catch (ArgumentOutOfRangeException) { return false; }
        }

        /// <summary>Jalali year/month/day to a Gregorian date; null when invalid.</summary>
        public static DateTime? ToDate(int y, int m, int d)
        {
            try { return Pc.ToDateTime(y, m, d, 0, 0, 0, 0); }
            catch (ArgumentOutOfRangeException) { return null; }
        }

        /// <summary>Parses "1404/03/15", "1404-3-5" (also with Persian/Arabic digits); returns false when it is not a valid Jalali date.</summary>
        public static bool TryParse(string s, out DateTime dt)
        {
            dt = default(DateTime);
            if (string.IsNullOrWhiteSpace(s)) return false;
            var chars = s.Trim().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] >= '۰' && chars[i] <= '۹') chars[i] = (char)('0' + (chars[i] - '۰'));
                else if (chars[i] >= '٠' && chars[i] <= '٩') chars[i] = (char)('0' + (chars[i] - '٠'));
            }
            var parts = new string(chars).Split(new[] { '/', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);
            int y, m, d;
            if (parts.Length != 3 || !int.TryParse(parts[0], out y) || !int.TryParse(parts[1], out m) || !int.TryParse(parts[2], out d)) return false;
            if (y < 1000 || y > 1999) return false;            // Gregorian-looking years are not Jalali input
            var r = ToDate(y, m, d);
            if (r == null) return false;
            dt = r.Value; return true;
        }

        public static string MonthName(int m) { return m >= 1 && m <= 12 ? MonthNames[m - 1] : ""; }

        public static int Quarter(int m) { return (m - 1) / 3 + 1; }
    }
}
