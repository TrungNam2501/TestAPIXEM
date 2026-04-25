using System;

namespace ServiceBBAPI.Helpers
{
    public static class ShiftHelper
    {
        public static string GetShiftId()
        {
            DateTime now = DateTime.Now;
            DateTime dayStart = now.Date.Add(new TimeSpan(7, 30, 0));
            DateTime dayEnd = now.Date.Add(new TimeSpan(19, 0, 0));
            return (now >= dayStart && now <= dayEnd) ? "1" : "2";
        }

        public static string GetShiftMayBB()
        {
            DateTime now = DateTime.Now;
            DateTime dayStart = now.Date.Add(new TimeSpan(7, 30, 0));
            DateTime dayEnd = now.Date.Add(new TimeSpan(19, 0, 0));
            return (now >= dayStart && now <= dayEnd) ? "0" : "1";
        }

        public static string GetPday(string shiftId)
        {
            DateTime now = DateTime.Now;
            DateTime midnightStart = now.Date;
            DateTime morningCutoff = now.Date.Add(new TimeSpan(7, 30, 0));

            if (shiftId == "2" && now >= midnightStart && now <= morningCutoff)
            {
                return now.AddDays(-1).ToString("yyyyMMdd");
            }
            return now.ToString("yyyyMMdd");
        }

        public static string GetPdayBB(string shiftId)
        {
            DateTime now = DateTime.Now;
            DateTime midnightStart = now.Date;
            DateTime morningCutoff = now.Date.Add(new TimeSpan(7, 30, 0));

            if (shiftId == "1" && now >= midnightStart && now <= morningCutoff)
            {
                return now.AddDays(-1).ToString("yyyyMMdd");
            }
            return now.ToString("yyyyMMdd");
        }

        public static string EncodeMonth(string month)
        {
            switch (month)
            {
                case "10": return "A";
                case "11": return "B";
                case "12": return "C";
                default: return month.Substring(1, 1);
            }
        }

        public static string BuildSpday(DateTime pday)
        {
            string year = pday.ToString("yyyy").Substring(2, 2);
            string month = EncodeMonth(pday.ToString("MM"));
            string day = pday.ToString("dd");
            return year + month + day;
        }
    }
}
