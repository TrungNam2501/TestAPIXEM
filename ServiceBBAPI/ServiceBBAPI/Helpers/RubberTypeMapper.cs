using System.Collections.Generic;

namespace ServiceBBAPI.Helpers
{
    public class RubberTypeResult
    {
        public string Makeo { get; set; }
        public string Ptype { get; set; }
        public string NormalizedTenkeo { get; set; }
    }

    public static class RubberTypeMapper
    {
        private static readonly Dictionary<string, (string makeo, string ptype)> TypeMap = new Dictionary<string, (string, string)>
        {
            { "RM",      ("RD", "3") },
            { "1",       ("RB", "2") },
            { "9",       ("RD", "3") },
            { "2",       ("RC", "2") },
            { "3",       ("RC", "2") },
            { "4",       ("RC", "2") },
            { "5",       ("RC", "2") },
            { "RE",      ("RR", "3") },
            { "R",       ("RR", "3") },
            { "92",      ("RD", "3") },
            { "",        ("RD", "3") },
            { "1-EDGE",  ("RB", "2") },
            { "2-EDGE",  ("RC", "2") },
            { "3-EDGE",  ("RC", "2") },
            { "4-EDGE",  ("RC", "2") },
            { "5-EDGE",  ("RC", "2") },
            { "9-EDGE",  ("RD", "3") },
            { "EDGE",    ("RD", "3") },
        };

        private static readonly Dictionary<string, (string suffix, string makeo, string ptype)> EdgeNormMap = new Dictionary<string, (string, string, string)>
        {
            { "1EDGE", ("-1-EDGE", "RB", "2") },
            { "2EDGE", ("-2-EDGE", "RC", "2") },
            { "3EDGE", ("-3-EDGE", "RC", "2") },
            { "4EDGE", ("-4-EDGE", "RC", "2") },
            { "5EDGE", ("-5-EDGE", "RC", "2") },
            { "9EDGE", ("-9-EDGE", "RD", "3") },
        };

        private static readonly Dictionary<string, (string suffix, string makeo, string ptype)> ThuNormMap = new Dictionary<string, (string, string, string)>
        {
            { "1THU", ("-1THU", "RB", "2") },
            { "2THU", ("-2THU", "RC", "2") },
            { "3THU", ("-3THU", "RC", "2") },
            { "4THU", ("-4THU", "RC", "2") },
            { "5THU", ("-5THU", "RC", "2") },
            { "9THU", ("-9THU", "RD", "3") },
        };

        public static RubberTypeResult Resolve(string tenkeo)
        {
            string partno = tenkeo.Trim();

            if (partno.Length <= 5)
            {
                return new RubberTypeResult { Makeo = "RD", Ptype = "3", NormalizedTenkeo = tenkeo };
            }

            string suffix = partno.Substring(6).ToUpper();

            if (TypeMap.TryGetValue(suffix, out var direct))
            {
                return new RubberTypeResult { Makeo = direct.makeo, Ptype = direct.ptype, NormalizedTenkeo = tenkeo };
            }

            if (EdgeNormMap.TryGetValue(suffix, out var edge))
            {
                string normalized = partno.Substring(0, 5) + edge.suffix;
                return new RubberTypeResult { Makeo = edge.makeo, Ptype = edge.ptype, NormalizedTenkeo = normalized };
            }

            if (ThuNormMap.TryGetValue(suffix, out var thu))
            {
                string normalized = partno.Substring(0, 5) + thu.suffix;
                return new RubberTypeResult { Makeo = thu.makeo, Ptype = thu.ptype, NormalizedTenkeo = normalized };
            }

            return new RubberTypeResult { Makeo = "", Ptype = "", NormalizedTenkeo = tenkeo };
        }
    }
}
