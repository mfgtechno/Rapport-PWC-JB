using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PWC_to_Jobboss
{
    public class Util
    {
        public static string ToString(object o)
        {
            if (o is DBNull) return string.Empty;
            if (o == null) return string.Empty;

            return o.ToString();
        }

        public static DateTime ToDateTime(object o)
        {
            string s = ToString(o);
            if (string.IsNullOrWhiteSpace(s) || !DateTime.TryParse(s, out DateTime d)) return DateTime.Parse("2000-01-01");
            return d;
        }

        public static DateTime ToDateTimeExcel(object o)
        {
            string s = ToString(o);

            if (int.TryParse(s, out int i)) return DateTime.FromOADate(i);

            if (string.IsNullOrWhiteSpace(s) || !DateTime.TryParse(s, out DateTime d)) return DateTime.Parse("2000-01-01");
            return d;
        }

        public static DateTime? ToDateTimeN(object o)
        {
            string s = ToString(o);
            if (string.IsNullOrWhiteSpace(s) || !DateTime.TryParse(s, out DateTime d)) return null;
            return d;
        }

        public static decimal ToDecimal(object o)
        {
            string s = ToString(o);
            if (string.IsNullOrWhiteSpace(s) || !decimal.TryParse(s, out decimal d)) return 0;
            return d;
        }

        public static long ToLong(object o)
        {
            string s = ToString(o);
            if (string.IsNullOrWhiteSpace(s) || !long.TryParse(s, out long d)) return 0;
            return d;
        }
    }
}
