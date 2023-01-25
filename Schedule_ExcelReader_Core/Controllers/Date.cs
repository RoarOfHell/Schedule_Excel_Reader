using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Schedule_ExcelReader_Core.Controllers
{
    public static class Date
    {
        public static DateTime ExtractFromString(string date)
        {
            StringBuilder result = new StringBuilder();

            List<Char> chars = new List<char>() { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.', ',', ':', '-' };
            date.ToList().ForEach(p =>
            {
                if (chars.Contains(p)) result.Append(p.ToString());
            });


            return DateTime.Parse(result.ToString());
        }
    }
}
