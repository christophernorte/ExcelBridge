using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelBridgeApi.Tools
{
    class CellHelper
    {
        public static string getCellStringPart(string rawCell)
        {
            Regex regex = new Regex(@"[a-zA-Z]+");
            Match match = regex.Match(rawCell);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "";
            }
        }

        public static uint getCellNumericPart(string rawCell)
        {
            Regex regex = new Regex(@"[\d]");
            Match match = regex.Match(rawCell);
            if (match.Success)
            {
                uint value = 0;
                if (uint.TryParse(match.Value, out value))
                {
                    return value;
                }
                return value;
            }
            else
            {
                return 0;
            }
        }
    }
}
