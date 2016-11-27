using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBridgeCli.Exceptions
{
    public class CliArgumentCellBadFormat : Exception
    {
        public CliArgumentCellBadFormat()
        {
        }

        public CliArgumentCellBadFormat(string message) : base(message)
        {
        }

        public CliArgumentCellBadFormat(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
