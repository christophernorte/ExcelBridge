using System;

namespace ExcelBridgeCli.Exceptions
{
    public class CliArgumentMissing : Exception
    {
        public CliArgumentMissing()
        {
        }

        public CliArgumentMissing(string message) : base(message)
        {
        }

        public CliArgumentMissing(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
