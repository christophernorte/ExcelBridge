using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;
using ExcelBridgeCli.Exceptions;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerCellUpdate : ModeRunnerAbstract,IModeRunner
    {
        public ModeRunnerCellUpdate(Options options) : base(options)
        {
            
        }

        public ModeRunnerResponse Run()
        {
            string cellLetterPart = getCellStringPart(options.Cell);
            uint cellNumericPart = getCellNumericPart(options.Cell);

            writer.UpdateCell(options.ExcelFilePath, options.SheetName, options.Value, cellNumericPart, cellLetterPart);
            return new ModeRunnerResponse();
        }

        private string getCellStringPart(string rawCell)
        {
            Regex regex = new Regex(@"[a-zA-Z]+");
            Match match = regex.Match(rawCell);
            if(match.Success)
            {
                return match.Value;
            }
            else
            {
                return "";
            }
        }

        private uint getCellNumericPart(string rawCell)
        {
            Regex regex = new Regex(@"[\d]");
            Match match = regex.Match(rawCell);
            if (match.Success)
            {
                uint value = 0;
                if(uint.TryParse(match.Value,out value))
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

        public void ValidateOptions()
        {
            if(options.Value == null)
            {
                throw new CliArgumentMissing("Value argument is mandatory in cell mode");
            }

            if (options.SheetName == null)
            {
                throw new CliArgumentMissing("SheetName argument is mandatory in cell mode");
            }

            if (options.Cell == null)
            {
                throw new CliArgumentMissing("Cell argument is mandatory in cell mode");
            }

            this.ValidateCell(options.Cell);
        }

        private void ValidateCell(string cell)
        {
            Regex regex = new Regex(@"^[a-zA-Z]+[\d]+$");

            if (!regex.IsMatch(cell))
            {
                throw new CliArgumentCellBadFormat("Cell argument must be an excel identifier with a letter and a numeric [" + cell+ "] given");
            }
        }
    }
}
