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
            string cellLetterPart = this.getCellStringPart(options.Cell);
            uint cellNumericPart = this.getCellNumericPart(options.Cell);

            writer.UpdateCell(options.ExcelFilePath, options.SheetName, options.Value, cellNumericPart, cellLetterPart);
            return new ModeRunnerResponse();
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
