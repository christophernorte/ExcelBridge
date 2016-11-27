using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;
using ExcelBridgeCli.Exceptions;
using System.IO;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerCellUpdate : ModeRunnerAbstract,IModeRunner
    {
        public ModeRunnerCellUpdate(Options options) : base(options)
        {
            
        }

        public ModeRunnerResponse Run()
        {
            writer.UpdateCell(options.ExcelFilePath, options.SheetName, options.Value, 1, "A");
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
            bool hasString = false;
            bool hasNum = false;

            StringReader sr = new StringReader(cell);
            foreach (char c in cell)
            {
                if (Char.IsDigit(c) || Char.IsNumber(c))
                {
                    hasNum = true;
                }

                if (Char.IsLetter(c))
                {
                    hasString = true;
                }
            }

            if (!hasNum || !hasString)
            {
                throw new CliArgumentCellBadFormat("Cell argument must be an excel identifier with a letter and a numeric [" + cell+ "] given");
            }
        }
    }
}
