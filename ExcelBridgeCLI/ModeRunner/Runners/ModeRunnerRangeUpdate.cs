using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;
using ExcelBridgeCli.Exceptions;
using System.Text.RegularExpressions;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerRangeUpdate : ModeRunnerAbstract, IModeRunner
    {
        public ModeRunnerRangeUpdate(Options options) : base(options)
        {
        }

        public ModeRunnerResponse Run()
        {
            

            return new ModeRunnerResponse();
        }

        public void ValidateOptions()
        {
            if (options.Value == null)
            {
                throw new CliArgumentMissing("Value argument is mandatory in cell mode");
            }

            if (options.SheetName == null)
            {
                throw new CliArgumentMissing("SheetName argument is mandatory in cell mode");
            }

            //this.ValidateCell(options.Cell);
        }

        private void ValidateCell(string cell)
        {
            Regex regex = new Regex(@"^([a-zA-Z]+[\d]+:?){1,}$");

            if (!regex.IsMatch(cell))
            {
                throw new CliArgumentCellBadFormat("Cell argument must be an excel identifier with a letter and a numeric [" + cell + "] given");
            }
        }
    }
}
