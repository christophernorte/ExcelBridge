using ExcelBridgeCli.Argument;
using ExcelBridgeCli.ModeRunner.Runners;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBridgeCli.ModeRunner
{
    public class ModeRunnerFactory
    {
        public static IModeRunner getRunner(string mode, Options options)
        {
            mode = mode.ToLower();
            switch (mode)
            {
                case "cell":
                    return new ModeRunnerCellUpdate(options);
                case "range":
                    return new ModeRunnerRangeUpdate(options);
                case "file":
                    return new ModeRunnerFileUpdate(options);
                default:
                    return new ModeRunnerCellUpdate(options);
            }
        }
    }
}
