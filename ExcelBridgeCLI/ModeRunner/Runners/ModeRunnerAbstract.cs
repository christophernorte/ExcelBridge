using ExcelBridgeCli.Argument;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public abstract class ModeRunnerAbstract
    {
        protected Options options;

        public ModeRunnerAbstract(Options options)
        {
            this.options = options;
        }
    }
}
