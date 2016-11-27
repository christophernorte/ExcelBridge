using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerRangeUpdate : ModeRunnerAbstract, IModeRunner
    {
        public ModeRunnerRangeUpdate(Options options) : base(options)
        {
        }

        public ModeRunnerResponse Run()
        {
            Console.WriteLine("Range update");
            return new ModeRunnerResponse();
        }

        public void ValidateOptions()
        {
            throw new NotImplementedException();
        }
    }
}
