using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerCellUpdate : ModeRunnerAbstract,IModeRunner
    {
        public ModeRunnerCellUpdate(Options options) : base(options)
        {
            
        }

        public ModeRunnerResponse Run()
        {
            Console.WriteLine("Cell update");
            return new ModeRunnerResponse();
        }

        public void ValidateOptions()
        {
            
        }
    }
}
