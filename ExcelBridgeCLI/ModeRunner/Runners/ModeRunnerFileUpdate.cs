using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeCli.Argument;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public class ModeRunnerFileUpdate : ModeRunnerAbstract, IModeRunner
    {
        public ModeRunnerFileUpdate(Options options) : base(options)
        {
            
        }

        public ModeRunnerResponse Run()
        {
            Console.WriteLine("File update");
            return new ModeRunnerResponse();
        }

        public void ValidateOptions()
        {
            throw new NotImplementedException();
        }
    }
}
