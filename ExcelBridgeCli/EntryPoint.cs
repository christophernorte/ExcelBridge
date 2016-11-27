using ExcelBridgeApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeApi.Writer;
using ExcelBridgeCli.Argument;
using ExcelBridgeCli.ModeRunner;

namespace ExcelBridgeCli
{
    class EntryPoint
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                IModeRunner modeRunner = ModeRunnerFactory.getRunner(options.Mode,options);
                modeRunner.Run();
            }
            /*
            WriterProcessor writer = new WriterProcessor();

            for (uint i = 1; i < 100; i++)
            {
                writer.UpdateCell("FichierInOrig.xlsx","sheet2", "insertByDotnet", i, "A");
                Console.WriteLine("Current line : " + i);
            }*/
            //string line = Console.ReadLine();
        }
    }
}
