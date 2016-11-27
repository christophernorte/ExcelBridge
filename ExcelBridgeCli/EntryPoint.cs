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
            ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
            ExcelBridgeController.Start(args);
            //string line = Console.ReadLine();
        }

    }
}
