using ExcelBridgeApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBridgeApi.Writer;

namespace ExcelBridgeCli
{
    class EntryPoint
    {
        static void Main(string[] args)
        {

            WriterProcessor writer = new WriterProcessor();

            for (uint i = 0; i < 100; i++)
            {
                writer.UpdateCell("FichierInOrig.xlsx", "insertByDotnet", i, "A");
                Console.WriteLine("Current line : " + i);
            }
            string line = Console.ReadLine();
        }
    }
}
