using System;
using System.Threading;

namespace ExcelBridgeCli
{
    class Program
    {
        static void Main(string[] args)
        {
            WriterService writer = new WriterService();

            for(uint i = 0;i < 100;i++)
            {
                writer.UpdateCell("TestGraphe.xlsx", "insertByDotnet", i, "A");
                Console.WriteLine("Current line : "+i);
            }

            Thread.Sleep(2000);
        }
        
    }
}
