using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text;
using System.Threading.Tasks;
using ExcelWriter.Writer;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            WriterService writer = new WriterService();

            for(uint i = 0;i < 2000;i++)
            {
                writer.UpdateCell("TestGraphe.xlsx", "insertByDotnet", i, "A");
                Console.WriteLine("Current line : "+i);
            }
            
            //writer.UpdateCell("FichierIn.xlsx", "80", 3, "B");
            //writer.UpdateCell("FichierIn.xlsx", "80", 2, "C");
            //writer.UpdateCell("FichierIn.xlsx", "20", 3, "C");

            Thread.Sleep(2000);
        }
        
    }
}
