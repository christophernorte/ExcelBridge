using ExcelBridgeApi.Writer;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBridgeTestUnit
{
    [TestClass]
    public class BenchmarkTest
    {

        [TestInitialize]
        public void Initialize()
        {
            Console.WriteLine("TestClass1.Initialize()");
        }

        [TestMethod]
        public void Debit_WhenAmountIsMoreThanBalance_ShouldThrowArgumentOutOfRange()
        {
            WriterProcessor writer = new WriterProcessor();

            for (uint i = 0; i < 100; i++)
            {
                writer.UpdateCell("FichierInOrig.xlsx", "sheet2", "insertByDotnet", i, "A");
                Console.WriteLine("Current line : " + i);
            }
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Console.WriteLine("TestClass1.TestCleanup()");
        }
    }

}
