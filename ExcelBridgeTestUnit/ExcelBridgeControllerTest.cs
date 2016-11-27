using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelBridgeCli;

namespace ExcelBridgeTestUnit
{
    [TestClass]
    public class ExcelBridgeControllerTest
    { 

        [TestInitialize]
        public void Initialize()
        {
            Console.WriteLine("TestClass1.Initialize()");
        }

        [TestMethod]
        public void cell_InjectGoodArgument_ShouldEndWithExcelFileCellUpdated()
        {
            ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();

            ExcelBridgeController.Start(new string[] { "-m cell", "-e path" });
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Console.WriteLine("TestClass1.TestCleanup()");
        }
    }
}
