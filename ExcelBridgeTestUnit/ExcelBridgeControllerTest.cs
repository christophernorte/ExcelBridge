using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelBridgeCli;
using ExcelBridgeCli.Exceptions;

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
            ExcelBridgeController.Start(new string[] { "-m", "cell", "-e", "FichierInOrig.xlsx", "-s", "Sheet2", "-c", "A1", "-v", "InsertedFromTest" });
        }

        [TestMethod]
        public void cell_ExecuteInRangeMode_ShouldEndWithExcelFileCellUpdated()
        {
            ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
            ExcelBridgeController.Start(new string[] { "-m", "range", "-e", "FichierInOrig.xlsx", "-s", "Sheet2", "-v", "InsertedFromTest" });
        }


        [TestMethod]
        public void cell_InjectBadArgument_ShouldStopProcess()
        {
            try
            {
                ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
                ExcelBridgeController.Start(new string[] { "-m", "cell", "-e", "FichierInOrig.xlsx", "-s", "Sheet2", "-c", "A1" });
                Assert.Fail("An exception should have been thrown");
            }
            catch (CliArgumentMissing ae)
            {
                Assert.AreEqual("Value argument is mandatory in cell mode", ae.Message);
            }
            catch (Exception e)
            {
                Assert.Fail(string.Format("Unexpected exception of type {0} caught: {1}", e.GetType(), e.Message));
            }
        }

        [TestMethod]
        public void cell_InjectBadCellArgument_ShouldStopProcess()
        {
            try
            {
                ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
                ExcelBridgeController.Start(new string[] { "-m", "cell", "-e", "FichierInOrig.xlsx", "-s", "Sheet2", "-c", "AAA", "-v", "InsertedFromTest" });
                Assert.Fail("An exception should have been thrown");
            }
            catch (CliArgumentCellBadFormat ae)
            {
                Assert.AreEqual("Cell argument must be an excel identifier with a letter and a numeric [AAA] given", ae.Message);
            }
            catch (Exception e)
            {
                Assert.Fail(string.Format("Unexpected exception of type {0} caught: {1}", e.GetType(), e.Message));
            }
        }

        [TestMethod]
        public void cell_InjectGoodCellArgument_ShouldPassArgumentValidationProcess()
        {
            try
            {
                ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
                ExcelBridgeController.Start(new string[] { "-m", "cell", "-e", "FichierInOrig.xlsx", "-s", "Sheet2", "-c", "A2", "-v", "InsertedFromTest" });
            }
            catch (Exception e)
            {
                Assert.Fail(string.Format("Unexpected exception of type {0} caught: {1}", e.GetType(), e.Message));
            }
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Console.WriteLine("TestClass1.TestCleanup()");
        }
    }
}
