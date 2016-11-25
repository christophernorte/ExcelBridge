using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestUnitExcelWriter
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
        { // 
            StringAssert.Contains("Test", "Test");

        }

        [TestCleanup]
        public void TestCleanup()
        {
            Console.WriteLine("TestClass1.TestCleanup()");
        }
    }

}