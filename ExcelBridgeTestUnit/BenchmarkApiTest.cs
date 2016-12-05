using ExcelBridgeApi.Writer;
using ExcelBridgeCore.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelBridgeTestUnit
{
    [TestClass]
    public class BenchmarkApiTest
    {

        [TestInitialize]
        public void Initialize()
        {
            Console.WriteLine("TestClass1.Initialize()");
        }

        [TestMethod]
        public void Cell_InsertLargeAmountOfData()
        {
            //WriterProcessor writer = new WriterProcessor();

            //for (uint i = 0; i < 100; i++)
            //{
            //    writer.UpdateCell("FichierInOrig.xlsx", "sheet2", "insertByDotnet", i, "A");
            //    Console.WriteLine("Current line : " + i);
            //}
        }

        [TestMethod]
        public void Range_InsertLargeAmountOfData_SameColumn_10000Cells()
        {
            WriterProcessor writer = new WriterProcessor();

            Range range = new Range("A");
            for (uint i = 1; i < 10000; i++)
            {
                range.addCellCore(new CellCore(i, "insert InsertLargeAmountOfData : " + i.ToString()));
            }

            writer.UpdateRange("FichierInOrig.xlsx", "Sheet2", range);
        }

        [TestMethod]
        public void Range_InsertLargeAmountOfData_MultiColumn_50000Cells()
        {
            WriterProcessor writer = new WriterProcessor();

            List<string> listColumn = new List<string>();

            listColumn.Add("A");
            // gap
            listColumn.Add("C");
            listColumn.Add("D");
            // More gap
            listColumn.Add("G");
            listColumn.Add("H");

            List<Range> listRange = new List<Range>();
           
            foreach (string column in listColumn)
            {
                Range range = new Range(column);
                for (uint i = 1; i < 10000; i++)
                {
                    range.addCellCore(new CellCore(i, "insert InsertMultiColumn ["+ column + "] line [" + i.ToString()+"]"));
                }
                listRange.Add(range);
            }

            writer.UpdateListOfRange("FichierInOrig.xlsx", "Sheet2", listRange);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Console.WriteLine("TestClass1.TestCleanup()");
        }
    }

}
