﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelWriter.Writer
{
    class WriterService
    {
        public void UpdateCell(string docName, string text, uint rowIndex, string columnName)
        {
            try { 
                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
                {
                    WorksheetPart worksheetPart = this.GetWorksheetPartByName(spreadSheet, "Sheet2");

                    if (worksheetPart != null)
                    {
                        Cell cell = GetCell(worksheetPart.Worksheet, columnName, rowIndex);

                        cell.CellValue = new CellValue(text);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                        // Save the worksheet.
                        worksheetPart.Worksheet.Save();
                    }
                }
            }
            catch (ArgumentNullException e)
            {
                Console.Error.WriteLine(e.ToString());
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }

        }

        private WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart) document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private Cell GetCell(Worksheet worksheet,string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
        }


        // Given a worksheet and a row index, return the row.
        private Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }
}
