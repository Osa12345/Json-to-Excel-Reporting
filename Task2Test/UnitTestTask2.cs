using GemBox.Spreadsheet;
using NUnit.Framework;
using System.Collections.Generic;


namespace Task2Test
{
    public class UnitTestTask2
    {
        [Test]
        public void Test1()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook = ExcelFile.Load("Hotel Rates.xlsx");

            List<string> headerNames = new List<string>()
            {
                "ARRIVAL_DATE",
                "DEPARTURE_DATE",
                "PRICE",
                "CURRENCY",
                "RATENAME",
                "ADULTS",
                "BREAKFAST_INCLUDED"
            };

            bool isMatch = false;
            foreach (var header in workbook.Worksheets[0].Rows[0].AllocatedCells)
            {
                isMatch = headerNames.Contains(header.StringValue) ? true : false;
                if (!isMatch)
                {
                    break;
                }
            }

            Assert.AreEqual(isMatch, true);
        }
    }
}