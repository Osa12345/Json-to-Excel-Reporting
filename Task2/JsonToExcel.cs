using GemBox.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.IO;

namespace Task2
{
    public class JsonToExcel
    {
        public JsonData LoadJson()
        {
            string fileName = "task 2 - hotelrates.json";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);

            using (StreamReader jsonStream = new StreamReader(path))
            {
                string json = jsonStream.ReadToEnd();
                return JsonConvert.DeserializeObject<JsonData>(json);
            }
        }

        public string CreateExcel(JsonData items)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile workbook = new ExcelFile();
            ExcelWorksheet worksheet = workbook.Worksheets.Add("Hotel Info");

            var cells = worksheet.Cells;
            var rows = worksheet.Rows;
            var columns = worksheet.Columns;

            rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
            rows[0].Style.Font.Color = SpreadsheetColor.FromName(ColorName.LightBlue);

            columns[0].SetWidth(5, LengthUnit.Centimeter);
            columns[1].SetWidth(5, LengthUnit.Centimeter);
            columns[2].SetWidth(5, LengthUnit.Centimeter);
            columns[3].SetWidth(5, LengthUnit.Centimeter);
            columns[4].SetWidth(5, LengthUnit.Centimeter);
            columns[5].SetWidth(5, LengthUnit.Centimeter);
            columns[6].SetWidth(5, LengthUnit.Centimeter);


            // Define header values
            cells[0, 0].Value = "ARRIVAL_DATE";
            cells[0, 1].Value = "DEPARTURE_DATE";
            cells[0, 2].Value = "PRICE";
            cells[0, 3].Value = "CURRENCY";
            cells[0, 4].Value = "RATENAME";
            cells[0, 5].Value = "ADULTS";
            cells[0, 6].Value = "BREAKFAST_INCLUDED";


            int row = 0;
            foreach (HotelRate val in items.hotelRates)
            {
                cells[++row, 0].Value = String.Format("{0:dd.MM.yy}", val.targetDay);
                cells[row, 1].Value = String.Format("{0:dd.MM.yy}", val.targetDay.AddDays(val.los)); ;
                cells[row, 2].Value = val.price.numericFloat;
                cells[row, 3].Value = val.price.currency;
                cells[row, 4].Value = val.rateName;
                cells[row, 5].Value = val.adults;
                if (val.rateTags[0].name == "breakfast")
                {
                    cells[row, 6].Value = val.rateTags[0].shape ? 1 : 0;
                }

                if (row % 2 != 0)
                {
                    rows[row].Style.FillPattern.SetSolid(SpreadsheetColor.FromName(ColorName.LightBlue));
                }

            }
            var filterRange = cells.GetSubrangeAbsolute(0, 0, items.hotelRates.Count, 6);
            filterRange.Filter().Apply();

            // Save excel file
            string fileName = string.Format("Hotel Rates{0}.xlsx", DateTime.Now.ToString("MM-dd-yyyy-HH-mm"));
            workbook.Save(fileName);
            return Path.Combine(Environment.CurrentDirectory) + "/" + fileName;

        }
    }
}
