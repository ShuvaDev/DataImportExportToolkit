using CsvHelper;
using DataImportExportToolkit.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Rotativa.AspNetCore;
using System.Globalization;

namespace DataImportExportToolkit.Controllers
{
    public class ExportToolkit : Controller
    {
        public IActionResult GeneratePDF()
        {
            Person person = new Person() { Name = "Shuva Dev", Age = 22 };

            return new ViewAsPdf("GeneratePDF", person, ViewData)
            {
                PageMargins = new Rotativa.AspNetCore.Options.Margins() { Top = 20, Bottom = 20, Left = 20, Right = 20 },
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Portrait
            };
        }

        public async Task<IActionResult> GenerateCSV()
        {
            List<Person> persons = new()
            {
                new() {Name = "Shuva", Age = 23},
                new() {Name = "Shanto", Age = 20}
            };


            MemoryStream memoryStream = new MemoryStream();
            StreamWriter streamWriter = new StreamWriter(memoryStream);

            CsvWriter csvWriter = new CsvWriter(streamWriter, CultureInfo.InvariantCulture, leaveOpen: true);

            csvWriter.WriteHeader<Person>();
            csvWriter.NextRecord();

            await csvWriter.WriteRecordsAsync(persons);
            //csvWriter.WriteField(persons[0].Name);
            await streamWriter.FlushAsync();

            memoryStream.Position = 0;

            return File(memoryStream, "text/csv", "test.csv");
        }

        public async Task<IActionResult> GenerateExcel()
        {
            MemoryStream memoryStream = new MemoryStream();
            using(ExcelPackage excelPackage = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells["A1"].Value = "Person Name";
                workSheet.Cells[2, 1].Value = "Shuva Dev";
                workSheet.Cells["B1"].Value = "Age";
                workSheet.Cells[2, 2].Value = "23";

                workSheet.Cells["A1:B2"].AutoFitColumns();

                using(ExcelRange headerCells = workSheet.Cells["A1:B1"])
                {
                    headerCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    headerCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    headerCells.Style.Font.Bold = true;
                }

                await excelPackage.SaveAsAsync(memoryStream);

            }
            memoryStream.Position = 0;
            return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test.xlsx");
        }

        public async Task UploadDataFromExcelFile(IFormFile formFile)
        {
            MemoryStream memoryStream = new MemoryStream();
            await formFile.CopyToAsync(memoryStream);

            using (ExcelPackage excelPackage = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets["Countries"];

                int rowCount = workSheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string? cellValue = workSheet.Cells[row, 1].Value.ToString();

                    if(!string.IsNullOrEmpty(cellValue))
                    {
                        Console.WriteLine(cellValue);
                        // add cellValue to db
                    }
                }
            }
        }

    }
}
