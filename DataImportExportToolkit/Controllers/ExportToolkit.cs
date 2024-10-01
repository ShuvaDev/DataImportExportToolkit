using CsvHelper;
using DataImportExportToolkit.Models;
using Microsoft.AspNetCore.Mvc;
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


    }
}
