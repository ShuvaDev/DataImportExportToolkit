using DataImportExportToolkit.Models;
using Microsoft.AspNetCore.Mvc;
using Rotativa.AspNetCore;

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
    }
}
