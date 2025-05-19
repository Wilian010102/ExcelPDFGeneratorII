using ExcelPDFGeneratorII.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace ExcelPDFGeneratorII.Controllers
{
    public class PDFGeneratorController : Controller
    {
        public IActionResult Index(string model)
        {
            var employees = JsonSerializer.Deserialize<IEnumerable<Employee>>(model);
            return View(employees);
        }
    }
}
