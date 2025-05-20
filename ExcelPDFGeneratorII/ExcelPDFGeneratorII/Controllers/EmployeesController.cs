using ClosedXML.Excel;
using ExcelPDFGeneratorII.Interfaces;
using Microsoft.AspNetCore.Mvc;
using PuppeteerSharp;
using System.Text.Json;

namespace ExcelPDFGeneratorII.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmployeesController : ControllerBase
    {
        private readonly IEmployeeRepository _employeeRepository;

        public EmployeesController(IEmployeeRepository employeeRepository)
        {
            _employeeRepository = employeeRepository;
        }

        [HttpGet("excel-report")]
        public async Task<IActionResult> GetExcelReport()
        {
            var employees = await _employeeRepository.GetAllEmployeesAsync();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Employees");

            worksheet.Cell(1, 1).Value = "Id";
            worksheet.Cell(1, 2).Value = "First Name";
            worksheet.Cell(1, 3).Value = "Last Name";
            worksheet.Cell(1, 4).Value = "Salary";
            worksheet.Cell(1, 5).Value = "Department";

            int row = 2;
            foreach (var emp in employees)
            {
                worksheet.Cell(row, 1).Value = emp.EmployeeId;
                worksheet.Cell(row, 2).Value = emp.FirstName;
                worksheet.Cell(row, 3).Value = emp.LastName;
                worksheet.Cell(row, 4).Value = emp.Salary;
                worksheet.Cell(row, 5).Value = emp.Department?.Name;
                row++;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Seek(0, SeekOrigin.Begin);

            return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "EmployeesReport.xlsx");
        }

        [HttpGet("pdf-report")]
        public async Task<IActionResult> GetPdfReport()
        {
            var employees = await _employeeRepository.GetAllEmployeesAsync();

            var options = new LaunchOptions { Headless = true };
            await new BrowserFetcher().DownloadAsync();

            using var browser = await Puppeteer.LaunchAsync(options);
            using var page = await browser.NewPageAsync();

            var url = Url.ActionLink("Index", "PDFGenerator", new
            {
                model = JsonSerializer.Serialize(employees)
            });

            await page.GoToAsync(url);
            var pdf = await page.PdfDataAsync();

            return File(pdf, "application/pdf", "EmployeesReport.pdf");
        }
    }
}
