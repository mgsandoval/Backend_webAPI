using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Backend_webAPI_ASP.NETCore.Models;
using ExcelDataReader;

namespace Backend_webAPI_ASP.NETCore.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    public IActionResult ExcelFileReader()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> ExcelFileReader(IFormFile file)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Upload file to server
        // string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");
        string path = Directory.GetCurrentDirectory();
        Console.WriteLine(path);
        if (file != null && file.Length > 0)
        {
            var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\Uploads";
             
            if (!Directory.Exists(uploadDirectory))
            {
                Console.WriteLine("Directorio inexistente");
                Directory.CreateDirectory(uploadDirectory);
            }

            var filePath = Path.Combine(uploadDirectory, file.FileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Read file
            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var excelData = new List<List<object>>();
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            var rowData = new List<object>();
                            for (int column =0; column < reader.FieldCount; column++)
                            {
                                rowData.Add(reader.GetValue(column));
                            }
                            excelData.Add(rowData);
                        }
                    } while (reader.NextResult());
                    ViewBag.ExcelData = excelData;
                }
            }
        }
        return View();
    }
}
