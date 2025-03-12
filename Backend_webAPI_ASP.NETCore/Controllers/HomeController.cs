using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Backend_webAPI_ASP.NETCore.Models;

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
        if (file == null || file.Length <= 0)
        {
            var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\uploads";

            if (!Directory.Exists(uploadDirectory))
            {
                Directory.CreateDirectory(uploadDirectory);
            }
        }
        return View();
    }
}
