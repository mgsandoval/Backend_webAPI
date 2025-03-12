//using ExcelDataReader;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using System.Diagnostics;

//namespace Backend_webAPI.Controllers
//{
//    public class HomeController
//    {
//        private readonly ILogger<HomeController> _logger;

//        public HomeController(ILogger<HomeController> logger)
//        {
//            _logger = logger;
//        }

//        public IActionResult Index()
//        {
//            return View();
//        }

//        public IActionResult Privacy()
//        {
//            return View();
//        }

//        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]

//        public IActionResult Error()
//        {
//            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//        }

//        [HttpPost]
//        public async Task<IActionResult> ExcelFileReader(IFormFile file)
//        {
//            if (file == null || file.Length <= 0)
//            {
//                var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\uploads";

//                if (!Directory.Exists(uploadDirectory))
//                {
//                    Directory.CreateDirectory(uploadDirectory);
//                }
//            }
//        }

//    }
//}
