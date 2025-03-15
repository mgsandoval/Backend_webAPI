using Microsoft.AspNetCore.Mvc;
using Backend_webAPI_ASP.NETCore.Models;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Microsoft.EntityFrameworkCore;
using Backend_webAPI_ASP.NETCore.Models.ViewModels;

namespace Backend_webAPI_ASP.NETCore.Controllers
{
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

        [HttpPost]
        public IActionResult MostrarDatos([FromForm] IFormFile ArchivoExcel)
        {
            Stream stream = ArchivoExcel.OpenReadStream();

            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0);

            int cantidadFilas = HojaExcel.LastRowNum;

            List<VMContacto> lista = new List<VMContacto>();

            for (int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);
                lista.Add(new VMContacto
                {
                    Codigo                  = fila.GetCell(0) .ToString(),
                    Nombre_razon_social     = fila.GetCell(1) .ToString(),
                    Tipo_cliente            = fila.GetCell(2) .ToString(),
                    Moneda                  = fila.GetCell(3) .ToString(),
                    Telefono1               = fila.GetCell(4) .ToString(),
                    Telefono_movil          = fila.GetCell(5) .ToString(),
                    Correo_electronico      = fila.GetCell(6) .ToString(),
                    RTN                     = fila.GetCell(7) .ToString(),
                    Direccion               = fila.GetCell(8) .ToString(),
                    Vendedor                = fila.GetCell(9) .ToString(),
                    Territorio              = fila.GetCell(10).ToString(),
                    Nombre_completo         = fila.GetCell(11).ToString(),
                    Nombre                  = fila.GetCell(12).ToString(),
                    Apellido                = fila.GetCell(13).ToString(),
                    Telefono_fijo           = fila.GetCell(14).ToString(),
                    Movil_personal          = fila.GetCell(15).ToString(),
                    Correo_electronico2     = fila.GetCell(16).ToString(),
                    Destino                 = fila.GetCell(17).ToString(),
                    Id_direccion            = fila.GetCell(18).ToString(),
                    Nombre_direccion2       = fila.GetCell(19).ToString(),
                    Nombre_direccion3       = fila.GetCell(20).ToString(),
                    Ciudad                  = fila.GetCell(21).ToString(),
                    Condado                 = fila.GetCell(22).ToString(),
                    Condiciones_pago        = fila.GetCell(23).ToString(),
                    Lista_precios           = fila.GetCell(24).ToString(),
                    Limite_credito          = fila.GetCell(25).ToString(),
                    Cuenta_mayor_sucursal   = fila.GetCell(26).ToString()
                });
            }

            return StatusCode(StatusCodes.Status200OK, lista);
        }

        //[HttpPost]
        //public IActionResult EnviarDatos([FromForm] IFormFile ArchivoExcel)
        //{
        //    Stream stream = ArchivoExcel.OpenReadStream();

        //    IWorkbook MiExcel = null;

        //    if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
        //    {
        //        MiExcel = new XSSFWorkbook(stream);
        //    }
        //    else
        //    {
        //        MiExcel = new HSSFWorkbook(stream);
        //    }

        //    ISheet HojaExcel = MiExcel.GetSheetAt(0);

        //    int cantidadFilas = HojaExcel.LastRowNum;

        //    List<Contacto> lista = new List<Contacto>();

        //    for (int i = 1; i <= cantidadFilas; i++)
        //    {

        //        IRow fila = HojaExcel.GetRow(i);

        //        lista.Add(new Contacto
        //        {
        //            Nombre = fila.GetCell(0).ToString(),
        //            Apellido = fila.GetCell(1).ToString(),
        //            Telefono = fila.GetCell(2).ToString(),
        //            Correo = fila.GetCell(3).ToString(),

        //        });
        //    }

        //    _dbocontext.BulkInsert(lista);

        //    return StatusCode(StatusCodes.Status200OK, new { mensaje = "ok" });
        //}


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}