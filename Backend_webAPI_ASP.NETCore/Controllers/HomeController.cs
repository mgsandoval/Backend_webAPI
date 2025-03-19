using Backend_webAPI_ASP.NETCore.Models;
using Backend_webAPI_ASP.NETCore.Models.ViewModels;
using Backend_webAPI_ASP.NETCore.Controllers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Newtonsoft.Json;
using SAPbobsCOM;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.Encodings.Web;
using System.Globalization;


namespace Backend_webAPI_ASP.NETCore.Controllers
{
    // Clase HomeController que hereda de Controller, es el controlador principal de la aplicación web
    public class HomeController : Controller
    {
        private SAPbobsCOM.Company myCompany;
        private SAPbobsCOM.BusinessPartners mySN;
        private int err;
        private string errMsg;

        // Constructor de la clase HomeController
        private readonly ILogger<HomeController> _logger;

        // Constructor de la clase HomeController
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }


        /* Método para mostrar los datos del archivo Excel, recibe un archivo de tipo IFormFile y 
         * devuelve una lista de tipo VMContacto 
        */
        [HttpPost]
        public async Task<IActionResult> MostrarDatos([FromForm] IFormFile ArchivoExcel)
        {
            // Se lee el archivo Excel y se guarda en un stream
            Stream stream = ArchivoExcel.OpenReadStream();
            // Se crea una instancia de la clase IWorkbook de NPOI para leer el archivo Excel
            IWorkbook MiExcel = null;

            /* Se verifica si el archivo es de extensión .xlsx o .xls: 
             * si es .xlsx se crea una instancia de XSSFWorkbook, 
             * si es .xls se crea una instancia de HSSFWorkbook 
            */
            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }
            // Se obtiene la primera hoja del archivo Excel
            ISheet HojaExcel = MiExcel.GetSheetAt(0);
            // Se obtiene la cantidad de filas de la hoja Excel
            int cantidadFilas = HojaExcel.LastRowNum;

            // Se crea una lista de tipo VMContacto, donde se almacenan los datos del archivo Excel
            List<VMContacto> lista = new List<VMContacto>();

            // Se recorren las filas del archivo Excel y se guardan en la lista
            for (int i = 1; i <= cantidadFilas; i++)
            {
                // Se obtiene la fila actual
                IRow fila = HojaExcel.GetRow(i);
                // Se agregan los datos de la fila a la lista
                lista.Add(new VMContacto
                {
                    // Se obtienen los datos de cada celda de la fila
                    Codigo                = fila.GetCell(0) .ToString(),
                    Nombre_razon_social   = fila.GetCell(1) .ToString(),
                    Tipo_cliente          = fila.GetCell(2) .ToString(),
                    Moneda                = fila.GetCell(3) .ToString(),
                    Telefono1             = fila.GetCell(4) .ToString(),
                    Telefono_movil        = fila.GetCell(5) .ToString(),
                    Correo_electronico    = fila.GetCell(6) .ToString(),
                    RTN                   = fila.GetCell(7) .ToString(),
                    Direccion             = fila.GetCell(8) .ToString(),
                    Vendedor              = fila.GetCell(9) .ToString(),
                    Territorio            = fila.GetCell(10).ToString(),
                    Nombre_completo       = fila.GetCell(11).ToString(),
                    Nombre                = fila.GetCell(12).ToString(),
                    Apellido              = fila.GetCell(13).ToString(),
                    Telefono_fijo         = fila.GetCell(14).ToString(),
                    Movil_personal        = fila.GetCell(15).ToString(),
                    Correo_electronico2   = fila.GetCell(16).ToString(),
                    Destino               = fila.GetCell(17).ToString(),
                    Id_direccion          = fila.GetCell(18).ToString(),
                    Nombre_direccion2     = fila.GetCell(19).ToString(),
                    Nombre_direccion3     = fila.GetCell(20).ToString(),
                    Ciudad                = fila.GetCell(21).ToString(),
                    Condado               = fila.GetCell(22).ToString(),
                    Condiciones_pago      = fila.GetCell(23).ToString(),
                    Lista_precios         = fila.GetCell(24).ToString(),
                    Limite_credito        = fila.GetCell(25).ToString(),
                    Cuenta_mayor_sucursal = fila.GetCell(26).ToString()
                });
            }
            // Se retorna la lista con los datos del archivo Excel
            return StatusCode(StatusCodes.Status200OK, lista);
        }

        /* Método para enviar los datos del archivo Excel a SAP B1, recibe un archivo de tipo IFormFile 
         * y devuelve un mensaje de tipo string, los mensajes se imprimen en la consola de la aplicación web
        */
        [HttpPost]
        public async Task<IActionResult> EnviarDatos([FromForm] List<Contacto> contactos)
        {
            //-------------- Código nuevo: --------------------------
            // Conexión a SAP B1
            try
            {
                Console.WriteLine("Intentando conectarse a SAP B1...");
                myCompany = new Company();
                myCompany.SLDServer = "192.168.100.241";
                myCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                myCompany.Server = "192.168.100.241";
                myCompany.CompanyDB = "PRUEBAS";
                myCompany.UserName = "dev";
                myCompany.Password = "Admin123";
                myCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La;

                err = myCompany.Connect();

                if (err != 0)
                {
                    Console.WriteLine("Error de conexión: " + myCompany.GetLastErrorDescription());
                    return StatusCode(StatusCodes.Status500InternalServerError, new { mensaje = "Error de conexión a SAP B1" });
                }
                else
                {
                    Console.WriteLine("Conectado a SAP B1 exitosamente");
                    foreach (var contacto in contactos)
                    {
                        mySN = (SAPbobsCOM.BusinessPartners)myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                        mySN.CardCode = contacto.Codigo;
                        mySN.CardName = contacto.Nombre_razon_social;
                        mySN.AdditionalID = contacto.Tipo_cliente;
                        mySN.FederalTaxID = contacto.RTN;
                        mySN.Phone1 = contacto.Telefono1;

                        err = mySN.Add();
                        if (err != 0)
                        {
                            Console.WriteLine("Error al crear el cliente: " + myCompany.GetLastErrorDescription(), "Error");
                        }
                        else
                        {
                            Console.WriteLine("Cliente creado exitosamente");
                        }
                    }

                    myCompany.Disconnect();
                    Console.WriteLine("Desconectado de SAP B1");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error general al conectar a SAP B1: " + ex.Message);
                return StatusCode(StatusCodes.Status500InternalServerError, new { message = "Error general al conectar a SAP B1" });
            }
            return StatusCode(StatusCodes.Status200OK, new { mensaje = "Datos enviados a SAP B1 exitosamente" });

        }

        //-------------- Código anterior: -----------------------
        // Se lee el archivo Excel y se guarda en un stream
        /*Stream stream = ArchivoExcel.OpenReadStream(); // Esta línea está provocando 1 error/excepción

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

        List<Contacto> lista = new List<Contacto>();

        for (int i = 1; i <= cantidadFilas; i++)
        {
            IRow fila = HojaExcel.GetRow(i);
            lista.Add(new Contacto
            {
                Codigo = fila.GetCell(0).ToString(),
                Nombre_razon_social = fila.GetCell(1).ToString(),
                Tipo_cliente = fila.GetCell(2).ToString(),
                Moneda = fila.GetCell(3).ToString(),
                Telefono1 = fila.GetCell(4).ToString(),
                Telefono_movil = fila.GetCell(5).ToString(),
                Correo_electronico = fila.GetCell(6).ToString(),
                RTN = fila.GetCell(7).ToString(),
                Direccion = fila.GetCell(8).ToString(),
                Vendedor = fila.GetCell(9).ToString(),
                Territorio = fila.GetCell(10).ToString(),
                Nombre_completo = fila.GetCell(11).ToString(),
                Nombre = fila.GetCell(12).ToString(),
                Apellido = fila.GetCell(13).ToString(),
                Telefono_fijo = fila.GetCell(14).ToString(),
                Movil_personal = fila.GetCell(15).ToString(),
                Correo_electronico2 = fila.GetCell(16).ToString(),
                Destino = fila.GetCell(17).ToString(),
                Id_direccion = fila.GetCell(18).ToString(),
                Nombre_direccion2 = fila.GetCell(19).ToString(),
                Nombre_direccion3 = fila.GetCell(20).ToString(),
                Ciudad = fila.GetCell(21).ToString(),
                Condado = fila.GetCell(22).ToString(),
                Condiciones_pago = fila.GetCell(23).ToString(),
                Lista_precios = fila.GetCell(24).ToString(),
                Limite_credito = fila.GetCell(25).ToString(),
                Cuenta_mayor_sucursal = fila.GetCell(26).ToString()
            });
        }

        HttpContext.Session.SetString("lista", JsonConvert.SerializeObject(lista));

        // Conexión a SAP B1
        try
        {
            Console.WriteLine("Intentando conectarse a SAP B1...");
            Company myCompany = new Company();
            myCompany.SLDServer = "192.168.100.241"; // IP o nombre de servidor
            myCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012; // Tipo de servidor de base de datos (MSSQL, 2012, 2014, 2016, HANA)
            myCompany.Server = "192.168.100.241"; // Nombre de la instancia de SQL Server
            myCompany.CompanyDB = "PRUEBAS"; // Nombre de la base de datos
            myCompany.UserName = "dev"; // Usuario de SAP B1
            myCompany.Password = "Admin123"; // Contraseña de SAP B1
            myCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La; // Lenguaje de la base de datos, en este caso español de Latinoamérica


            err = myCompany.Connect();

            if (err != 0)
            {
                Console.WriteLine("Error de conexión: " + myCompany.GetLastErrorDescription());
            }
            else
            {
                Console.WriteLine("Conectado a SAP B1 exitosamente");
                myCompany.Disconnect();
                Console.WriteLine("Desconectado de SAP B1");


                //foreach (var contacto in lista)
                //{
                //    mySN = (SAPbobsCOM.BusinessPartners)myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners); // Se crea una instancia de la clase BusinessPartners de la DI API
                //    mySN = (SAPbobsCOM.BusinessPartners)myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners); // Se crea una instancia de la clase BusinessPartners de la DI API
                //    mySN.CardCode = contacto.Codigo; // Código del cliente
                //    mySN.CardName = contacto.Nombre_razon_social; // Nombre del cliente
                //    mySN.AdditionalID = contacto.Tipo_cliente; // Tipo de identificación adicional
                //    mySN.FederalTaxID = contacto.RTN; // Identificación fiscal
                //    mySN.Phone1 = contacto.Telefono1; // Teléfono

                // Se intenta agregar el cliente a la base de datos
                //    err = mySN.Add();
                //    if (err != 0)
                //    {
                //        Console.WriteLine("Error al crear el cliente: " + myCompany.GetLastErrorDescription(), "Error");
                //    }
                //    else
                //    {
                //        Console.WriteLine("Cliente creado exitosamente");
                //    }
                //}
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error general al conectar a SAP B1: " + ex.Message);
        }

        return StatusCode(StatusCodes.Status200OK, new { mensaje = "ok" });
    }*/

        public IActionResult Privacy()
        {
            return View();
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = System.Diagnostics.Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}