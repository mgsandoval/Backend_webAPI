using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Backend_webAPI_ASP.NETCore
{
    public partial class CreacionClientes
    {
        private SAPbobsCOM.Company myCompany; // Objeto de la clase Company de la DI API
        private SAPbobsCOM.BusinessPartners mySN; // Objeto de la clase BusinessPartners de la DI API
        private int error; // Variable para almacenar el código de error
        private string errorMessage; // Variable para almacenar el mensaje de error

        public void CrearCliente()
        {
            // Se intenta conectar a SAP B1
            try
            {
                // Se muestra un mensaje indicando que se está intentando conectar a SAP B1
                Console.WriteLine("Intentando conectarse a SAP B1...");
                // Se crea una instancia de la clase Company de la DI API
                myCompany = new SAPbobsCOM.Company();
                // Se establecen los parámetros de conexión
                //myCompany.LicenseServer = "192.168.100.241:40000"; // IP o nombre de servidor
                myCompany.SLDServer = "192.168.100.241"; // IP o nombre de servidor

                myCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012; // Tipo de servidor de base de datos (MSSQL, 2012, 2014, 2016, HANA)
                myCompany.Server = "192.168.100.241"; // Nombre de la instancia de SQL Server
                myCompany.CompanyDB = "PRUEBAS"; // Nombre de la base de datos
                myCompany.UserName = "dev"; // Usuario de SAP B1
                myCompany.Password = "Admin123"; // Contraseña de SAP B1
                myCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La; // Lenguaje de la base de datos, en este caso español de Latinoamérica

                // Se intenta conectar a SAP B1
                // error = 0 indica que la conexión fue exitosa
                error = myCompany.Connect();

                // Si el error es diferente de 0, entonces hubo un error al conectar, en caso contrario, la conexión fue exitosa.
                if (error != 0) // Conexión fallida
                {
                    Console.WriteLine("Error al conectar a SAP B1: " + myCompany.GetLastErrorDescription(), "Error");
                }
                else // Conexión exitosa
                {
                    Console.WriteLine("Conectado a SAP B1 exitosamente");
                    // aquí deben ir las operaciones desde DI
                    // por ejemplo, creación de clientes, facturas...
                    mySN = (SAPbobsCOM.BusinessPartners)myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners); // Se crea una instancia de la clase BusinessPartners de la DI API
                    mySN.CardCode = "C0001"; // Código del cliente
                    mySN.CardName = "Cliente creado desde DI API 01"; // Nombre del cliente
                    mySN.AdditionalID = "CF"; // Tipo de identificación adicional
                    mySN.FederalTaxID = "000000000000"; // Identificación fiscal
                    mySN.Phone1 = "2255"; // Teléfono

                    // Se intenta agregar el cliente a la base de datos
                    error = mySN.Add();

                    if (error != 0)
                    {
                        Console.WriteLine("Error al crear el cliente: " + myCompany.GetLastErrorDescription(), "Error");
                    }
                    else
                    {
                        Console.WriteLine("Cliente creado exitosamente");
                    }


                    if (myCompany.Connected == true)
                    {
                        myCompany.Disconnect();
                        Console.WriteLine("SAP B1 ha sido desconectado");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error general al conectar a SAP B1: " + ex.Message);
            }
        }
    }
}
