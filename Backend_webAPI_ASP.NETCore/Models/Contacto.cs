using System.Collections.Generic;

namespace Backend_webAPI_ASP.NETCore.Models;

public class Contacto
{
    public int IdContacto { get; set; }
    public string? Nombre { get; set; }
    public string? Apellido { get; set; }
    public string? Telefono { get; set; }
    public string? Correo { get; set; }
}
