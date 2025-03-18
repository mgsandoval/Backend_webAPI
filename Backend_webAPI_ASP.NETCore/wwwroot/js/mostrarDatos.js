function mostrarDatos() {
    let input = document.getElementById("inputExcel");

    const formData = new FormData();

    formData.append("ArchivoExcel", input.files[0]);

    fetch("Home/MostrarDatos", {
        method: "POST",
        body: formData
    })
        .then((response) => { return response.json() })
        .then((dataJson) => {

            dataJson.forEach((item) => {
                $("#tbData tbody").append(
                    $("<tr>").append(
                        $("<td>").text(item.codigo),
                        $("<td>").text(item.nombre_razon_social),
                        $("<td>").text(item.tipo_cliente),
                        $("<td>").text(item.moneda),
                        $("<td>").text(item.telefono1),
                        $("<td>").text(item.telefono_movil),
                        $("<td>").text(item.correo_electronico),
                        $("<td>").text(item.rtn),
                        $("<td>").text(item.direccion),
                        $("<td>").text(item.vendedor),
                        $("<td>").text(item.territorio),
                        $("<td>").text(item.nombre_completo),
                        $("<td>").text(item.nombre),
                        $("<td>").text(item.apellido),
                        $("<td>").text(item.telefono_fijo),
                        $("<td>").text(item.movil_personal),
                        $("<td>").text(item.correo_electronico2),
                        $("<td>").text(item.destino),
                        $("<td>").text(item.id_direccion),
                        $("<td>").text(item.nombre_direccion2),
                        $("<td>").text(item.nombre_direccion3),
                        $("<td>").text(item.ciudad),
                        $("<td>").text(item.condado),
                        $("<td>").text(item.condiciones_pago),
                        $("<td>").text(item.lista_precios),
                        $("<td>").text(item.limite_credito),
                        $("<td>").text(item.cuenta_mayor_sucursal)
                    )
                );
            });
        });
    input = document.getElementById("inputExcel").value = "";
}
