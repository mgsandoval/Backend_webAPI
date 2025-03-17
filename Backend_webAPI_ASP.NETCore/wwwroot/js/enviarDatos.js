function enviarDatos() {
    //Obtener los datos de la tabla
    alert("Hola mundo!!!!!!!!!!");
    let table = document.getElementById("tbData");
    let data = [];
    for (let i = 1; i < table.rows.length; i++) {
        let row = table.rows[i];
        let rowData = {
            codigo: row.cells[0].innerText,
            nombre_razon_social: row.cells[1].innerText,
            tipo_cliente: row.cells[2].innerText,
            moneda: row.cells[3].innerText,
            telefono1: row.cells[4].innerText,
            telefono_movil: row.cells[5].innerText,
            correo_electronico: row.cells[6].innerText,
            rtn: row.cells[7].innerText,
            direccion: row.cells[8].innerText,
            vendedor: row.cells[9].innerText,
            territorio: row.cells[10].innerText,
            nombre_completo: row.cells[11].innerText,
            nombre: row.cells[12].innerText,
            apellido: row.cells[13].innerText,
            telefono_fijo: row.cells[14].innerText,
            movil_personal: row.cells[15].innerText,
            correo_electronico2: row.cells[16].innerText,
            destino: row.cells[17].innerText,
            id_direccion: row.cells[18].innerText,
            nombre_direccion2: row.cells[19].innerText,
            nombre_direccion3: row.cells[20].innerText,
            ciudad: row.cells[21].innerText,
            condado: row.cells[22].innerText,
            condiciones_pago: row.cells[23].innerText,
            lista_precios: row.cells[24].innerText,
            limite_credito: row.cells[25].innerText,
            cuenta_mayor_sucursal: row.cells[26].innerText
        };
        data.push(rowData);
    }
    alert("Los datos han sido enviados");

    //fetch("Home/EnviarDatos", {
    //    method: "POST",
    //    headers: {
    //        "Content-Type": "application/json"
    //    },
    //    body: JSON.stringify(data)
    //})
    //    .then((response) => { return response.json() })
    //    .then((dataJson) => {
    //        console.log(dataJson);
    //    });
}
