function enviarDatos() {
    // Obtenemos la tabla
    let table = document.getElementById("tbData");
    // Creamos un arreglo para almacenar los datos
    let data = [];

    console.log("Preparando datos para enviar...");

    // Recorremos la tabla y almacenamos los datos en el arreglo
    for (let i = 1; i < table.rows.length; i++) {
        // Obtenemos la fila
        let row = table.rows[i];
        // Obtenemos los datos de la fila
        let rowData = {
            Codigo: row.cells[0].innerText.trim(),
            Nombre_razon_social: row.cells[1].innerText.trim(),
            Tipo_cliente: row.cells[2].innerText.trim(),
            Moneda: row.cells[3].innerText.trim(),
            Telefono1: row.cells[4].innerText.trim(),
            Telefono_movil: row.cells[5].innerText.trim(),
            Correo_electronico: row.cells[6].innerText.trim(),
            RTN: row.cells[7].innerText.trim(),
            Direccion: row.cells[8].innerText.trim(),
            Vendedor: row.cells[9].innerText.trim(),
            Territorio: row.cells[10].innerText.trim()
        };
        data.push(rowData);
    };

    fetch("Home/EnviarDatos", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify(data)
    })
        // Response 1: Obtenemos la respuesta del servidor y la mostramos en la consola
        .then(response => {
            if (!response.ok) {
                console.log("[Response 1.]");
                throw new Error("Error en la API: " + response.statusText);
            }
            return response.json();
        })
        // Response 2: Obtenemos la respuesta del servidor y la mostramos en la consola
        .then(responseJson => {
            alert(responseJson.message);
            console.log("Respuesta del servidor:", responseJson);
            console.log("[Response 2.]");
        })
        // Catch: Mostramos un mensaje de error en la consola
        .catch(error => {
            console.log("Error al enviar los datos:", error);
            console.log("[Catch.]");
            alert("Ocurrió un error al enviar los datos. Revisa la consola.");
        });

    //input = document.getElementById("inputExcel").value = "";
}

