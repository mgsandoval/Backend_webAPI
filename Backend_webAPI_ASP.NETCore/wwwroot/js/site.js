function mostrarDatos() {
    const input = document.getElementById("inputExcel")

    const formData = new FormData()

    formData.append("ArchivoExcel", input.files[0])

    fetch("Home/MostrarDatos", {
        method: "POST",
        body: formData
    })
        .then((response) => { return response.json() })
        .then((dataJson) => {

            dataJson.forEach((item) => {
                $("#tbData tbody").append(
                    $("<tr>").append(
                        $("<td>").text(item.nombre),
                        $("<td>").text(item.apellido),
                        $("<td>").text(item.telefono),
                        $("<td>").text(item.correo)
                    )
                )
            })
        })
}

function enviarDatos() {
    
}