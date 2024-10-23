const xlsx = require('xlsx');

const obtenerTn = () => {
    // Leer archivo CSV
    const excel = xlsx.readFile("./ventaTN.csv"); // Eliminamos {type: 'binary'}

    // Convertir la hoja a JSON
    const hojaNombre = excel.SheetNames[0];  // Obtener el nombre de la primera hoja
    const datos = xlsx.utils.sheet_to_json(excel.Sheets[hojaNombre], {
        cellDates: true
    });

    // Crear un objeto para agrupar por 'Número de orden'
    const agrupados = {};

    datos.forEach(dato => {
        const numeroOrden = dato['Número de orden'];
        const producto = dato['Nombre del producto'];

        // Si el número de orden ya está en el objeto, incrementar la cantidad y concatenar los productos
        if (agrupados[numeroOrden]) {
            agrupados[numeroOrden]['Cantidad'] += 1;
            agrupados[numeroOrden]['Nombre del producto'] += ' - ' + producto;
        } else {
            // Si no está, inicializar el objeto con los datos
            agrupados[numeroOrden] = {
                'Estado del pago': dato['Estado del pago'],
                'Fecha': dato['Fecha'],
                'Número de orden': numeroOrden,
                'Cantidad': 1,
                'Nombre del comprador': dato['Nombre del comprador'],
                'Nombre del producto': producto,
                'Medio de pago': dato['Medio de pago'],
                'Identificador de la transacción en el medio de pago': dato['Identificador de la transacción en el medio de pago']
            };
        }
    });

    // Convertir el objeto de agrupados en un array para procesar y guardar
    const datosAgrupados = Object.values(agrupados);

    return datosAgrupados;
}

const guardarEnExcel = (datos, nombreArchivo) => {
    // Crear hoja a partir de los datos procesados
    const hoja = xlsx.utils.json_to_sheet(datos);
    const libro = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(libro, hoja, 'Datos Cruzados');

    // Guardar en archivo Excel
    xlsx.writeFile(libro, nombreArchivo);
}

// Ejemplo de uso
const datosFinales = obtenerTn();
guardarEnExcel(datosFinales, 'datos_cruzados.xlsx');
