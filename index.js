// Verificar que el elemento existe antes de agregar el event listener
const excelMp = document.getElementById('excelFile1');
if (excelMp) {
    excelMp.addEventListener('change', function(event) {
        let archivoMp = event.target.files[0];
        if (archivoMp) {
            document.querySelector('#uploadBox1 .file-upload-text').innerText = `Archivo seleccionado: ${archivoMp.name}`;
            document.getElementById('uploadBox1').classList.add('file-upload-success');

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                window.archivoMp = XLSX.read(data, { type: 'array' });
            };
            reader.readAsArrayBuffer(archivoMp);
        }
    });
}

// Manejo del archivo de Tienda Nube (asegúrate de que el ID existe)
const excelTn = document.getElementById('excelFile2');
if (excelTn) {
    excelTn.addEventListener('change', function(event) {
        let archivoTn = event.target.files[0];
        if (archivoTn) {
            document.querySelector('#uploadBox2 .file-upload-text').innerText = `Archivo seleccionado: ${archivoTn.name}`;
            document.getElementById('uploadBox2').classList.add('file-upload-success');

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                window.archivoTn = XLSX.read(data, { type: 'array' });
            };
            reader.readAsArrayBuffer(archivoTn);
        }
    });
}

// Función para obtener datos de Tienda Nube
const obtenerTn = () => {
    if (!window.archivoTn) {
        alert("Archivo de Tienda Nube no cargado.");
        return [];
    }
    const hojaNombre = window.archivoTn.SheetNames[0];
    const datos = XLSX.utils.sheet_to_json(window.archivoTn.Sheets[hojaNombre], { cellDates: true });

    const agrupados = {};
    datos.forEach((dato, index) => {
        if (index === 0) return; // Saltar la primera fila si es encabezado

        const numeroOrden = dato['Número de orden']; 
        const producto = String(dato['Nombre del producto']); 

        if (agrupados[numeroOrden]) {
            agrupados[numeroOrden]['Cantidad'] += 1;
            agrupados[numeroOrden]['Nombre del producto'] += ' - ' + producto;
        } else {
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

    return Object.values(agrupados);
};

// Función para obtener datos de Mercado Pago
const obtenerMp = () => {
    if (!window.archivoMp) {
        console.error("Archivo de Mercado Pago no cargado.");
        return [];
    }
    const nombreHoja = window.archivoMp.SheetNames[0];
    const datos = XLSX.utils.sheet_to_json(window.archivoMp.Sheets[nombreHoja], { cellDates: true });

    return datos.map(dato => ({
        'Número de operación de Mercado Pago': Number(dato['NÃºmero de operaciÃ³n de Mercado Pago (operation_id)']),
        'Valor del producto': Number(dato['Valor del producto (transaction_amount)']),
        'Tarifa de Mercado Pago': Number(dato['Tarifa de Mercado Pago (mercadopago_fee)']),
        'Comisión por uso de plataforma de terceros': Number(dato['ComisiÃ³n por uso de plataforma de terceros (marketplace_fee)']),
        'Monto recibido': Number(dato['Monto recibido (net_received_amount)']),
        'Cuotas (installments)': Number(dato['Cuotas (installments)']),
        'Costos de financiación (financing_fee)': Number(dato['Costos de financiaciÃ³n (financing_fee)'])
    }));
};

// Función para cruzar la información
const cruzarInfo = () => {
    if (!window.archivoMp || !window.archivoTn) {
        console.error('Por favor, carga ambos archivos antes de cruzar la información.');
        return [];
    }

    const datosActualizados = obtenerTn();
    const datosMp = obtenerMp();

    return datosActualizados.map(dato => {
        const identificador = parseInt(dato['Identificador de la transacción en el medio de pago'], 10);
        const datosMpEncontrado = datosMp.find(mp => mp['Número de operación de Mercado Pago'] === identificador);
        console.log(identificador)
        console.log(datosMpEncontrado)

        if (datosMpEncontrado) {
            const valorProducto = Math.abs(datosMpEncontrado['Valor del producto']);
            const tarifaMercadoPago = Math.abs(datosMpEncontrado['Tarifa de Mercado Pago']);
            const comisionTerceros = Math.abs(datosMpEncontrado['Comisión por uso de plataforma de terceros']);
            const costosFinanciacion = Math.abs(datosMpEncontrado['Costos de financiación (financing_fee)']);
            const montoRecibido = datosMpEncontrado['Monto recibido'];
            const impuestoIIBB = valorProducto - montoRecibido - costosFinanciacion - tarifaMercadoPago;

            return {
                'Estado del pago': dato['Estado del pago'],
                'Fecha': dato['Fecha'],
                'Número de orden': dato['Número de orden'],
                'Nombre de cliente': dato['Nombre del comprador'],
                'Nombre del producto': dato['Nombre del producto'],
                'Cantidad': dato['Cantidad'],
                'Ingreso Bruto': valorProducto,
                'Impuesto IIBB': impuestoIIBB,
                'Cargo de Mercado Pago y Comisión de terceros': tarifaMercadoPago + comisionTerceros,
            };
        } else {
            return {
                'Estado del pago': dato['Estado del pago'],
                'Fecha': dato['Fecha'],
                'Número de orden': dato['Número de orden'],
                'Nombre de cliente': dato['Nombre del comprador'],
                'Nombre del producto': dato['Nombre del producto'],
                'Cantidad': dato['Cantidad'],
            };
        }
    });
};

// Función para guardar datos en Excel
const guardarEnExcel = async (datos, nombreArchivo) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Datos Cruzados');

    worksheet.columns = [
        { header: 'Estado del pago', key: 'Estado del pago', width: 20 },
        { header: 'Fecha', key: 'Fecha', width: 15 },
        { header: 'Número de orden', key: 'Número de orden', width: 20 },
        { header: 'Nombre de cliente', key: 'Nombre de cliente', width: 25 },
        { header: 'Nombre del producto', key: 'Nombre del producto', width: 30 },
        { header: 'Cantidad', key: 'Cantidad', width: 10 },
        { header: 'Ingreso Bruto', key: 'Ingreso Bruto', width: 15 },
        { header: 'Impuesto IIBB', key: 'Impuesto IIBB', width: 15 },
        { header: 'Cargo de Mercado Pago y Comisión de terceros', key: 'Cargo de Mercado Pago y Comisión de terceros', width: 15 },
    ];

    datos.forEach(item => {
        worksheet.addRow(item);
    });

    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' }
    };

    datos.forEach((item, index) => {
        const row = worksheet.getRow(index + 2);
        row.getCell('Ingreso Bruto').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
        row.getCell('Impuesto IIBB').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0CB' } };
        row.getCell('Cargo de Mercado Pago y Comisión de terceros').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000' } };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nombreArchivo;
    a.click();
    window.URL.revokeObjectURL(url);
};

// Evento para generar el Excel al presionar el botón
document.getElementById('generarBtn').addEventListener('click', async () => {
    const datos = cruzarInfo();
    if (datos.length > 0) {
        await guardarEnExcel(datos, 'Datos_Cruzados.xlsx');
    } else {
        console.error('No hay datos para procesar.');
    }
});
