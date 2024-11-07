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
    const datos = XLSX.utils.sheet_to_json(window.archivoTn.Sheets[hojaNombre], { cellDates: true, raw: false });

    // Inicializamos el objeto agrupados fuera del map
    const agrupados = {};

    datos.map((dato) => {
        const numeroOrden = parseInt(dato['Número de orden']); 
        const producto = String(dato['Nombre del producto'] || '');
        
        if (agrupados[numeroOrden]) {
            agrupados[numeroOrden]['Cantidad'] += 1;
            agrupados[numeroOrden]['Nombre del producto'] += ' - ' + producto;
        } else {
            agrupados[numeroOrden] = {
                'Estado del pago': String(dato['Estado del pago'] || ''),
                'Fecha': dato['Fecha'] ? String(dato['Fecha']) : '',
                'Número de orden': numeroOrden,
                'Cantidad': 1,
                'Nombre del comprador': String(dato['Nombre del comprador'] || ''),
                'Nombre del producto': producto,
                'Medio de pago': String(dato['Medio de pago'] || ''),
                'Identificador de la transacción en el medio de pago': parseInt(dato['Identificador de la transacción en el medio de pago'] || 0)
            };
        }
    });

    // Convertir el objeto agrupado en un array de valores
    return Object.values(agrupados);
};

// Función para obtener datos de Mercado Pago
const obtenerMp = () => {
    if (!window.archivoMp) {
        alert("Archivo de Mercado Pago no cargado.");
        return [];
    }
    const nombreHoja = window.archivoMp.SheetNames[0];
    const datos = XLSX.utils.sheet_to_json(window.archivoMp.Sheets[nombreHoja], { cellDates: true });

    return datos.map(dato => ({
        'Número de operación de Mercado Pago': Number(dato['NÃºmero de operaciÃ³n de Mercado Pago (operation_id)']),
        'Medio de pago': String(dato['Medio de pago']),
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
        alert('Por favor, carga ambos archivos antes de cruzar la información.');
        return [];
    }

    const datosTn = obtenerTn();
    const datosMp = obtenerMp();
    
    const fechaHoy = new Date().toLocaleDateString();

    return datosTn
        .filter(dato => dato['Estado del pago'] !== 'Rechazado') // Filtra los elementos con 'Estado del pago' == 'Rechazado'
        .map(dato => {
            const identificador = parseInt(dato['Identificador de la transacción en el medio de pago'], 10);
            const datosMpEncontrado = datosMp.find(mp => mp['Número de operación de Mercado Pago'] === identificador);

            if (datosMpEncontrado) {
                const valorProducto = Math.abs(Number(datosMpEncontrado['Valor del producto']));
                const tarifaMercadoPago = Math.abs(Number(datosMpEncontrado['Tarifa de Mercado Pago']));
                const comisionTerceros = Math.abs(Number(datosMpEncontrado['Comisión por uso de plataforma de terceros']));
                const costosFinanciacion = Math.abs(Number(datosMpEncontrado['Costos de financiación (financing_fee)']));
                const montoRecibido = Math.abs(Number(datosMpEncontrado['Monto recibido']));
                const impuestoDebitoCredito = valorProducto * 0.006;
                
                const impuestoIIBB = Math.abs(valorProducto - montoRecibido - comisionTerceros - costosFinanciacion - tarifaMercadoPago - impuestoDebitoCredito);
                const comisionTotal = Math.abs(tarifaMercadoPago + comisionTerceros);

                return {
                    'Estado del pago': dato['Estado del pago'],
                    'Fecha': fechaHoy,
                    'Número de orden': dato['Número de orden'],
                    'Nombre de cliente': dato['Nombre del comprador'],
                    'Nombre del producto': dato['Nombre del producto'],
                    'Cantidad': dato['Cantidad'],
                    'Ingreso Bruto': valorProducto,
                    'Impuesto IIBB': impuestoIIBB,
                    'Impuesto Deb y Cred': impuestoDebitoCredito,
                    'Intereses Cuotas': costosFinanciacion,
                    'Cargo de Mercado Pago y Comisión de terceros': comisionTotal,
                    'Ingreso Neto': montoRecibido
                };
            } else {
                return {
                    'Estado del pago': dato['Estado del pago'],
                    'Fecha': fechaHoy,
                    'Número de orden': dato['Número de orden'],
                    'Nombre de cliente': dato['Nombre del comprador'],
                    'Nombre del producto': dato['Nombre del producto'],
                    'Cantidad': dato['Cantidad'],
                    'Ingreso Bruto': '',
                    'Impuesto IIBB': '',
                    'Impuesto Deb y Cred': '',
                    'Intereses Cuotas': '',
                    'Cargo de Mercado Pago y Comisión de terceros': '',
                    'Ingreso Neto': ''
                };
            }
        });
};

const guardarEnExcel = async (datos, nombreArchivo) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Datos Cruzados');

    worksheet.columns = [
        { header: 'ESTADO DE PAGO', key: 'Estado del pago', width: 15 },
        { header: 'FECHA', key: 'Fecha', width: 10 },
        { header: 'NÚMERO ORDEN', key: 'Número de orden', width: 15 },
        { header: 'NOMBRE CLIENTE', key: 'Nombre de cliente', width: 25 },
        { header: 'CANTIDAD', key: 'Cantidad', width: 10 },
        { header: 'PRODUCTO', key: 'Nombre del producto', width: 50 },
        { header: 'INGRESO BRUTO', key: 'Ingreso Bruto', width: 15 },
        { header: 'IMPUESTOS IIBB', key: 'Impuesto IIBB', width: 15 },
        { header: 'IMPUESTOS DEB Y CRED', key: 'Impuesto Deb y Cred', width: 15 },
        { header: 'INTERESES CUOTAS', key: 'Intereses Cuotas', width: 15 },
        { header: 'CARGO MERCADOPAGO Y COMISICIONES TERCEROS', key: 'Cargo de Mercado Pago y Comisión de terceros', width: 15 },
        { header: 'INGRESO NETO', key: 'Ingreso Neto', width: 15 },
    ];

    // Definir el estilo del borde
    const borderStyle = {
        top: { style: 'thin', color: { argb: '000000' } },
        left: { style: 'thin', color: { argb: '000000' } },
        bottom: { style: 'thin', color: { argb: '000000' } },
        right: { style: 'thin', color: { argb: '000000' } }
    };

    // Agregar los datos a la hoja
    datos.forEach(item => {
        const row = worksheet.addRow(item);
        
        // Agregar bordes a cada celda de la fila
        row.eachCell((cell) => {
            cell.border = borderStyle;
        });
    });

    // Estilo para la cabecera
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' }
    };

    // Estilos de fondo para ciertas celdas
    datos.forEach((item, index) => {
        const row = worksheet.getRow(index + 2); // Comienza en la fila 2 para evitar la cabecera

        // Aplicando estilos de fondo
        row.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '62d397' } }; // INGRESO BRUTO
        row.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'eb9999' } }; // IMPUESTO IIBB
        row.getCell(9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'eb9999' } }; // IMPUESTOS DEB Y CRED
        row.getCell(10).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'eb9999' } }; // INTERESES CUOTAS
        row.getCell(11).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'eb9999' } }; // CARGO MERCADOPAGO Y COMISICIONES TERCEROS
        row.getCell(12).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '62d397' } }; // INGRESO NETO
    });

    // Generar el archivo
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
        alert('No hay datos para procesar.');
    }
});
