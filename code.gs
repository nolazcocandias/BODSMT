//------------------------ ID CLIENTES/CTRL BODEGA--------------------------------------------------------
const CRTLBOD = "1oLY-_7Ql1toJl2o9ZgyErbC_zorS8Tr2hEYMAeT0knk";
const FSMT = "1T7HG8KEs-Ob7f3M5atEVN9Yn7IeORGp0QGvggB62ELw";
const HOJA_ORDENES = "Ordenes de Produccion";
//--------------MOSTRAR FORMULARIOS-----------------------------------------------------------------------
function onOpen() {
  
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Summit')
    .addSubMenu(ui.createMenu('Ingreso')
      .addItem('Ingreso Productos', 'mostrarFormularioIngresoMP')
      //.addItem('Ingreso Orden de Trabajo', 'mostrarFormularioIngresoOC')
      .addItem('Ingresar Insumo', 'mostrarFormularioIngresoInsumo')
    )
    .addItem('Reubicar Productos', 'mostrarFormularioReubicarSMT')
    .addItem('Salida por Ubicacion Selectiva', 'mostrarFormularioSalida')
    .addItem('Salida por Picking Automatico', 'mostrarFormularioSalidaRapida')
    .addToUi();

  ui.createMenu('Imprimir')
    .addSubMenu(ui.createMenu('Imprimir Documento')
      .addItem('Documento Ingreso', 'imprimirDocumentoIngreso')
      .addItem('Documento Reubicacion', 'imprimirDocumentoReubicacion')
      .addItem('Documento Picking', 'imprimirDocumentoPicking')
    )
    .addToUi();
}

//----------------------- IMPRIMIR DOCUMENTOS-------------------------------------------------------------
function imprimirDocumentoIngreso() {
  abrirPDF('1Fuojppxk56U_K5RTld8ivGTb4EEVuUv4');
}
function imprimirDocumentoReubicacion() {
  abrirPDF('1uzfTFuLExzBk6OlebJGPZ9lCsebd3WZc');
}
function imprimirDocumentoPicking() {
  abrirPDF('1dYuytETmFKqZrHg0XmXoTUscAOKPiYUt');
}
function abrirPDF(fileId) {
  var file = DriveApp.getFileById(fileId);
  var url = file.getUrl();
  var html = HtmlService.createHtmlOutput('<html><script>'
    + 'window.open("' + url + '");google.script.host.close();'
    + '</script></html>')
    .setWidth(60)
    .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo PDF...');
}
//----------------------- FORMULARIOS SUMMIT--------------------------------------------------------------

function mostrarFormularioSalidaRapida() {
  const html = HtmlService.createHtmlOutputFromFile('SalidaRapida')
    .setWidth(900)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Salida rápida por SKU');
}

function mostrarFormularioReubicarSMT() {
  const html = HtmlService.createHtmlOutputFromFile('ReubicarSMT')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Reubicar Productos");
}
function mostrarFormularioIngresoMP() {
  const html = HtmlService.createHtmlOutputFromFile('IngresoMP')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Ingreso Productos");
}
function mostrarFormularioIngresoOC() {
  const html = HtmlService.createHtmlOutputFromFile('IngresoOC')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Ingreso Orden de Trabajo");
}
function mostrarFormularioIngresoInsumo() {
  const html = HtmlService.createHtmlOutputFromFile('IngresoInsumo')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Ingreso Insumos");
}
function mostrarFormularioSalida() {
  const html = HtmlService.createHtmlOutputFromFile('SalidaSMT')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Salida Summit');
}

//-------------------------FUNCIONES SUMMIT------------------------------------------------------------------------
function obtenerOpcionesSMT(tipo, formulario) {
  const ss = SpreadsheetApp.openById(FSMT);
  const sheet = ss.getSheetByName("Maestro");
  const datos = sheet.getRange("A3:D" + sheet.getLastRow()).getValues(); // Rango de columnas A a D

  let opciones;

  if (formulario === 'IngresoMP') {
    // Filtro para IngresoMP: ACTIVO y Si
    opciones = datos
      .filter(fila => fila[0] === "ACTIVO" && fila[1] === "Si")
      .map(fila => fila[3]) // Columna D: Datos finales para la lista
      .filter((valor, indice, self) => valor && self.indexOf(valor) === indice); // Elimina duplicados y valores vacíos

  } else if (formulario === 'IngresoOC') {
    // Filtro para IngresoOC: ACTIVO solamente
    opciones = datos
      .filter(fila => fila[0] === "ACTIVO")
      .map(fila => fila[3]) // Columna D: Datos finales para la lista
      .filter((valor, indice, self) => valor && self.indexOf(valor) === indice); // Elimina duplicados y valores vacíos
  }

  return { tipo, opciones };
}

function guardarDatosIngresoSMT(datos) {
  const ssRecepcion = SpreadsheetApp.openById(FSMT);
  const hojaRecepcion = ssRecepcion.getSheetByName('Recepcion Materia Prima');
  
  const ssUbicaciones = SpreadsheetApp.openById(CRTLBOD);
  const hojaUbicaciones = ssUbicaciones.getSheetByName('Ubicaciones');
  const hojaMaestro = ssRecepcion.getSheetByName('Maestro');

  if (!hojaRecepcion) {
    throw new Error('No se encontró la hoja de Recepcion Materia Prima');
  }
  if (!hojaUbicaciones || !hojaMaestro) {
    throw new Error('No se encontraron todas las hojas necesarias');
  }

  try {
    // Obtener datos del Maestro en una sola lectura (Columnas C y D)
    const datosMaestro = hojaMaestro.getRange(2, 3, hojaMaestro.getLastRow() - 1, 2).getValues();
    const datosUbicaciones = hojaUbicaciones.getRange(2, 4, hojaUbicaciones.getLastRow() - 1, 1).getValues();
    
    datos.productos.forEach(producto => {
      // Guardar en la hoja "Recepcion Materia Prima" (ahora incluye proveedor al final)
      hojaRecepcion.appendRow([
        datos.fecha,
        datos.guia,
        producto.sku,
        producto.producto,
        producto.cantidad,
        producto.pallets,
        producto.ubicacion, // Guardar la ubicación en la columna G
        datos.proveedor     // <--- NUEVO campo proveedor al final
      ]);

      // Buscar el SKU para el producto actual basándose en la descripción
      const sku = producto.sku;

      // Preparar actualizaciones para la hoja "Ubicaciones"
      const updates = [];
      const ubicacionIndex = datosUbicaciones.findIndex(row => row[0] === producto.ubicacion);
      if (ubicacionIndex !== -1) {
        const filaUbicacion = ubicacionIndex + 2;

        const datosUbicacionesParte1 = [
          [
            datos.fecha,
            datos.guia,
          ]
        ];

        const datosUbicacionesParte2 = [
          [
            producto.producto,
            producto.cantidad,
            datos.origen
          ]
        ];

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 5, 1, 2),
          values: datosUbicacionesParte1
        });

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 8, 1, 3),
          values: datosUbicacionesParte2
        });

        // Actualización adicional para la columna "O" con el valor "SUMMIT"
        const datosUbicacionesParte3 = [
          ["SUMMIT"]
        ];

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 15, 1, 1),
          values: datosUbicacionesParte3
        });

        // Actualización adicional para la columna "G" con el SKU correspondiente
        const datosUbicacionesParte4 = [
          [sku]
        ];

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 7, 1, 1), // Columna G es la columna 7
          values: datosUbicacionesParte4
        });

        // Actualización adicional para la columna "I" con la cantidad correspondiente
        const datosUbicacionesParte5 = [
          [producto.cantidad]
        ];

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 9, 1, 1), // Columna I es la columna 9
          values: datosUbicacionesParte5
        });

        // Actualización adicional para la columna "J" con el valor "PROVEEDOR"
        const datosUbicacionesParte6 = [
          ["PROVEEDOR"]
        ];

        updates.push({
          sheet: hojaUbicaciones,
          range: hojaUbicaciones.getRange(filaUbicacion, 10, 1, 1), // Columna J es la columna 10
          values: datosUbicacionesParte6
        });
      }

      // Ejecutar actualizaciones para este producto en la hoja "Ubicaciones"
      updates.forEach(update => {
        update.range.setValues(update.values);
      });
    });

    return {
      success: true,
      message: 'Datos guardados correctamente'
    };

  } catch (error) {
    Logger.log(`Error al guardar datos: ${error.message}`);
    throw new Error(`Error al guardar datos: ${error.message}`);
  }
}

function obtenerSKUSMT(producto) {
  const ss = SpreadsheetApp.openById(FSMT);
  const hojaMaestro = ss.getSheetByName('Maestro');
  const datosMaestro = hojaMaestro.getRange("C3:F" + hojaMaestro.getLastRow()).getValues(); // Columnas C y D
  
  const filaProducto = datosMaestro.find(row => row[1] === producto);
  if (filaProducto) {
    return filaProducto[0];  // SKU está en la columna C (índice 0)
  } else {
    return '';
  }
}
function buscarSkuOrigenEnMaestro(sku) {
  const hoja = SpreadsheetApp.openById(FSMT).getSheetByName("Maestro");
  const datos = hoja.getRange("C2:F" + hoja.getLastRow()).getValues(); // Obtener datos desde la fila 2 hasta la última fila de las columnas C a F

  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === sku) { // Comparar con la columna C (índice 0)
      return datos[i][3]; // Devolver el valor de la columna F (índice 3)
    }
  }
  
  throw new Error('No se encontró el SKU de origen en la hoja Maestro');
}

function llenarListaUbicacionesOC() {
  const hoja = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  const datos = hoja.getRange("BD3:BH" + hoja.getLastRow()).getValues();

  const ubicacionesPorSKU = {};

  datos.forEach(fila => {
    const ubicacion = fila[0]; // Columna BD (índice 0)
    const sku = fila[1]; // Columna BE (índice 1)
    const referencia = fila[4]; // Columna BH (índice 4)

    if (!ubicacionesPorSKU[sku]) {
      ubicacionesPorSKU[sku] = [];
    }
    ubicacionesPorSKU[sku].push(ubicacion);

    if (sku !== referencia) {
      if (!ubicacionesPorSKU[referencia]) {
        ubicacionesPorSKU[referencia] = [];
      }
      ubicacionesPorSKU[referencia].push(ubicacion);
    }
  });

  return ubicacionesPorSKU;
}

function actualizarUbicaciones(skuOrigenInput) {
  const fila = skuOrigenInput.closest('tr');
  const skuInput = fila.querySelector('.sku');
  const ubicacionSelect = fila.querySelector('.ubicacion');
  const skuOrigen = skuOrigenInput.value;
  const sku = skuInput.value;

  let ubicaciones = [];

  if (skuOrigen === sku) {
    ubicaciones = ubicacionesPorSKU[sku] || [];
  } else {
    ubicaciones = ubicacionesPorSKU[skuOrigen] || [];
  }

  cargarOpcionesUbicacion(ubicacionSelect, ubicaciones);
  ubicacionSelect.disabled = ubicaciones.length === 0;
  if (ubicaciones.length === 0) {
    ubicacionSelect.classList.add('ubicacion-bloqueada');
    Swal.fire('Error', 'No hay ubicaciones disponibles para el SKU proporcionado', 'error');
  } else {
    ubicacionSelect.classList.remove('ubicacion-bloqueada');
  }
}

function cargarOpcionesUbicacion(select, ubicaciones) {
  select.innerHTML = '<option value="" disabled selected>-- Selecciona una ubicación --</option>';
  ubicaciones.forEach(ubicacion => {
    const opt = document.createElement("option");
    opt.value = ubicacion;
    opt.textContent = ubicacion;
    select.appendChild(opt);
  });
}

function obtenerMaximosPorUbicacion() {
  try {
    const ss = SpreadsheetApp.openById(CRTLBOD);
    const hoja = ss.getSheetByName('Ubicaciones');
    const datos = hoja.getRange('BD3:BJ' + hoja.getLastRow()).getValues();
    
    // Crear objeto con los máximos por ubicación
    const maximosPorUbicacion = {};
    datos.forEach(fila => {
      if (fila[0]) { // Si hay un valor en la columna BD
        maximosPorUbicacion[fila[0]] = fila[fila.length - 1]; // El último valor es la columna BJ
      }
    });
    
    return maximosPorUbicacion;
  } catch (error) {
    Logger.log('Error al obtener máximos por ubicación: ' + error.toString());
    return {};
  }
}
function guardarOCSMT(productos) {
  const sheetOrdenes = SpreadsheetApp.openById(FSMT).getSheetByName(HOJA_ORDENES);
  const sheetUbicaciones = SpreadsheetApp.openById(CRTLBOD).getSheetByName('Ubicaciones');

  productos.forEach(producto => {
    const nuevaFila = [
      producto.fecha,
      producto.oc,
      producto.cliente,
      producto.skuOrigen,
      producto.sku,
      producto.descripcion,
      producto.maximoAProducir,
      producto.ubicacion,
      producto.unidades,
      "EN PREPARACION"
    ];
    sheetOrdenes.appendRow(nuevaFila);

    // Proceso en la hoja "Ubicaciones"
    const ubicacionesData = sheetUbicaciones.getDataRange().getValues();
    for (let i = 1; i < ubicacionesData.length; i++) {
      if (ubicacionesData[i][3] === producto.ubicacion) { // Comparar con la columna D (índice 3)
        const nuevaCantidad = ubicacionesData[i][8] - producto.unidades; // Restar unidades a la columna I (índice 8)
        if (nuevaCantidad === 0) {
          // Borrar columnas E, F, G, H, I, J, O
          sheetUbicaciones.getRange(i + 1, 5, 1, 6).clearContent(); // E, F, G, H, I, J
          sheetUbicaciones.getRange(i + 1, 15, 1, 1).clearContent(); // O
        } else {
          // Actualizar la columna I con la nueva cantidad
          sheetUbicaciones.getRange(i + 1, 9).setValue(nuevaCantidad);
        }
        break; // Salir del bucle una vez encontrada la coincidencia
      }
    }
  });

  return { success: true };
}

function obtenerDatosUbicacionesSMT() {
  const hoja= SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones")
  const datos = hoja.getRange("BS3:BV" + hoja.getLastRow()).getValues();///////////////AC2:AG
  return datos.filter(fila => fila.some(celda => celda !== ""));
}

function registrarReubicacionSMT(datos) {
  const ssu = SpreadsheetApp.openById(CRTLBOD);
  const ss = SpreadsheetApp.openById(FSMT);
  const hojas = {
    reubicaciones: ss.getSheetByName('Reubicaciones'),
    ubicaciones: ssu.getSheetByName('Ubicaciones'),
    maestro: ss.getSheetByName('Maestro')
  };

  if (!Object.values(hojas).every(hoja => hoja)) {
    throw new Error('No se encontraron todas las hojas necesarias');
  }

  try {
    const updates = [];

    const filaReubicacion = [
      [
        datos.fecha,
        datos.ubicacionAntigua,
        datos.sku,
        datos.descripcion,
        datos.cantidadEnUbicacion,
        datos.cantidadAMover,
        datos.ubicacionNueva
      ]
    ];

    updates.push({
      sheet: hojas.reubicaciones,
      range: hojas.reubicaciones.getRange(hojas.reubicaciones.getLastRow() + 1, 1, 1, filaReubicacion[0].length),
      values: filaReubicacion
    });

    const ubicacionesData = hojas.ubicaciones.getRange(1, 1, hojas.ubicaciones.getLastRow(), 15).getValues();
    const indexAntigua = ubicacionesData.findIndex(row => row[3] === datos.ubicacionAntigua);
    const indexNueva = ubicacionesData.findIndex(row => row[3] === datos.ubicacionNueva);

    if (indexAntigua === -1) {
      throw new Error('La ubicación antigua no se encontró');
    }

    const cantidadEnUbicacion = parseFloat(ubicacionesData[indexAntigua][8]) || 0;
    const cantidadAMover = parseFloat(datos.cantidadAMover);
    const nuevaCantidad = cantidadEnUbicacion - cantidadAMover;
    const origenIngreso = ubicacionesData[indexAntigua][9];
    const documentoOriginal = ubicacionesData[indexAntigua][5];

    if (nuevaCantidad <= 0) {
      const rangosALimpiar = [
        hojas.ubicaciones.getRange(indexAntigua + 1, 5, 1, 6), // Columnas E a J
        hojas.ubicaciones.getRange(indexAntigua + 1, 15, 1, 1) // Columna O
      ];
      rangosALimpiar.forEach(rango => updates.push({
        sheet: hojas.ubicaciones,
        range: rango,
        action: 'clear'
      }));
    } else {
      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexAntigua + 1, 9),
        values: [[nuevaCantidad]]
      });
    }

    if (indexNueva === -1) {
      const nuevaFila = [
        ['', '', '', datos.ubicacionNueva, datos.fecha, documentoOriginal, datos.sku,
         datos.descripcion, cantidadAMover, origenIngreso, '', '', '', '', 'SUMMIT']
      ];
      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(hojas.ubicaciones.getLastRow() + 1, 1, 1, 15),
        values: nuevaFila
      });
    } else {
      const cantidadEnUbicacionNueva = parseFloat(ubicacionesData[indexNueva][8]) || 0;
      const nuevaCantidadTotal = cantidadEnUbicacionNueva + cantidadAMover;

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 5),
        values: [[datos.fecha]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 6),
        values: [[documentoOriginal]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 7),
        values: [[datos.sku]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 8),
        values: [[datos.descripcion]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 9),
        values: [[nuevaCantidadTotal]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 10),
        values: [[origenIngreso]]
      });

      updates.push({
        sheet: hojas.ubicaciones,
        range: hojas.ubicaciones.getRange(indexNueva + 1, 15),
        values: [['SUMMIT']]
      });
    }

    updates.forEach(update => {
      if (update.action === 'clear') {
        update.range.clearContent();
      } else {
        update.range.setValues(update.values);
      }
    });

    Logger.log(`Reubicación completada: ${datos.ubicacionAntigua} -> ${datos.ubicacionNueva}`);

    return {
      success: true,
      message: 'Reubicación completada exitosamente'
    };

  } catch (error) {
    Logger.log(`Error en reubicación: ${error.message}`);
    throw new Error(`Error al procesar la reubicación: ${error.message}`);
  }
}
function obtenerUbicacionesDisponibles() {
  const ssu= SpreadsheetApp.openById(CRTLBOD);
  const hojaUbicaciones = ssu.getSheetByName("Ubicaciones");
  
  // Obtener todas las ubicaciones posibles (columna Z)
  const todasUbicaciones = hojaUbicaciones.getRange("Z2:Z" + hojaUbicaciones.getLastRow())
    .getValues()
    .flat()
    .filter(ubicacion => ubicacion !== "");

  // Obtener ubicaciones ocupadas (columna AC)
  const ubicacionesOcupadas = hojaUbicaciones.getRange("AC2:AC" + hojaUbicaciones.getLastRow())
    .getValues()
    .flat()
    .filter(ubicacion => ubicacion !== "");

  // Filtrar ubicaciones disponibles y excluir las que comienzan con R-8-B
  const ubicacionesDisponibles = todasUbicaciones.filter(ubicacion => 
    !ubicacionesOcupadas.includes(ubicacion) && 
    !ubicacion.toString().match(/^R-8-B\d+$/)  // Excluir ubicaciones R-8-B#
  );

  return {
    success: true,
    opciones: ubicacionesDisponibles
  };
}
function obtenerDatosUbicaciones() {
  const hoja= SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones")
  const datos = hoja.getRange("BL3:BP" + hoja.getLastRow()).getValues();///////////////AC2:AG
  return datos.filter(fila => fila.some(celda => celda !== ""));
}

function obtenerOpcionesUbicacion() {
  const hoja= SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones")
  const datos = hoja.getRange("Z2:Z" + hoja.getLastRow()).getValues();//Z2:z CAMBIO SI NO FUNK
  return datos.flat().filter((valor, i, self) => valor && self.indexOf(valor) === i);
}

function obtenerUbicacionesOcupadasPorSKU(sku) {
  try {
    const sheet= SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones")
    
    // Verificar si la hoja existe
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Ubicaciones'");
    }

    // Obtener el último registro con datos
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) { // Si solo hay encabezados o está vacía
      return [];
    }

    // Obtener todos los datos en un solo llamado
    const dataRange = sheet.getRange("AC2:AF" + lastRow);//ac2:af
    const data = dataRange.getValues();

    // Crear array para almacenar ubicaciones
    const ubicacionesOcupadas = [];

    // Procesar los datos usando forEach para mejor rendimiento
    data.forEach(row => {
      // Verificar que la fila tenga datos válidos y coincida con el SKU
      if (row[1] && row[1].toString() === sku.toString() && row[0]) {
        // Agregar solo ubicaciones únicas
        if (!ubicacionesOcupadas.includes(row[0])) {
          ubicacionesOcupadas.push(row[0]);
        }
      }
    });

    // Registrar en el log para debugging
    Logger.log(`Se encontraron ${ubicacionesOcupadas.length} ubicaciones para el SKU: ${sku}`);
    
    return ubicacionesOcupadas;

  } catch (error) {
    Logger.log(`Error en obtenerUbicacionesOcupadasPorSKU: ${error.message}`);
    return [];
  }
}
function obtenerDatosSKUSMT() {
  const sheet = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  const dataRange = sheet.getRange("BD3:BG");
  const data = dataRange.getValues();

  const datosSKU = data.map(row => ({
    ubicacion: row[0],
    sku: row[1],
    descripcion: row[2],
    cantidad: row[3],
  }));

  return datosSKU;
}
function registrarSalida(datos) {
  var hojaControlBodega = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  var hojaSalidas = SpreadsheetApp.openById(FSMT).getSheetByName("Salidas");

  // Obtener los datos del formulario
  var fecha = datos.fecha;
  var documento = datos.documento;
  var pallets = datos.pallets || 0;
  var cliente = datos.cliente;
  var filasSeleccionadas = datos.filasSeleccionadas;

  // Obtener datos de la hoja de ubicaciones
  var ubicacionesData = hojaControlBodega.getRange(2, 1, hojaControlBodega.getLastRow() - 1, 15).getValues();
  var ubicacionesUpdates = [];

  // Procesar cada fila seleccionada
  filasSeleccionadas.forEach(function (fila, index) {
    var ubicacion = fila.ubicacion;
    var sku = fila.sku;
    var descripcion = fila.descripcion;
    var cantidad = fila.cantidad;

    // Buscar la ubicación en la hoja de control de bodega
    var ubicacionIndex = ubicacionesData.findIndex(row => row[3] === ubicacion);

    if (ubicacionIndex !== -1) {
      var cantidadActual = ubicacionesData[ubicacionIndex][8];
      var nuevaCantidad = cantidadActual - cantidad;

      if (nuevaCantidad <= 0) {
        ubicacionesUpdates.push({
          row: ubicacionIndex + 2,
          clearColumns: [5, 6, 7, 8, 9, 10, 15] // Columnas E, F, G, H, I, J y O
        });
      } else {
        ubicacionesUpdates.push({
          row: ubicacionIndex + 2,
          column: 9,
          value: nuevaCantidad
        });
      }

      // Guardar los datos en la hoja de salidas
      if (index === 0) {
        hojaSalidas.appendRow([fecha, documento, sku, descripcion, cantidad, pallets, ubicacion, cliente]);
      } else {
        hojaSalidas.appendRow([fecha, documento, sku, descripcion, cantidad, 0, ubicacion, cliente]);
      }
    }
  });

  // Aplicar actualizaciones a la hoja de ubicaciones
  ubicacionesUpdates.forEach(update => {
    if (update.clearColumns) {
      update.clearColumns.forEach(col => {
        hojaControlBodega.getRange(update.row, col).clearContent();
      });
    } else {
      hojaControlBodega.getRange(update.row, update.column).setValue(update.value);
    }
  });
}
//-----------------ASDF---------
function buscarUbicacionesParaRetiro(sku, cantidadSolicitada) {
  const sheet = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,15).getValues();
  let ubicaciones = [];

  // ubicacion = col 4, sku = col 7, stock = col 9
  for (let i = 0; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().trim() === sku.toString().trim() && Number(data[i][8]) > 0) {
      ubicaciones.push({
        ubicacion: data[i][3],
        stock: Number(data[i][8]),
        descripcion: data[i][7] || ''
      });
    }
  }

  // Ordena por el último número/letra de la ubicación, de mayor a menor (a la mano)
  ubicaciones.sort((a, b) => {
    const posA = parseInt((a.ubicacion||'').split('-').pop(), 36) || 0;
    const posB = parseInt((b.ubicacion||'').split('-').pop(), 36) || 0;
    return posB - posA;
  });

  let cantidadRestante = cantidadSolicitada;
  let detalle = [];
  for (let ub of ubicaciones) {
    if (cantidadRestante <= 0) break;
    let cantidadTomar = Math.min(ub.stock, cantidadRestante);
    detalle.push({ ubicacion: ub.ubicacion, cantidad: cantidadTomar, descripcion: ub.descripcion });
    cantidadRestante -= cantidadTomar;
  }

  return {
    detalle: detalle,
    suficiente: cantidadRestante <= 0,
    cantidadFaltante: cantidadRestante > 0 ? cantidadRestante : 0,
    sku: sku
  };
}

// Guarda y descuenta la salida rápida, reusa lógica de registrarSalida

function registrarSalidaRapidaMultiple(datos) {
  var ss = SpreadsheetApp.openById(ID); // Usa tu ID de Spreadsheet
  var sheet = ss.getSheetByName("Salidas");
  var usuario = Session.getActiveUser().getEmail() || "Desconocido";
  var fecha = datos.fecha;
  var documento = datos.documento;
  var resultados = datos.resultados; // [{sku, detalle:[{ubicacion, cantidad}], ...}]
  // Ya no usamos datos.cliente

  resultados.forEach(function(res) {
    // Si tienes la descripción disponible en "res.descripcion", úsala. Si no, déjala en blanco o búscala si necesario.
    var descripcion = ""; // Puedes mejorarlo si necesitas la descripción
    // Pero normalmente en tu flujo la descripción la maneja el frontend

    res.detalle.forEach(function(ub) {
      // Ajusta el orden de las columnas según tu hoja "Salidas"
      // Ejemplo: Fecha | Documento | SKU | Descripción | Cantidad | Ubicación | Usuario
      var fila = [
        fecha,                  // Col A
        documento,              // Col B
        res.sku,                // Col C
        descripcion,            // Col D
        ub.cantidad,            // Col E
        ub.ubicacion,           // Col F
        usuario                 // Col G
        // Col H ya no se usa para Cliente
      ];
      sheet.appendRow(fila);
    });
  });

  return {mensaje: "Salida registrada correctamente."};
}

function buscarUbicacionesParaRetiroMultiple(datos) {
  // datos = array de {sku, cantidad}
  var sheet = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  var data = sheet.getDataRange().getValues();
  var resultados = [];

  datos.forEach(function(pedido) {
    var sku = pedido.sku;
    var cantidadSolicitada = pedido.cantidad;
    var detalle = [];
    var cantidadAcumulada = 0;
    var descripcion = "";

    for (var i = 1; i < data.length; i++) { // Empieza en 1 si la primera fila es encabezado
      var row = data[i];
      if (row[6] == sku && row[8] > 0) { // Ajusta índices según tu hoja
        if (descripcion === "" && row[7]) descripcion = row[7];
        var cantidadDisponible = row[8];
        var aRetirar = Math.min(cantidadSolicitada - cantidadAcumulada, cantidadDisponible);
        if (aRetirar > 0) {
          detalle.push({ubicacion: row[3], cantidad: aRetirar}); // <-- aquí va la columna D
          cantidadAcumulada += aRetirar;
        }
        if (cantidadAcumulada >= cantidadSolicitada) break;
      }
    }
    resultados.push({
      sku: sku,
      descripcion: descripcion,
      detalle: detalle,
      suficiente: cantidadAcumulada >= cantidadSolicitada,
      cantidadFaltante: Math.max(0, cantidadSolicitada - cantidadAcumulada)
    });
  });

  return {resultados: resultados};
}

function registrarSalidaRapidaMultiple(datos) {
  var hojaUbicaciones = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  var hojaSalidas = SpreadsheetApp.openById(FSMT).getSheetByName("Salidas");

  var fecha = datos.fecha;
  var documento = datos.documento;
  var resultados = datos.resultados;

  // Leer todos los datos una vez para mejor performance
  var dataUbi = hojaUbicaciones.getRange(2, 1, hojaUbicaciones.getLastRow()-1, 15).getValues();

  resultados.forEach(function(res) {
    var sku = res.sku;
    var descripcion = res.descripcion || "";
    res.detalle.forEach(function(ub) {
      var ubicacion = ub.ubicacion;
      var cantidad = ub.cantidad;

      // Buscar fila en ubicaciones (col D = ubicacion física)
      var idx = dataUbi.findIndex(row => row[3] === ubicacion && row[6] == sku);
      if (idx !== -1) {
        var filaSheet = idx + 2;
        var stockActual = dataUbi[idx][8];
        var nuevoStock = stockActual - cantidad;

        // Si queda stock, actualiza; si no, limpia columnas E:J y O
        if (nuevoStock > 0) {
          hojaUbicaciones.getRange(filaSheet, 9).setValue(nuevoStock); // Col I
        } else {
          hojaUbicaciones.getRange(filaSheet, 5, 1, 6).clearContent(); // E:J
          hojaUbicaciones.getRange(filaSheet, 15).clearContent();      // O
        }
      }

      // Guarda en hoja "Salidas" (Summit)
      hojaSalidas.appendRow([
        fecha,          // A
        documento,      // B
        sku,            // C
        descripcion,    // D
        cantidad,       // E
        0,             // F (VACÍO)
        ubicacion       // G (ubicación física como texto)
        // H en adelante queda igual o vacío
      ]);
    });
  });

  return {mensaje: "Salida registrada correctamente."};
}

function buscarSugerenciasSKUDescripcion(query) {
  // Abre la hoja de ubicaciones
  const sheet = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  // Obtén solo columnas G, H, I, O (6, 7, 8, 15) para máxima velocidad y menor memoria
  const data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 10).getValues(); // G-O es 10 columnas desde G
  query = (query || "").toString().toLowerCase();

  // Agrupa y suma stock total por SKU solo para cliente SUMMIT y stock > 0
  const mapSKU = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sku = row[0] ? row[0].toString().trim() : "";     // G (ahora data[i][0])
    const descripcion = row[1] ? row[1].toString().trim() : ""; // H (data[i][1])
    const cantidad = Number(row[2]);                        // I (data[i][2])
    const cliente = row[8] ? row[8].toString() : "";        // O (data[i][8], la novena columna de las 10 seleccionadas)

    if (cliente !== "SUMMIT" || cantidad <= 0) continue;

    // Filtra por query en SKU o descripción
    if (
      (sku && sku.toLowerCase().includes(query)) ||
      (descripcion && descripcion.toLowerCase().includes(query))
    ) {
      if (!mapSKU[sku]) {
        mapSKU[sku] = { sku: sku, descripcion: descripcion, cantidad: 0 };
      }
      mapSKU[sku].cantidad += cantidad;
    }
  }

  // Devuelve los primeros 20 resultados ordenados por SKU
  return Object.values(mapSKU)
    .sort((a, b) => a.sku.localeCompare(b.sku))
    .slice(0, 20);
}
function obtenerTodosLosSKUS() {
  // Lee la hoja de 'Ubicaciones', columnas G (SKU), H (Descripción), I (Stock), O (Cliente)
  const sheet = SpreadsheetApp.openById(CRTLBOD).getSheetByName("Ubicaciones");
  const data = sheet.getRange(2, 7, sheet.getLastRow() - 1, 9).getValues(); // G a O (7 a 15 = 9 columnas)
  let skuMap = {};

  for (let i = 0; i < data.length; i++) {
    const sku = data[i][0] ? data[i][0].toString().trim() : "";    // G
    const descripcion = data[i][1] ? data[i][1].toString().trim() : ""; // H
    const cantidad = Number(data[i][2]);                               // I
    const cliente = data[i][8] ? data[i][8].toString().trim() : "";    // O

    if (!sku || !descripcion || !cliente || cantidad <= 0) continue;
    if (cliente === "SUMMIT") {
      if (!skuMap[sku]) {
        skuMap[sku] = {
          sku: sku,
          descripcion: descripcion,
          cantidad: 0
        };
      }
      skuMap[sku].cantidad += cantidad;
    }
  }

  // Solo devuelve los SKUs con cantidad total > 0
  return Object.values(skuMap).filter(item => item.cantidad > 0);
}

//-----------------------------------------------------------------------------

// Devuelve [{sku, descripcion}] desde Maestro (Col N=13, Col O=14, sin duplicados)
function obtenerOpcionesInsumos() {
  const ss = SpreadsheetApp.openById(FSMT); // FSMT debe ser el ID de tu archivo Maestro
  const hoja = ss.getSheetByName('Maestro');
  const data = hoja.getRange(3, 14, hoja.getLastRow()-2, 2).getValues(); // N3:O
  // Filtra duplicados y vacíos
  const map = {};
  data.forEach(([sku, descripcion]) => {
    if (sku && descripcion && !map[descripcion]) {
      map[descripcion] = sku;
    }
  });
  return Object.keys(map).map(desc => ({sku: map[desc], descripcion: desc}));
}

// Guarda los datos en la hoja "INSUMOS"
function guardarDatosIngresoInsumo(datos) {
  const ss = SpreadsheetApp.openById(FSMT); // o el ID de tu archivo con la hoja INSUMOS
  const hoja = ss.getSheetByName('Insumos');
  if (!hoja) throw new Error('No se encontró la hoja INSUMOS');
  datos.productos.forEach(prod => {
    hoja.appendRow([
      datos.fecha,
      datos.guia,
      prod.sku,
      prod.descripcion,
      prod.cantidad
    ]);
  });
  return {success:true};
}

function obtenerProveedoresFSMT() {
  // ID del Google Sheet FSMT
  var ss = SpreadsheetApp.openById(FSMT);
  var hoja = ss.getSheetByName("Maestro");
  if (!hoja) return [];
  var data = hoja.getRange("Q2:Q" + hoja.getLastRow()).getValues().flat();
  // Filtrar vacíos y devolver únicos
  var proveedoresUnicos = Array.from(new Set(data)).filter(function(p){ return p && p.trim() !== ""; });
  return proveedoresUnicos;
}
