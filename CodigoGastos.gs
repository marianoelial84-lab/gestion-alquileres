
/**
 * Extrae comisiones > 0 de la hoja 'Reservas' y las registra en 'Gastos'.
 * Usa Fecha de Salida (Col L), Depto (Col C) y asigna Estado:
 * - "Pagado" si el canal es Airbnb
 * - "Pendiente" si es otro canal no particular
 */
function procesarComisionesReservas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetReservas = ss.getSheetByName(SHEET_NAME_R);
  const sheetGastos = ss.getSheetByName(SHEET_NAME);

  if (!sheetReservas || !sheetGastos) {
    Logger.log("ERROR: Verifica que existan las hojas 'Reservas' y 'Gastos'.");
    return;
  }

  // 1. Obtener los IDs ya existentes en la hoja Gastos (Columna H / Indice 7) para evitar duplicados
  const datosGastos = sheetGastos.getDataRange().getValues();
  const idsExistentes = new Set();
  
  if (datosGastos.length > 1) {
    for (let i = 1; i < datosGastos.length; i++) {
      const id = datosGastos[i][7]; // Columna H (ID Oculto/Único)
      if (id) idsExistentes.add(String(id));
    }
  }

  // 2. Obtener datos de la hoja Reservas
  const datosReservas = sheetReservas.getDataRange().getValues();
  if (datosReservas.length <= 1) return;

  let nuevosGastos = 0;

  for (let f = 1; f < datosReservas.length; f++) {
    const fila = datosReservas[f];
    
    // Mapeo según la estructura indicada:
    const idReserva = fila[0];                     // Columna A: ID o Código de Reserva
    const depto = fila[2] || "General";             // Columna C: Depto
    const fechaSalida = fila[11];                   // Columna L: Fecha de Salida
    const comision = Number(fila[17]) || 0;        // Columna P: Comisión
    const canal = String(fila[9] || "").toLowerCase(); // Columna J: Canal

    // Procesamos solo comisiones mayores a 0
    if (comision > 0) {
      const idComision = "COM_" + String(idReserva);

      // Verificar que no se haya cargado anteriormente
      if (!idsExistentes.has(idComision)) {
        
        // Manejo de la fecha de salida (Columna L)
        let fechaObj = new Date();
        if (fechaSalida instanceof Date) {
          fechaObj = fechaSalida;
        } else if (fechaSalida) {
          fechaObj = new Date(fechaSalida);
        }

        // Determinar el estado según el canal
        const estadoGasto = canal.includes("airbnb") ? "Pagado" : "Pendiente";

        // Estructura en Gastos: [Fecha, Depto, Categoria, Detalle, Monto, Codigo, Estado, ID]
        sheetGastos.appendRow([
          fechaObj,
          depto,
          "Comisiones Plataformas",
          "Comisión " + (fila[9] || "Canal") + " - Reserva " + idReserva,
          comision,
          String(idReserva),
          estadoGasto,
          idComision
        ]);

        idsExistentes.add(idComision);
        nuevosGastos++;
      }
    }
  }

  Logger.log(`Proceso finalizado. Se agregaron ${nuevosGastos} comisiones a la hoja de Gastos.`);
}

/**
 * Procesa los correos de los últimos 30 días reenviados desde las cuentas indicadas,
 * extrae los datos de servicios (Luz, Gas, Agua, Internet), los registra en 'gastos'
 * y los etiqueta con 'Servicio_Cargado'.
 */
function procesarMailsServicios() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("ERROR: No se encontró la hoja con el nombre: " + SHEET_NAME);
    return;
  }

  // Crear la etiqueta 'Servicio_Cargado' en Gmail si aún no existe
  let etiqueta = GmailApp.getUserLabelByName("Servicio_Cargado");
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel("Servicio_Cargado");
  }

  // Búsqueda en Gmail de correos de los últimos 30 días sin la etiqueta aplicada
  const query = 'newer_than:30d -label:Servicio_Cargado (EDEN OR Camuzzi OR ABSA OR "AGUAS BONAERENSES" OR "Factura Personal" OR "saldo total a debitar")';
  const hilos = GmailApp.search(query);

  Logger.log(`Se encontraron ${hilos.length} hilos de correo para revisar.`);

  hilos.forEach(hilo => {
    const mensajes = hilo.getMessages();
    
    mensajes.forEach(mensaje => {
      const cuerpo = mensaje.getPlainBody();
      const asunto = mensaje.getSubject();
      const textoCompleto = (asunto + " " + cuerpo).toLowerCase();

      let gasto = null;

      // 1. LUZ (EDEN)
      if (textoCompleto.includes("edensa.com.ar") || textoCompleto.includes("NIS 163761801")) {
        gasto = extraerDatosLuz(cuerpo, asunto);
      } 
      // 2. GAS (CAMUZZI)
      else if (textoCompleto.includes("camuzzi") || textoCompleto.includes("6620/0-1806-0037808/2")) {
        gasto = extraerDatosGas(cuerpo, asunto);
      } 
      // 3. AGUA (ABSA)
      else if (textoCompleto.includes("absa") || textoCompleto.includes("aguas bonaerenses") || textoCompleto.includes("1026879")) {
        gasto = extraerDatosAgua(cuerpo, asunto);
      } 
      // 4. INTERNET & CABLE (PERSONAL / FLOW)
      else if (textoCompleto.includes("personal") || textoCompleto.includes("flow") || textoCompleto.includes("1003357656510001")) {
        gasto = extraerDatosInternet(cuerpo, asunto);
      }

      if (gasto) {
        const idUnico = new Date().getTime().toString() + "_" + Math.floor(Math.random() * 1000);
        
        // Formato de la hoja: Fecha(Vencimiento), Depto, Categoría, Detalle, Monto, Código, Estado, ID
        sheet.appendRow([
          gasto.fechaVencimiento, 
          "General", 
          gasto.categoria, 
          gasto.detalle, 
          gasto.monto, 
          gasto.codigo, 
          "DA", 
          idUnico
        ]);

        Logger.log(`Gasto guardado: ${gasto.categoria} - $${gasto.monto} - Vence: ${gasto.fechaVencimiento}`);
      }
    });

    // Aplica la etiqueta para no volver a cargarlo en la siguiente ejecución
    hilo.addLabel(etiqueta);
  });
}

// --- FUNCIONES EXTRACTORAS ---

function extraerDatosLuz(cuerpo, asunto) {
//  const matchFecha = cuerpo.match(/(?:1er Vencimiento|Vencimiento)\s*[\r\n]*\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
// 1. Intenta capturar "1er Vencimiento" seguido de la fecha (incluso con saltos de línea HTML)
  let matchFecha = cuerpo.match(/1er\s*Vencimiento[\s\S]*?(\d{1,2}\/\d{1,2}\/\d{4})/i);
  
  // 2. Si no encuentra "1er Vencimiento", busca la primera fecha válida DD/MM/YYYY en todo el cuerpo
  if (!matchFecha) {
    matchFecha = cuerpo.match(/(\d{1,2}\/\d{1,2}\/20\d{2})/);
  }

  const matchMonto = cuerpo.match(/\$\s*([\d\.,]+)/);

  return {
    categoria: "Luz",
    codigo: "163761801",
    detalle: "EDEN S.A. - Factura Digital",
    fechaVencimiento: matchFecha ? parsearFecha(matchFecha[1]) : new Date(),
    monto: matchMonto ? limpiarMonto(matchMonto[1]) : 0
  };
}

function extraerDatosGas(cuerpo, asunto) {
// 1. Busca "Vencimiento:" seguido de la fecha en el cuerpo
  let matchFecha = cuerpo.match(/Vencimiento:\s*[\r\n]*\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);

  // 2. Si no la encuentra, busca la primera fecha DD/MM/YYYY que aparezca en el cuerpo o asunto
  if (!matchFecha) {
    matchFecha = (asunto + " " + cuerpo).match(/(\d{1,2}\/\d{1,2}\/20\d{2})/);
  }

  const matchMonto = cuerpo.match(/Total:\s*[\r\n]*\$\s*([\d\.,]+)/i) || cuerpo.match(/\$\s*([\d\.,]+)/);

  return {
    categoria: "Gas",
    codigo: "6620/0-1806-0037808/2",
    detalle: "Camuzzi Gas",
    fechaVencimiento: matchFecha ? parsearFecha(matchFecha[1]) : new Date(),
    monto: matchMonto ? limpiarMonto(matchMonto[1]) : 0
  };
}

function extraerDatosAgua(cuerpo, asunto) {
  const matchFecha = cuerpo.match(/Vencimiento\s*\|\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  const matchMonto = cuerpo.match(/Importe\s*\|\s*\$?\s*([\d\.,]+)/i);

  return {
    categoria: "Agua",
    codigo: "1026879",
    detalle: "AGUAS BONAERENSES S.A.",
    fechaVencimiento: matchFecha ? parsearFecha(matchFecha[1]) : new Date(),
    monto: matchMonto ? limpiarMonto(matchMonto[1]) : 0
  };
}

function extraerDatosInternet(cuerpo, asunto) {
  const matchFecha = cuerpo.match(/Vencimiento\s*[\r\n]*\s*(\d{1,2}\/\d{1,2}\/\d{4})/i) || asunto.match(/vence el (\d{1,2}\/\d{1,2}\/\d{4})/i);
  const matchMonto = cuerpo.match(/Total a debitar\s*[\r\n]*\$\s*([\d\.,]+)/i) || asunto.match(/\$\s*([\d\.,]+)/);

  return {
    categoria: "Internet & Cable",
    codigo: "1003357656510001",
    detalle: "Personal - Fibra + Flow",
    fechaVencimiento: matchFecha ? parsearFecha(matchFecha[1]) : new Date(),
    monto: matchMonto ? limpiarMonto(matchMonto[1]) : 0
  };
}

// --- FUNCIONES AUXILIARES ---

function parsearFecha(cadenaFecha) {
  const partes = cadenaFecha.split('/');
  if (partes.length === 3) {
    return new Date(partes[2], partes[1] - 1, partes[0]);
  }
  return new Date();
}

function limpiarMonto(textoMonto) {
  let limpio = textoMonto.trim();
  if (limpio.includes(',') && limpio.includes('.')) {
    limpio = limpio.replace(/\./g, '').replace(',', '.');
  } else if (limpio.includes(',')) {
    limpio = limpio.replace(',', '.');
  }
  return parseFloat(limpio) || 0;
}

// ***********************************************************************************
//                     GASTOS WEB
// ***********************************************************************************

// Servir la aplicación web
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestión de Departamentos - Dashboard & Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Incluir archivos HTML auxiliares si los usaras
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene y procesa todas las métricas e indicadores desde las hojas 'Reservas' y 'Gastos'
 */
function obtenerDatosDashboard() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetReservas = ss.getSheetByName(SHEET_NAME_R);
  const sheetGastos = ss.getSheetByName(SHEET_NAME);

  const datosReservas = sheetReservas ? sheetReservas.getDataRange().getValues() : [];
  const datosGastos = sheetGastos ? sheetGastos.getDataRange().getValues() : [];

  // 1. Procesar Reservas
  // Estructura: A:ID, B:Canal, C:Depto, K:Check-in, L:Check-out, O:Ingreso Noche/Total, P:Comision...
  // Ajustar según columnas exactas:
  let reservas = [];
  if (datosReservas.length > 1) {
    for (let i = 1; i < datosReservas.length; i++) {
      const fila = datosReservas[i];
      if (!fila[0]) continue; // Saltar filas vacías

      const checkIn = fila[10] ? new Date(fila[10]) : null;
      const checkOut = fila[11] ? new Date(fila[11]) : null;
      
      // Cálculo de noches
      let noches = 0;
      if (checkIn && checkOut && !isNaN(checkIn) && !isNaN(checkOut)) {
        noches = Math.max(1, Math.round((checkOut - checkIn) / (1000 * 60 * 60 * 24)));
      }

      // Supongamos: Col B=Canal, Col C=Depto, Col M o N=Monto Bruto Total
      const canal = String(fila[9] || "Particular").trim();
      const depto = String(fila[2] || "General").trim();
      const montoIngreso = Number(fila[16] || fila[18] || 0); // Ajustar a columna de Ingreso Total

      reservas.push({
        id: fila[0],
        canal: canal,
        depto: depto,
        checkIn: checkIn ? checkIn.toISOString() : null,
        checkOut: checkOut ? checkOut.toISOString() : null,
        noches: noches,
        monto: montoIngreso
      });
    }
  }

  // 2. Procesar Gastos
  // Estructura: A:Fecha, B:Depto, C:Categoria, D:Detalle, E:Monto
  let gastos = [];
  if (datosGastos.length > 1) {
    for (let i = 1; i < datosGastos.length; i++) {
      const fila = datosGastos[i];
      if (!fila[0]) continue;

      const fecha = fila[0] ? new Date(fila[0]) : null;
      gastos.push({
        fecha: fecha ? fecha.toISOString() : null,
        depto: String(fila[1] || "General").trim(),
        categoria: String(fila[2] || "Varios").trim(),
        detalle: fila[3] || "",
        monto: Number(fila[4]) || 0
      });
    }
  }

  return {
    reservas: reservas,
    gastos: gastos,
    deptosDisponibles: ["Depto 2", "Depto 3", "Depto 4", "Depto 5", "Depto 6"] // Deptos activos
  };
}

function obtenerGastos() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Solo está el encabezado
  data.shift(); // Quita encabezados

  // Formateamos las filas para asegurar que la fecha sea texto usable y los tipos de datos sean limpios
  return data.map(fila => {
    let fechaTexto = '';
    if (fila[0] instanceof Date) {
      const dia = String(fila[0].getDate()).padStart(2, '0');
      const mes = String(fila[0].getMonth() + 1).padStart(2, '0');
      const anio = fila[0].getFullYear();
      fechaTexto = `${anio}-${mes}-${dia}`; // Formato ISO estandar
    } else {
      fechaTexto = fila[0] ? String(fila[0]) : '';
    }

    return [
      fechaTexto,               // Col A: Fecha
      String(fila[1] || ''),    // Col B: Depto
      String(fila[2] || ''),    // Col C: Categoría
      String(fila[3] || ''),    // Col D: Detalle
      Number(fila[4]) || 0,     // Col E: Monto $
      String(fila[5] || '')     // Col F: ID Oculto
    ];
  });
}

function cargarGasto(datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const id = new Date().getTime().toString();
  
  let fechaObjeto = "";
  if (datos.fecha) {
    const partes = datos.fecha.split('-');
    fechaObjeto = new Date(partes[0], partes[1] - 1, partes[2]);
  }

  const montoNumero = parseFloat(datos.monto) || 0;

  // Inserta la fila completa
  sheet.appendRow([fechaObjeto, datos.depto, datos.categoria, datos.detalle, montoNumero, id]);
  
  return obtenerGastos();
}

function eliminarGasto(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] == id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return obtenerGastos();
}

function modificarGasto(datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  let fechaObjeto = "";
  if (datos.fecha) {
    const partes = datos.fecha.split('-');
    fechaObjeto = new Date(partes[0], partes[1] - 1, partes[2]);
  }
  const montoNumero = parseFloat(datos.monto) || 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][5] == datos.id) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[fechaObjeto, datos.depto, datos.categoria, datos.detalle, montoNumero]]);
      break;
    }
  }
  return obtenerGastos();
}