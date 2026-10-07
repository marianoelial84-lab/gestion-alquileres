// CONFIGURACIÓN GLOBAL
const ID_PLANTILLA_DOC = '1h0wgNIqsNqzvJhMV8aXhCP3Myy0BQIR2AzLEWO9yEkk';
const ID_CARPETA_DESTINO = '1MttTHfWV83L5G12COF5WDv3nzN-lKXVC';
const EMAIL_DESTINO = 'marianoelial84@gmail.com';
const SHEET_ID = '1ItAtAdla9N3YqRbT1VoXnLRRGC5so4hIbjlvSOvdTDc';
const SHEET_NAME_R = 'Reservas';
const SHEET_NAME = 'Gastos';

/**
 * Función que se ejecutará automáticamente cada hora mediante un Trigger.
 */
function verificarYEnviarContratos() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_R);
  if (!sheet) {
    Logger.log("ERROR: No se encontró la hoja con el nombre: " + SHEET_NAME_R);
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const ahora = new Date();

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];

    if (!fila[2] && !fila[5] && !fila[10]) continue;

    const cantNoches = fila[12] || '1';
    const depto = fila[2] || 'Sin Depto';
    const nombre = fila[5] || '';
    const dni = fila[6] || '';
    const tel = fila[7] || '';
    
    const datosLocatario = `${nombre} DNI: ${dni} Tel.: ${tel}`;
    const fechaIngresoRaw = fila[10];
    const fechaSalidaRaw = fila[11];
    const contratoGenerado = fila[21]; // Columna T

    if (contratoGenerado && contratoGenerado.toString().toUpperCase().includes('ENVIADO')) {
      continue;
    }

    const fechaIngreso = normalizarFecha(fechaIngresoRaw);
    const fechaSalida = normalizarFecha(fechaSalidaRaw);

    if (!fechaIngreso) continue;

    const esProximo = esFechaDeIngresoProxima(fechaIngreso, ahora);

    if (esProximo) {
      try {
        const datosReserva = {
          locatario: datosLocatario,
          depto: depto,
          cantNoches: cantNoches + (typeof cantNoches === 'number' ? " noches" : ""),
          fechaIngreso: formatearFechaCorta(fechaIngreso),
          fechaSalida: fechaSalida ? formatearFechaCorta(fechaSalida) : "-",
          fechaContratoTexto: obtenerFechaEnTexto(ahora)
        };

        // Genera el PDF, lo guarda en Drive y obtiene el Link + Blob
        const resultadoContrato = generarPDFContrato(datosReserva);

        const asunto = `Contrato de Locación - Reserva ${depto} (${datosReserva.fechaIngreso} - ${nombre})`;
        const cuerpo = `Hola,\n\nSe ha verificado un nuevo ingreso para el ${depto}.\nAdjunto encontrarás el contrato de locación correspondiente.\n\nDatos de la reserva:\n- Locatario: ${datosLocatario}\n- Departamento: ${depto}\n- Check-in: ${datosReserva.fechaIngreso}\n- Check-out: ${datosReserva.fechaSalida}\n\nEnlace al contrato en Drive: ${resultadoContrato.url}\n\nSaludos,\nApart Chivilcoy`;

        GmailApp.sendEmail(EMAIL_DESTINO, asunto, cuerpo, {
          attachments: [resultadoContrato.blob],
          name: 'Sistema Apart Chivilcoy'
        });

        // Escribir en Columna  (22): Estado de envío
        sheet.getRange(i + 1, 22).setValue('ENVIADO ' + new Date().toLocaleString());
        
        // Escribir en Columna  (23): Link directo al PDF guardado en Drive
        sheet.getRange(i + 1, 23).setValue(resultadoContrato.url);
        
        Logger.log(`-> ÉXITO: Contrato enviado y link guardado en Columna V para la fila ${i + 1}.`);

      } catch (error) {
        Logger.log(`-> ERROR al procesar fila ${i + 1}: ` + error.toString());
      }
    }
  }
}

/**
 * Crea la copia, genera el PDF, lo guarda en la carpeta de Drive
 * y retorna el Blob (para el mail) y la URL (para la Columna V).
 */
function verificarYEnviarContratos() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_R);
  if (!sheet) {
    Logger.log("ERROR: No se encontró la hoja con el nombre: " + SHEET_NAME_R);
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const ahora = new Date();

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];

    if (!fila[2] && !fila[5] && !fila[10]) continue;

    const cantNoches = fila[12] || '1';
    const depto = fila[2] || 'Sin Depto';
    const nombre = fila[5] || '';
    const dni = fila[6] || '';
    const tel = fila[7] || '';
    
    const datosLocatario = `${nombre} DNI: ${dni} Tel.: ${tel}`;
    const fechaIngresoRaw = fila[10];
    const fechaSalidaRaw = fila[11];
    const contratoGenerado = fila[21]; // Columna V (índice 21)

    if (contratoGenerado && contratoGenerado.toString().toUpperCase().includes('ENVIADO')) {
      continue;
    }

    const fechaIngreso = normalizarFecha(fechaIngresoRaw);
    const fechaSalida = normalizarFecha(fechaSalidaRaw);

    if (!fechaIngreso) continue;

    const esProximo = esFechaDeIngresoProxima(fechaIngreso, ahora);

    if (esProximo) {
      try {
        const datosReserva = {
          locatario: datosLocatario,
          depto: depto,
          cantNoches: cantNoches + (typeof cantNoches === 'number' ? " noches" : ""),
          fechaIngreso: formatearFechaCorta(fechaIngreso),
          fechaSalida: fechaSalida ? formatearFechaCorta(fechaSalida) : "-",
          fechaContratoTexto: obtenerFechaEnTexto(ahora)
        };

        // Genera el PDF, lo guarda en Drive y obtiene el objeto resultado
        const resultadoContrato = generarPDFContrato(datosReserva);

        const asunto = `Contrato de Locación - Reserva ${depto} (${datosReserva.fechaIngreso} - ${nombre})`;
        const cuerpo = `Hola,\n\nSe ha verificado un nuevo ingreso para el ${depto}.\nAdjunto encontrarás el contrato de locación correspondiente.\n\nDatos de la reserva:\n- Locatario: ${datosLocatario}\n- Departamento: ${depto}\n- Check-in: ${datosReserva.fechaIngreso}\n- Check-out: ${datosReserva.fechaSalida}\n\nEnlace al contrato en Drive: ${resultadoContrato.url}\n\nSaludos,\nApart Chivilcoy`;

        // Se pasa directamente el Blob asegurado
        GmailApp.sendEmail(EMAIL_DESTINO, asunto, cuerpo, {
          attachments: [resultadoContrato.blob],
          name: 'Sistema Apart Chivilcoy'
        });

        // Registrar en Columna (22): Estado de envío
        sheet.getRange(i + 1, 22).setValue('ENVIADO ' + new Date().toLocaleString());
        
        // Registrar en Columna (23): Link directo al PDF guardado en Drive
        sheet.getRange(i + 1, 23).setValue(resultadoContrato.url);
        
        Logger.log(`-> ÉXITO: Contrato enviado y link guardado en Columna V para la fila ${i + 1}.`);

      } catch (error) {
        Logger.log(`-> ERROR al procesar fila ${i + 1}: ` + error.toString());
      }
    }
  }
}

/**
 * Crea la copia, genera el PDF, lo guarda en la carpeta de Drive
 * y retorna el Blob válido (para el mail) y la URL (para la Columna V).
 */
function generarPDFContrato(datos) {
  const archivoPlantilla = DriveApp.getFileById(ID_PLANTILLA_DOC);
  const carpetaDestino = DriveApp.getFolderById(ID_CARPETA_DESTINO);
  
  const copiaDoc = archivoPlantilla.makeCopy(`Contrato_${datos.depto}_${datos.locatario}`, carpetaDestino);
  const doc = DocumentApp.openById(copiaDoc.getId());
  const body = doc.getBody();

  body.replaceText('{{DATOS_LOCATARIO}}', datos.locatario);
  body.replaceText('{{DEPTO}}', datos.depto);
  body.replaceText('{{CANTIDAD_NOCHES}}', datos.cantNoches);
  body.replaceText('{{FECHA_INGRESO}}', datos.fechaIngreso);
  body.replaceText('{{FECHA_SALIDA}}', datos.fechaSalida);
  body.replaceText('{{FECHA_CONTRATO_TEXTO}}', datos.fechaContratoTexto);

  doc.saveAndClose();

  const tempBlob = copiaDoc.getAs('application/pdf');
  tempBlob.setName(`Contrato_Locacion_${datos.depto}_${datos.fechaIngreso.replace(/\//g, '-')}.pdf`);

  // Guardar el archivo PDF en Drive
  const archivoPdfGuardado = carpetaDestino.createFile(tempBlob);
  Logger.log("Archivo PDF creado exitosamente en Drive: " + archivoPdfGuardado.getName());

  // Eliminar la copia de edición en Google Docs
  copiaDoc.setTrashed(true);

  // Retornar el blob extraído del archivo oficial guardado
  return {
    blob: archivoPdfGuardado.getBlob(),
    url: archivoPdfGuardado.getUrl()
  };
}

// FUNCIONES AUXILIARES DE FECHA
function normalizarFecha(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  
  if (typeof val === 'string') {
    const partes = val.split(/[\/\-]/);
    if (partes.length === 3) {
      if (partes[0].length === 4) { // Formato YYYY-MM-DD
        return new Date(partes[0], partes[1] - 1, partes[2]);
      } else { // Formato DD/MM/YYYY
        return new Date(partes[2], partes[1] - 1, partes[0]);
      }
    }
  }
  return new Date(val);
}

function esFechaDeIngresoProxima(fechaIngreso, fechaActual) {
  if (!(fechaIngreso instanceof Date) || isNaN(fechaIngreso.getTime())) return false;
  
  // Normalizar a medianoche (00:00:00) para comparar únicamente la fecha
  const fIngreso = new Date(fechaIngreso.getFullYear(), fechaIngreso.getMonth(), fechaIngreso.getDate());
  const fActual = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), fechaActual.getDate());
  
  const diffDias = (fIngreso - fActual) / (1000 * 60 * 60 * 24);
  
  // Incluye reservas para HOY (0 días) o para los próximos 2 días
  return diffDias >= 0 && diffDias <= 2;
}

function formatearFechaCorta(fecha) {
  const d = new Date(fecha);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const anio = d.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

function obtenerFechaEnTexto(fecha) {
  const meses = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];
  const d = new Date(fecha);
  const dia = d.getDate();
  const mes = meses[d.getMonth()];
  const anio = d.getFullYear();
  
  return `${dia} de ${mes} del año ${anio}`;
}