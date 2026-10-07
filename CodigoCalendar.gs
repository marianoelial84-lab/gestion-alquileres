/**
 * Revisa la hoja 'Reservas' y crea un evento en Google Calendar para cada Check-in.
 * Notifica 1 día antes a las 10:00 AM y el mismo día a las 09:00 AM.
 * Usa la Columna X (24) para registrar el ID del evento de Calendar.
 */
function sincronizarReservasCalendar() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetReservas = ss.getSheetByName(SHEET_NAME_R);

  if (!sheetReservas) {
    Logger.log("ERROR: No se encontró la hoja 'Reservas'.");
    return;
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const datos = sheetReservas.getDataRange().getValues();
  if (datos.length <= 1) return;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const idReserva = fila[0];            // Columna A: ID / Ref
    const canal = fila[9] || "Particular";    // Columna J: Canal
    const depto = fila[2] || "Sin Depto";  // Columna C: Depto
    const nombre = fila[5] || "Huésped "+canal;   // Columna F: Nombre
    const tel = fila[7] || "";             // Columna H: Teléfono
    const fechaIngresoRaw = fila[10];     // Columna K: Check-in
    const fechaSalidaRaw = fila[11];      // Columna L: Check-out
    const idExistente = fila[23];         // Columna X (Índice 23 / Columna 24): ID Evento Calendar

    // Parsear fechas
    const fechaIngreso = (fechaIngresoRaw instanceof Date) ? fechaIngresoRaw : (fechaIngresoRaw ? new Date(fechaIngresoRaw) : null);
    const fechaSalida = (fechaSalidaRaw instanceof Date) ? fechaSalidaRaw : (fechaSalidaRaw ? new Date(fechaSalidaRaw) : null);

    // Verificar que la reserva sea válida y no haya sido agendada previamente
    if (fechaIngreso && !isNaN(fechaIngreso) && (!idExistente || !String(idExistente).startsWith("CAL_RES_"))) {

      // Formatear fechas para el detalle
      const checkInStr = fechaIngreso.toLocaleDateString('es-AR');
      const checkOutStr = fechaSalida ? fechaSalida.toLocaleDateString('es-AR') : '-';

      // Título del Evento
      const titulo = `Reserva ${depto} - ${nombre}`;

      // Descripción del evento
      const descripcion = `Detalles de la Reserva:\n` +
                          `• Depto: ${depto}\n` +
                          `• Huésped: ${nombre}\n` +
                          `• Teléfono: ${tel}\n` +
                          `• Canal: ${canal}\n` +
                          `• Check-in: ${checkInStr}\n` +
                          `• Check-out: ${checkOutStr}\n` +
                          `• ID Reserva: ${idReserva}`;

      // Crear evento de todo el día para la fecha de ingreso (Check-in)
      const evento = calendar.createAllDayEvent(titulo, fechaIngreso, {
        description: descripcion
      });

      evento.removeAllReminders();

      // 1. Notificación 1 día antes a las 10:00 AM
      const fechaNotif1 = new Date(fechaIngreso);
      fechaNotif1.setDate(fechaNotif1.getDate() - 1);
      fechaNotif1.setHours(10, 0, 0, 0);

      const minAntes1 = Math.round((fechaIngreso.getTime() - fechaNotif1.getTime()) / (1000 * 60));
      if (minAntes1 > 0) {
        evento.addPopupReminder(minAntes1);
      }

      // 2. Notificación el mismo día a las 09:00 AM
      const fechaNotif2 = new Date(fechaIngreso);
      fechaNotif2.setHours(10, 0, 0, 0);

      const minAntes2 = Math.round((fechaIngreso.getTime() - fechaNotif2.getTime()) / (1000 * 60));
      if (minAntes2 > 0) {
        evento.addPopupReminder(minAntes2);
      }

      // Guardar el ID del evento en la Columna X (Columna 24 / Range fila i+1, col 24)
      const idCal = "CAL_RES_" + evento.getId();
      sheetReservas.getRange(i + 1, 24).setValue(idCal);

      Logger.log(`Reserva agendada en Calendar: ${titulo} para el ${checkInStr}`);
    }
  }
}


/**
 * Revisa la hoja 'Gastos' y crea un evento en Google Calendar SOLAMENTE para:
 * Luz, Gas, Agua e Internet & Cable.
 * Notifica 1 día antes a las 10:00 AM.
 */
function sincronizarVencimientosCalendar() {
  const sheetGastos = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheetGastos) {
    Logger.log("ERROR: No se encontró la hoja 'Gastos'.");
    return;
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const datos = sheetGastos.getDataRange().getValues();
  if (datos.length <= 1) return;

  // Categorías permitidas para agendar en Calendar
  const categoriasPermitidas = ["Luz", "Gas", "Agua", "Internet & Cable"];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    
    const fechaVencimiento = fila[0]; // Columna A: Fecha
    const depto = fila[1];            // Columna B: Depto
    const categoria = String(fila[2] || "").trim(); // Columna C: Categoría
    const detalle = fila[3];          // Columna D: Detalle
    const monto = fila[4];            // Columna E: Monto $
    const codigo = fila[5];           // Columna F: Código
    const estado = fila[6];           // Columna G: Estado
    const idExistente = fila[7];      // Columna H: ID Único / ID Evento

    // 1. Filtrar solo por las categorías permitidas
    const esServicioPermitido = categoriasPermitidas.some(
      cat => cat.toLowerCase() === categoria.toLowerCase()
    );

    // 2. Verificar que sea fecha válida, categoría permitida y que no se haya agendado previamente
    if (esServicioPermitido && fechaVencimiento instanceof Date && !isNaN(fechaVencimiento) && (!idExistente || !String(idExistente).startsWith("CAL_"))) {
      
      const titulo = `${categoria} - $${Number(monto).toLocaleString('es-AR', { minimumFractionDigits: 2 })}`;

      const descripcion = `Vencimiento de servicio:\n` +
                          `• Depto: ${depto}\n` +
                          `• Detalle: ${detalle}\n` +
                          `• Código / Ref: ${codigo}\n` +
                          `• Estado: ${estado}`;

      const evento = calendar.createAllDayEvent(titulo, fechaVencimiento, {
        description: descripcion
      });

      // Notificación 1 día antes a las 10:00 AM
      const fechaNotificacion = new Date(fechaVencimiento);
      fechaNotificacion.setDate(fechaNotificacion.getDate() - 1);
      fechaNotificacion.setHours(10, 0, 0, 0);

      const minutosAntes = Math.round((fechaVencimiento.getTime() - fechaNotificacion.getTime()) / (1000 * 60));

      if (minutosAntes > 0) {
        evento.removeAllReminders();
        evento.addPopupReminder(minutosAntes);
      }

      // Guardar ID en la Columna H
      const idCal = "CAL_" + evento.getId();
      sheetGastos.getRange(i + 1, 8).setValue(idCal);

      Logger.log(`Evento agendado: ${titulo} (${fechaVencimiento.toLocaleDateString()})`);
    }
  }
}