function cargarReservasBooking() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_R);
  if (!sheet) return;

  // Buscar correos de confirmación de Booking que aún no hayan sido cargados
  var query = 'from:noreply@booking.com subject:"¡Nueva reserva!" -label:Booking_Cargado';
  var threads = GmailApp.search(query);
  
  // Crear o buscar la etiqueta "Booking_Cargado" para no duplicar reservas
  var label = GmailApp.getUserLabelByName("Booking_Cargado");
  if (!label) {
    label = GmailApp.createLabel("Booking_Cargado");
  }

  var meses = {
    'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
    'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
    'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
  };

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var subject = message.getSubject();
      
      // Extrae ID de reserva y fecha de llegada desde el Asunto
      // Ej: "Booking.com - ¡Nueva reserva! (6247935461, sábado, 26 de septiembre de 2026)"
      var idMatch = subject.match(/\((\d+),/);
      var dateMatch = subject.match(/,\s*([^)]+)\)/);

      if (idMatch && dateMatch) {
        var idReserva = idMatch[1]; // 6247935461
        var rawFechaLlegada = dateMatch[1]; // sábado, 26 de septiembre de 2026

        // Convertir la fecha de llegada a formato DD/MM/YYYY
        var fechaLlegadaMatch = rawFechaLlegada.match(/(\d{1,2})\s+de\s+([a-zA-Z]+)\s+de\s+(\d{4})/i);
        var fechaLlegadaFormatted = "";
        if (fechaLlegadaMatch) {
          var dia = ("0" + fechaLlegadaMatch[1]).slice(-2);
          var mesTexto = fechaLlegadaMatch[2].toLowerCase();
          var mes = meses[mesTexto] || "01";
          var anio = fechaLlegadaMatch[3];
          fechaLlegadaFormatted = dia + "/" + mes + "/" + anio;
        }

        // Obtener la fecha de recepción del correo como Fecha de Reserva (DD/MM/YYYY)
        var fechaEmail = message.getDate();
        var diaRes = ("0" + fechaEmail.getDate()).slice(-2);
        var mesRes = ("0" + (fechaEmail.getMonth() + 1)).slice(-2);
        var anioRes = fechaEmail.getFullYear();
        var fechaReservaFormatted = diaRes + "/" + mesRes + "/" + anioRes;

        // Encontrar la primera fila vacía en la columna A (ID/Código)
        var values = sheet.getRange("A:A").getValues();
        var lastRow = 1;
        for (var i = 1; i < values.length; i++) {
          if (values[i][0] === "" || values[i][0] === null) {
            lastRow = i + 1;
            break;
          }
        }
        if (lastRow === 1 && values[0][0] !== "") {
          lastRow = values.length + 1;
        }

        // Cargar los datos extraídos en sus respectivas columnas
        sheet.getRange(lastRow, 1).setValue(idReserva);             // Col A: ID/Código
        sheet.getRange(lastRow, 2).setValue(fechaReservaFormatted);  // Col B: Fecha Reserva
        sheet.getRange(lastRow, 10).setValue("Booking");            // Col J: Canal
        sheet.getRange(lastRow, 11).setValue(fechaLlegadaFormatted); // Col K: Llegada
        sheet.getRange(lastRow, 20).setValue("Pendiente");          // Col R: Estado Pago
      }
    });
    
    // Asignar etiqueta en Gmail para evitar reprocesar este correo en el futuro
    thread.addLabel(label);
  });

}