function cargarReservasAirbnb() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_R);
  if (!sheet) return;

  // 1. Cargar IDs existentes para evitar duplicados
  var existingIds = sheet.getRange("A2:A" + Math.max(sheet.getLastRow(), 2))
                        .getValues()
                        .flat()
                        .map(function(id) { return String(id).trim().toLowerCase(); });

  // 2. Gestionar etiqueta anti-duplicados
  var labelName = "Airbnb-Cargado";
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  // 3. Buscar SOLO correos recientes (últimos 30 días) que contengan "Reserva" y "Airbnb"
  var query = 'subject:"Reserva" "Airbnb" newer_than:30d -label:' + labelName;
  var threads = GmailApp.search(query);

  var meses = {
    'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
    'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
    'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
  };

  threads.forEach(function(thread) {
    var threadLabels = thread.getLabels().map(function(l) { return l.getName(); });
    if (threadLabels.indexOf(labelName) !== -1) return;

    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var body = message.getPlainBody();
      
      // Fecha de recepción del correo (Fecha de Reserva)
      var fechaEmail = message.getDate();
      var diaRes = ("0" + fechaEmail.getDate()).slice(-2);
      var mesRes = ("0" + (fechaEmail.getMonth() + 1)).slice(-2);
      var anioRes = fechaEmail.getFullYear();
      var fechaReservaFormatted = diaRes + "/" + mesRes + "/" + anioRes;

      // Extraer Nombre del Huésped
      var nombreHuesped = "";
      var matchNombre = body.match(/([^\n\r]+)[\n\r]+\s*Persona que reserva/i);
      if (matchNombre) {
        nombreHuesped = matchNombre[1].trim();
      }

      // Extraer Fecha de Llegada
      var fechaLlegadaFormatted = "";
      var llegadaObj = null;
      var matchLlegada = body.match(/Llegada[\s\S]*?(\d{1,2})\s+de\s+([a-zA-Z]+)\s+de\s+(\d{4})/i);
      if (matchLlegada) {
        var diaL = ("0" + matchLlegada[1]).slice(-2);
        var mesL = meses[matchLlegada[2].toLowerCase()] || "01";
        var anioL = matchLlegada[3];
        fechaLlegadaFormatted = diaL + "/" + mesL + "/" + anioL;
        llegadaObj = new Date(anioL, parseInt(mesL) - 1, parseInt(diaL));
      }

      // SI NO ES UNA RESERVA VÁLIDA (Mails viejos o promocionales), SALTEAR
      if (!fechaLlegadaFormatted && !nombreHuesped) {
        return;
      }

      // Extraer Fecha de Salida
      var fechaSalidaFormatted = "";
      var salidaObj = null;
      var matchSalida = body.match(/Salida[\s\S]*?(\d{1,2})\s+de\s+([a-zA-Z]+)\s+de\s+(\d{4})/i);
      if (matchSalida) {
        var diaS = ("0" + matchSalida[1]).slice(-2);
        var mesS = meses[matchSalida[2].toLowerCase()] || "01";
        var anioS = matchSalida[3];
        fechaSalidaFormatted = diaS + "/" + mesS + "/" + anioS;
        salidaObj = new Date(anioS, parseInt(mesS) - 1, parseInt(diaS));
      }

      // Generar ID único inteligente (airbnb_AAAAMMDD_nombre)
      var nameClean = nombreHuesped.toLowerCase().replace(/[^a-z0-9]/g, "");
      var fechaId = llegadaObj ? (anioL + mesL + diaL) : (anioRes + mesRes + diaRes);
      var idReserva = "airbnb_" + fechaId + (nameClean ? "_" + nameClean : "");

      // Si el ID ya existe en la planilla, ignorar
      if (existingIds.indexOf(idReserva.toLowerCase()) !== -1) {
        return; 
      }

      // Cálculo de Noches
      var noches = "";
      if (llegadaObj && salidaObj) {
        var diffTime = Math.abs(salidaObj - llegadaObj);
        noches = Math.round(diffTime / (1000 * 60 * 60 * 24));
      }

      // Cantidad de Pax
      var pax = "";
      var matchPax = body.match(/Viajeros[\s\S]*?(\d+)\s*(?:adulto|viajero|huésped|persona)/i);
      if (matchPax) {
        pax = parseInt(matchPax[1]);
      }

      // Buscar la primera fila vacía
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

      // Escribir datos
      sheet.getRange(lastRow, 1).setValue(idReserva);             // Col A: ID
      sheet.getRange(lastRow, 2).setValue(fechaReservaFormatted);  // Col B: Fecha Reserva
      sheet.getRange(lastRow, 6).setValue(nombreHuesped);          // Col F: Nombre Huésped
      sheet.getRange(lastRow, 9).setValue(pax);                    // Col I: Pax
      sheet.getRange(lastRow, 10).setValue("Airbnb");              // Col J: Canal
      sheet.getRange(lastRow, 11).setValue(fechaLlegadaFormatted); // Col K: Llegada
      sheet.getRange(lastRow, 12).setValue(fechaSalidaFormatted);  // Col L: Salida
      sheet.getRange(lastRow, 13).setValue(noches);                // Col M: Noches
      sheet.getRange(lastRow, 20).setValue("Pagado");              // Col T: Estado Pago
      sheet.getRange(lastRow, 21).setValue("Santander");           // Col U: Forma Pago

      existingIds.push(idReserva.toLowerCase());
    });

    thread.addLabel(label);
  });
}