/**
 * Google Apps Script Backend for Tech Crew Scheduling
 * Deployed as a Web App to handle:
 * 1. Contact form submissions (saves to sheet)
 * 2. Schedule booking (saves to calendar, sends invites, saves to sheet)
 * 3. Calendar availability checking (lists busy slot intervals on a specific date)
 */

function doPost(e) {
  var response = {};
  try {
    var data = JSON.parse(e.postData.contents);
    
    if (data.tipo_formulario === 'Obtener Disponibilidad') {
      var busySlots = getBusyEventsForDate(data.fecha);
      response = {
        status: 'success',
        busySlots: busySlots
      };
    } else if (data.tipo_formulario === 'Agendar Llamada') {
      // 1. Create event in calendar
      var event = createCalendarEvent(data);
      
      // 2. Send custom HTML confirmation emails to client and crew
      sendEmails(data);
      
      // 3. Append booking information to spreadsheet (if bound to spreadsheet)
      saveToSheet(data, 'Llamadas Agendadas');
      
      response = {
        status: 'success',
        eventId: event ? event.getId() : null
      };
    } else {
      // Formulario de Contacto u otros
      saveToSheet(data, 'Contacto');
      response = {
        status: 'success'
      };
    }
  } catch (err) {
    response = {
      status: 'error',
      message: err.toString()
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gets start and end times of all events on a specific date string (YYYY-MM-DD)
 * to let the frontend filter out busy slots.
 * Queries a 3-day window to prevent timezone mismatches.
 */
function getBusyEventsForDate(dateStr) {
  var parts = dateStr.split("-");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  var day = parseInt(parts[2], 10);
  
  // Rango ampliado de 3 días (un día antes y un día después)
  // Esto evita desfases de zonas horarias entre los servidores de Google y el navegador del cliente
  var start = new Date(year, month, day - 1, 0, 0, 0);
  var end = new Date(year, month, day + 1, 23, 59, 59);
  
  // Buscar específicamente el calendario de techcrewcr@gmail.com
  var calendar = CalendarApp.getCalendarById("techcrewcr@gmail.com");
  if (!calendar) {
    calendar = CalendarApp.getDefaultCalendar();
  }
  
  var events = calendar.getEvents(start, end);
  var busySlots = [];
  
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    // Skip all-day events since they usually represent reminders/holidays, not busy time blocks
    if (event.isAllDayEvent()) {
      continue;
    }
    busySlots.push({
      start: event.getStartTime().toISOString(),
      end: event.getEndTime().toISOString()
    });
  }
  
  return busySlots;
}

/**
 * Creates Google Calendar event with guest invitations
 */
function createCalendarEvent(data) {
  // Buscar específicamente el calendario de techcrewcr@gmail.com
  var calendar = CalendarApp.getCalendarById("techcrewcr@gmail.com");
  if (!calendar) {
    calendar = CalendarApp.getDefaultCalendar();
  }
  
  var title = data.evento_titulo || 'Llamada de Consultoría - Tech Crew';
  var start = new Date(data.fecha_inicio_iso);
  var end = new Date(data.fecha_fin_iso);
  
  var options = {
    description: data.evento_descripcion || 'Llamada estratégica',
    sendInvites: true
  };
  
  if (data.cliente_email) {
    options.guests = data.cliente_email;
  }
  
  return calendar.createEvent(title, start, end, options);
}

/**
 * Appends data row to active Spreadsheet sheet
 */
function saveToSheet(data, sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return; // Silent return if script is run stand-alone without sheet
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = ['Fecha de Registro', 'Nombre', 'Email', 'Telefono', 'Detalles'];
      sheet.appendRow(headers);
    }
    
    var row = [
      new Date(),
      data.nombre || '',
      data.email || data.cliente_email || '',
      data.telefono || '',
      JSON.stringify(data)
    ];
    sheet.appendRow(row);
  } catch (e) {
    Logger.log('Spreadsheet error: ' + e.toString());
  }
}

/**
 * Sends HTML confirmation emails with calendar icons to both client and crew
 */
function sendEmails(data) {
  var clientEmail = data.email || data.cliente_email;
  var crewEmail = data.correo_crew || 'techcrewcr@gmail.com';
  
  if (!clientEmail) return;

  var nombre = data.nombre || 'Cliente';
  var servicio = data.servicio_interes || 'Asesoría General';
  var fecha = data.fecha || '';
  var hora = data.hora || '';
  
  var dateParts = fecha.split("-");
  var formattedDate = dateParts.length === 3 ? dateParts[2] + "/" + dateParts[1] + "/" + dateParts[0] : fecha;
  
  // Plantilla HTML de confirmación para el cliente
  var clientHtml = 
    '<div style="font-family: Arial, sans-serif; background-color: #0d0d0d; color: #ffffff; padding: 25px; border-radius: 12px; max-width: 600px; border: 1px solid #f1cd00; margin: 0 auto;">' +
      '<div style="text-align: center; margin-bottom: 20px;">' +
        '<span style="font-size: 45px;">📅</span>' +
        '<h2 style="color: #f1cd00; margin: 10px 0;">¡Llamada Confirmada!</h2>' +
      '</div>' +
      '<p style="font-size: 16px; line-height: 1.5; color: #e0e0e0;">Hola <strong>' + nombre + '</strong>,</p>' +
      '<p style="font-size: 15px; line-height: 1.5; color: #cccccc;">Tu sesión de consultoría estratégica con el equipo de <strong>Tech Crew</strong> ha sido agendada con éxito. A continuación encuentras los detalles de tu cita:</p>' +
      
      '<div style="background-color: rgba(241, 205, 0, 0.08); border-left: 4px solid #f1cd00; padding: 15px; margin: 20px 0; border-radius: 4px;">' +
        '<table style="width: 100%; border-collapse: collapse; color: #ffffff;">' +
          '<tr>' +
            '<td style="padding: 6px 0; width: 40px; font-size: 20px;">📅</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Fecha:</strong> ' + formattedDate + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; font-size: 20px;">🕒</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Hora:</strong> ' + hora + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; font-size: 20px;">💼</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Servicio:</strong> ' + servicio + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; font-size: 20px;">💻</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Plataforma:</strong> Google Meet (enlace en la invitación de Calendar)</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +
      
      '<p style="font-size: 14px; line-height: 1.5; color: #aaaaaa;">Hemos enviado una invitación de Google Calendar a tu correo electrónico. Por favor, asegúrate de <strong>aceptar la invitación</strong> para que el evento se guarde automáticamente en tu calendario.</p>' +
      '<div style="border-top: 1px solid rgba(255,255,255,0.1); margin-top: 25px; padding-top: 15px; text-align: center; font-size: 13px; color: #888888;">' +
        'Tech Crew © 2026 • Soluciones Digitales Automatizadas' +
      '</div>' +
    '</div>';

  // Plantilla HTML para avisar a Tech Crew
  var crewHtml = 
    '<div style="font-family: Arial, sans-serif; background-color: #0d0d0d; color: #ffffff; padding: 25px; border-radius: 12px; max-width: 600px; border: 1px solid #f1cd00; margin: 0 auto;">' +
      '<div style="text-align: center; margin-bottom: 20px;">' +
        '<span style="font-size: 45px;">🔔</span>' +
        '<h2 style="color: #f1cd00; margin: 10px 0;">Nueva Llamada Agendada</h2>' +
      '</div>' +
      '<p style="font-size: 16px; line-height: 1.5; color: #e0e0e0;">¡Hola Tech Crew!</p>' +
      '<p style="font-size: 15px; line-height: 1.5; color: #cccccc;">Un cliente ha reservado una sesión de consultoría. Aquí están los detalles:</p>' +
      
      '<div style="background-color: rgba(255,255,255,0.05); padding: 15px; margin: 20px 0; border-radius: 8px; border: 1px solid rgba(241, 205, 0, 0.2);">' +
        '<h3 style="color: #f1cd00; margin-top: 0; border-bottom: 1px solid rgba(241,205,0,0.2); padding-bottom: 8px;">Datos del Cliente</h3>' +
        '<table style="width: 100%; border-collapse: collapse; color: #ffffff;">' +
          '<tr>' +
            '<td style="padding: 5px 0; width: 120px; color: #aaa;"><strong>Nombre:</strong></td>' +
            '<td style="padding: 5px 0;">' + nombre + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 5px 0; color: #aaa;"><strong>Correo:</strong></td>' +
            '<td style="padding: 5px 0;">' + clientEmail + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 5px 0; color: #aaa;"><strong>Teléfono:</strong></td>' +
            '<td style="padding: 5px 0;">' + (data.telefono || 'No especificado') + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 5px 0; color: #aaa;"><strong>Solución:</strong></td>' +
            '<td style="padding: 5px 0;">' + servicio + '</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +
      
      '<div style="background-color: rgba(241, 205, 0, 0.08); border-left: 4px solid #f1cd00; padding: 15px; margin: 20px 0; border-radius: 4px;">' +
        '<h3 style="color: #f1cd00; margin-top: 0; border-bottom: 1px solid rgba(241,205,0,0.2); padding-bottom: 8px;">Detalles de la Cita</h3>' +
        '<table style="width: 100%; border-collapse: collapse; color: #ffffff;">' +
          '<tr>' +
            '<td style="padding: 6px 0; width: 40px; font-size: 20px;">📅</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Fecha:</strong> ' + formattedDate + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; font-size: 20px;">🕒</td>' +
            '<td style="padding: 6px 0; font-size: 15px;"><strong>Hora:</strong> ' + hora + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding: 6px 0; font-size: 20px;">✉️</td>' +
            '<td style="padding: 6px 0; font-size: 15px;">Google Calendar sincronizado.</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +
      '<div style="border-top: 1px solid rgba(255,255,255,0.1); margin-top: 25px; padding-top: 15px; text-align: center; font-size: 13px; color: #888888;">' +
        'Notificación del Sistema de Reservas Tech Crew' +
      '</div>' +
    '</div>';

  try {
    // Enviar confirmación al cliente
    MailApp.sendEmail({
      to: clientEmail,
      subject: "📅 Confirmación de tu Llamada de Consultoría - Tech Crew",
      htmlBody: clientHtml
    });
    
    // Enviar alerta al equipo de Tech Crew
    MailApp.sendEmail({
      to: crewEmail,
      subject: "🔔 Nueva Llamada Agendada: " + nombre + " - Tech Crew",
      htmlBody: crewHtml
    });
  } catch (e) {
    Logger.log("Error al enviar correos: " + e.toString());
  }
}
