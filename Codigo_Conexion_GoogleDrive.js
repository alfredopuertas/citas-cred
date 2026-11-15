/**
 * BACKEND GOOGLE APPS SCRIPT - SISTEMA DE CITAS CRED Y SULFATO FERROSO (C.S. CHILCA)
 * ====================================================================================
 * INSTRUCCIONES PASO A PASO:
 * 1. Sube y abre tu archivo 'Plantilla_Maestra_GoogleDrive_CRED.xlsx' en Google Drive
 *    y guárdalo como Hoja de Cálculo de Google (Archivo > Guardar como hojas de cálculo de Google).
 * 2. En el menú superior, haz clic en Extensiones > Apps Script.
 * 3. Borra todo el texto que salga y pega este código completo.
 * 4. Haz clic en el botón azul "Implementar" > "Nueva implementación".
 * 5. En "Seleccionar tipo" elige "Aplicación web".
 * 6. En "Ejecutar como" elige "Yo".
 * 7. En "Quién tiene acceso" elige "Cualquier persona".
 * 8. Haz clic en "Implementar", autoriza los permisos y copia la URL de la Aplicación Web.
 * 9. Pega esa URL en el panel de Ajustes de Administrador (PIN 9999) dentro del aplicativo.
 */

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsExist(ss);

  const personal = getSheetData(ss.getSheetByName('Personal_Autorizado'));
  const consultorios = getSheetData(ss.getSheetByName('Habilitar_Consultorio_B'));
  const citas = getSheetData(ss.getSheetByName('Registro_Citas_En_Vivo'));

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    timestamp: new Date().toISOString(),
    personal: personal,
    consultorios: consultorios,
    citas: citas
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // Evita cruces si 2 enfermeras agendan al mismo segundo

  try {
    const rawContent = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(rawContent);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetsExist(ss);

    if (payload.action === 'CREAR_CITA') {
      const sheet = ss.getSheetByName('Registro_Citas_En_Vivo');
      const d = payload.data || {};
      const fecha = d.date || d.fecha || '';
      const turno = d.shift === 'TARDE' ? 'Tarde' : (d.shift === 'MANANA' ? 'Mañana' : (d.turno || 'Mañana'));
      const hora = d.time || d.hora || '';
      const bloques = Number(d.blocks || d.bloques || 1);
      const consultorio = String(d.consultorio || 'A').replace('Consultorio ', '');
      const hcl = d.hcl || '';
      const dni = d.dni || '';
      const nombreMenor = d.name || d.nombreMenor || '';
      const fechaNac = d.dob || d.fechaNac || '';
      const servicio = d.service || d.servicio || 'CRED';
      const prematuro = (d.isPrematuro || d.prematuro === 'SI' || d.prematuro === true) ? 'SI' : 'NO';
      const nombrePadre = d.parentName || d.nombrePadre || '';
      const celular1 = d.phone1 || d.celular1 || '';
      const celular2 = d.phone2 || d.celular2 || '';
      const estado = d.status || d.estado || 'PENDIENTE';
      const registradoPor = d.bookedBy || d.registradoPor || '';
      const detalle = d.detail || '';

      // Si ya existe ese ID (ej. reprogramación), actualiza la fila en lugar de duplicarla
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 3; i < rows.length; i++) {
        if (String(rows[i][0]) === String(d.id)) {
          sheet.getRange(i + 1, 1, 1, 19).setValues([[
            d.id, fecha, turno, hora, bloques, bloques * 20,
            'Consultorio ' + consultorio, hcl, dni, nombreMenor, fechaNac,
            servicio, prematuro, nombrePadre, celular1, celular2, estado, registradoPor, detalle
          ]]);
          updated = true;
          break;
        }
      }

      if (!updated) {
        sheet.appendRow([
          d.id, fecha, turno, hora, bloques, bloques * 20,
          'Consultorio ' + consultorio, hcl, dni, nombreMenor, fechaNac,
          servicio, prematuro, nombrePadre, celular1, celular2, estado, registradoPor, detalle
        ]);
      }

    } else if (payload.action === 'ACTUALIZAR_ESTADO') {
      const sheet = ss.getSheetByName('Registro_Citas_En_Vivo');
      const rows = sheet.getDataRange().getValues();
      for (let i = 3; i < rows.length; i++) {
        if (String(rows[i][0]) === String(payload.id)) {
          sheet.getRange(i + 1, 17).setValue(payload.estado);
          if (payload.nuevaFecha) sheet.getRange(i + 1, 2).setValue(payload.nuevaFecha);
          if (payload. nuevoTurno) sheet.getRange(i + 1, 3).setValue(payload.nuevoTurno);
          if (payload.nuevaHora) sheet.getRange(i + 1, 4).setValue(payload.nuevaHora);
          if (payload.nuevoConsultorio) sheet.getRange(i + 1, 7).setValue('Consultorio ' + String(payload.nuevoConsultorio).replace('Consultorio ', ''));
          break;
        }
      }

    } else if (payload.action === 'ELIMINAR_CITA') {
      const sheet = ss.getSheetByName('Registro_Citas_En_Vivo');
      const rows = sheet.getDataRange().getValues();
      for (let i = 3; i < rows.length; i++) {
        if (String(rows[i][0]) === String(payload.id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }

    } else if (payload.action === 'TOGGLE_CONSULTORIO_B') {
      const sheet = ss.getSheetByName('Habilitar_Consultorio_B');
      const rows = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 3; i < rows.length; i++) {
        const cellDate = formatCellDate(rows[i][0]);
        if (cellDate === payload.fecha) {
          sheet.getRange(i + 1, 3).setValue(payload.habilitarB ? 'SI' : 'NO');
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([payload.fecha, 'ACTIVO', payload.habilitarB ? 'SI' : 'NO', '', '']);
      }

    } else if (payload.action === 'GUARDAR_PERSONAL') {
      const sheet = ss.getSheetByName('Personal_Autorizado');
      if (Array.isArray(payload.personal) && payload.personal.length > 0) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 4) {
          sheet.getRange(4, 1, lastRow - 3, 6).clearContent();
        }
        const newRows = payload.personal.map(u => [
          u.id || '',
          u.name || '',
          u.role || 'Enfermera',
          u.code || '',
          u.pin || '1001',
          u.active ? 'ACTIVO' : 'SUSPENDIDO'
        ]);
        sheet.getRange(4, 1, newRows.length, 6).setValues(newRows);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function formatCellDate(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'America/Lima', 'yyyy-MM-dd');
  }
  return String(val).substring(0, 10);
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 4) return [];
  const headers = values[2];
  return values.slice(3).filter(r => r[0] !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (!h) return;
      let val = row[i];
      if (Object.prototype.toString.call(val) === '[object Date]') {
        // Si la columna es Hora_Inicio, formatear HH:mm; si es fecha, yyyy-MM-dd
        if (String(h).toLowerCase().indexOf('hora') !== -1) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone() || 'America/Lima', 'HH:mm');
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone() || 'America/Lima', 'yyyy-MM-dd');
        }
      }
      obj[h] = val;
    });
    return obj;
  });
}

function ensureSheetsExist(ss) {
  if (!ss.getSheetByName('Personal_Autorizado')) {
    const s1 = ss.insertSheet('Personal_Autorizado');
    s1.getRange(1, 1).setValue('CENTRO DE SALUD CHILCA - GESTIÓN DE PERSONAL AUTORIZADO (CRED / SULFATO FERROSO)');
    s1.getRange(3, 1, 1, 6).setValues([['ID_Usuario', 'Nombre_Completo', 'Rol_Sistema', 'DNI_Personal', 'PIN_Acceso', 'Estado']]);
  }
  if (!ss.getSheetByName('Habilitar_Consultorio_B')) {
    const s2 = ss.insertSheet('Habilitar_Consultorio_B');
    s2.getRange(1, 1).setValue('PROGRAMACIÓN DE CONSULTORIOS POR FECHA (ORDEN DE JEFATURA)');
    s2.getRange(3, 1, 1, 5).setValues([['Fecha (AAAA-MM-DD)', 'Consultorio_A', 'Habilitar_Consultorio_B', 'Enfermera_Turno_Mañana', 'Enfermera_Turno_Tarde']]);
  }
  if (!ss.getSheetByName('Registro_Citas_En_Vivo')) {
    const s3 = ss.insertSheet('Registro_Citas_En_Vivo');
    s3.getRange(1, 1).setValue('BASE DE DATOS EN VIVO DE CITAS AGENDADAS (CRED Y SULFATO FERROSO)');
    s3.getRange(3, 1, 1, 19).setValues([[
      'ID_Cita', 'Fecha_Cita', 'Turno', 'Hora_Inicio', 'Bloques_20min', 'Duracion_Min',
      'Consultorio', 'Nro_HCl', 'DNI_Menor', 'Nombre_Menor', 'Fecha_Nacimiento',
      'Servicio', 'CRED_Prematuro', 'Nombre_Padre_Madre', 'Celular_Principal',
      'Celular_Secundario', 'Estado_Cita', 'Registrado_Por', 'Detalle_Nota'
    ]]);
  }
}
