// ============================================================
//  REGISTRO DE TURNOS — Google Apps Script
//  Hoja: https://docs.google.com/spreadsheets/d/1OpweQoePlFN_5lISoFja3FWGc35iHv9aGvr6etYdrpM
// ============================================================
//
//  PASOS PARA PUBLICAR COMO WEB APP:
//  1. Abre: https://script.google.com  → "Nuevo proyecto"
//  2. Pega TODO este código (borra el código vacío que aparece)
//  3. Menú → Implementar → Nueva implementación
//  4. Tipo: "Aplicación web"
//       - Ejecutar como: "Yo"
//       - Quién tiene acceso: "Cualquier persona"
//  5. Clic en "Implementar" → Autoriza permisos
//  6. Copia la URL que aparece (empieza con https://script.google.com/macros/s/...)
//  7. Pega esa URL en el campo "Configuración" del index.html
// ============================================================

var SHEET_ID = '1OpweQoePlFN_5lISoFja3FWGc35iHv9aGvr6etYdrpM';

var HEADERS = [
  'Timestamp',
  'Fecha',
  'Empleado',
  'Lugar',
  'Hora Entrada',
  'Hora Salida',
  'Horas Trabajadas',
  'Lonas',
  'Préstamos',
  'Novedades'
];

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheets()[0];

  // Crear encabezados si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1976D2')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var sheet  = getSheet();

    // Acepta tanto un registro individual como un array (envío de quincena)
    var registros = Array.isArray(data) ? data : [data];

    registros.forEach(function(r) {
      sheet.appendRow([
        r.timestamp       || new Date().toLocaleString('es-MX'),
        r.fecha           || '',
        r.empleado        || '',
        r.area            || '',
        r.entrada         || '',
        r.salida          || '',
        r.horasTrabajadas || '',
        r.lonas           || 0,
        r.prestamos       || 0,
        r.novedades       || ''
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', count: registros.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Lee registros filtrados por año / mes / quincena ────────────────────────
//  GET ?year=2026&month=6&q=1
//  Devuelve: { status:'ok', registros:[{fecha, empleado, lugar, entrada, ...}] }
//  Si hay duplicados (mismo empleado+fecha), gana el registro más reciente.
function doGet(e) {
  try {
    var p     = e.parameter;
    var year  = parseInt(p.year  || 0);
    var month = parseInt(p.month || 0);
    var q     = parseInt(p.q     || 0);

    if (!year || !month || !q) {
      return jsonOut({ status: 'error', message: 'Parámetros requeridos: year, month, q' });
    }

    var startDay = (q === 1) ? 1  : 16;
    var endDay   = (q === 1) ? 15 : new Date(year, month, 0).getDate();

    var rows  = getSheet().getDataRange().getValues();
    var mapa  = {};   // empleado_fecha → registro (última escritura gana)

    for (var i = 1; i < rows.length; i++) {
      var row   = rows[i];
      var fecha = toFechaStr(row[1]);
      if (!fecha) continue;

      var parts = fecha.split('-');
      if (parts.length !== 3) continue;
      var rowY = parseInt(parts[0]);
      var rowM = parseInt(parts[1]);
      var rowD = parseInt(parts[2]);

      if (rowY !== year || rowM !== month || rowD < startDay || rowD > endDay) continue;

      var emp = (row[2] || '').toString().trim();
      if (!emp) continue;

      mapa[emp + '_' + fecha] = {
        fecha:     fecha,
        empleado:  emp,
        lugar:     (row[3] || '').toString(),
        entrada:   (row[4] || '').toString(),
        salida:    (row[5] || '').toString(),
        lonas:     Number(row[7]) || 0,
        prestamos: Number(row[8]) || 0,
        novedades: (row[9] || '').toString()
      };
    }

    return jsonOut({ status: 'ok', registros: Object.values(mapa) });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// Convierte un valor de celda (Date u objeto) a 'YYYY-MM-DD'
// Usa UTC para evitar que el huso horario desplace el día al leer desde Sheets
function toFechaStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getUTCFullYear();
    var m = String(val.getUTCMonth() + 1).padStart(2, '0');
    var d = String(val.getUTCDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return val.toString().split('T')[0].trim();
}

// Helper para respuestas JSON
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Diagnóstico: muestra en Logger qué fechas ve el script en la hoja ────────
//  1. Abre el editor de Apps Script
//  2. Cambia los valores de year/month/q al período que quieres probar
//  3. Selecciona "testDoGet" en el menú desplegable de funciones y clic Ejecutar
//  4. Ve a "Registros de ejecución" para ver el resultado
function testDoGet() {
  var e = { parameter: { year: '2026', month: '6', q: '1' } };
  var result = doGet(e);
  Logger.log(result.getContent());
}

// ── Función de prueba (ejecutar manualmente desde el editor) ─────────────────
function testInsert() {
  var sheet = getSheet();
  sheet.appendRow([
    new Date().toLocaleString('es-MX'),
    '2026-06-06',
    'Empleado Prueba',
    'Taller',
    '08:00',
    '16:00',
    '8h 00m',
    5,
    0,
    'Prueba manual desde editor'
  ]);
  Logger.log('Fila de prueba insertada correctamente.');
}
