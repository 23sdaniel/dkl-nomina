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

// Función de prueba (ejecutar manualmente desde el editor)
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
