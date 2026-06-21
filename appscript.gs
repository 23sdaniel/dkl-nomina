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
  'Nombre',
  'Lugar',
  'Hora Ingreso',
  'Hora Salida',
  'Horas Trabajadas',
  'Lonas',
  'Prestamo',
  'Novedades',
  'Descanso',
  'Ausencia'
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
    var data      = JSON.parse(e.postData.contents);
    var sheet     = getSheet();
    var registros = Array.isArray(data) ? data : [data];

    // Leer hoja una sola vez y construir índice EMPLEADO_FECHA → fila (1-based)
    var filas  = sheet.getDataRange().getValues();
    var indice = {};
    for (var i = 1; i < filas.length; i++) {
      var fe = toFechaStr(filas[i][1]);
      var em = (filas[i][2] || '').toString().trim().toUpperCase();
      if (fe && em) indice[em + '_' + fe] = i + 1;
    }

    registros.forEach(function(r) {
      var key = (r.empleado || '').toString().trim().toUpperCase()
                + '_' + (r.fecha || '').toString().trim();

      var fila = [
        r.timestamp       || new Date().toLocaleString('es-MX'),
        r.fecha           || '',
        r.empleado        || '',
        r.area            || '',
        r.entrada         || '',
        r.salida          || '',
        r.horasTrabajadas || '',
        r.lonas           || 0,
        r.prestamos       || 0,
        r.novedades       || '',
        r.descanso        || '',
        r.ausencia        || ''
      ];

      if (indice[key]) {
        // Ya existe → sobreescribir esa fila
        sheet.getRange(indice[key], 1, 1, fila.length).setValues([fila]);
      } else {
        // Nueva combinación empleado+fecha → agregar al final
        sheet.appendRow(fila);
        indice[key] = sheet.getLastRow(); // evita duplicados dentro del mismo envío
      }
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
    var debugRows = []; // solo se usa cuando ?debug=1

    for (var i = 1; i < rows.length; i++) {
      var row   = rows[i];
      var fecha = toFechaStr(row[1]);

      // Modo diagnóstico: captura los primeros 5 datos crudos
      if (p.debug === '1' && i <= 5) {
        debugRows.push({
          fila:       i + 1,
          raw:        String(rows[i][1]),
          esDate:     rows[i][1] instanceof Date,
          fechaParsed: fecha,
          empleado:   String(rows[i][2])
        });
      }

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
        lugar:     (row[3]  || '').toString(),
        entrada:   toTimeStr(row[4]),
        salida:    toTimeStr(row[5]),
        lonas:     Number(row[7]) || 0,
        prestamos: Number(row[8]) || 0,
        novedades: (row[9]  || '').toString(),
        descanso:  (row[10] || '').toString(),
        ausencia:  (row[11] || '').toString()
      };
    }

    var resultado = { status: 'ok', registros: Object.values(mapa) };
    if (p.debug === '1') {
      resultado.debug = { params: {year:year,month:month,q:q,startDay:startDay,endDay:endDay}, filas: debugRows };
    }
    return jsonOut(resultado);

  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// Convierte un valor de celda de Sheets a 'YYYY-MM-DD'
function toFechaStr(val) {
  if (!val) return '';
  if (typeof val === 'object' && typeof val.getFullYear === 'function') {
    var tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  }
  return String(val).trim().substring(0, 10);
}

// Convierte un valor de celda de Sheets a 'HH:mm'
// Las horas se guardan como objetos Date (base 30-dic-1899) en Google Sheets
function toTimeStr(val) {
  if (!val) return '';
  if (typeof val === 'object' && typeof val.getHours === 'function') {
    var tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  return String(val).trim();
}

// Helper para respuestas JSON
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Limpia duplicados ya existentes en la hoja (ejecutar una sola vez) ───────
//  Conserva la ÚLTIMA fila de cada combinación empleado+fecha y borra las demás.
//  Ejecuta desde el editor de Apps Script: selecciona "deduplicar" y clic Ejecutar.
function deduplicar() {
  var sheet = getSheet();
  var filas = sheet.getDataRange().getValues();
  var visto = {};       // clave → índice de la fila más reciente (1-based)
  var aBorrar = [];

  for (var i = 1; i < filas.length; i++) {
    var fe  = toFechaStr(filas[i][1]);
    var em  = (filas[i][2] || '').toString().trim().toUpperCase();
    var key = em + '_' + fe;
    if (!key || key === '_') continue;

    if (visto[key] !== undefined) {
      aBorrar.push(visto[key]); // la anterior se borra
    }
    visto[key] = i + 1; // fila más reciente gana (1-based)
  }

  // Borrar de abajo hacia arriba para no alterar los índices
  aBorrar.sort(function(a,b){ return b - a; });
  aBorrar.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });

  Logger.log('Duplicados eliminados: ' + aBorrar.length);
}

// ── Diagnóstico: muestra en Logger qué fechas ve el script en la hoja ────────
//  1. Abre el editor de Apps Script
//  2. Cambia los valores de year/month/q al período que quieres probar
//  3. Selecciona "testDoGet" en el menú desplegable de funciones y clic Ejecutar
//  4. Ve a "Registros de ejecución" para ver el resultado
function testDoGet() {
  var sheet = getSheet();
  var rows  = sheet.getDataRange().getValues();
  Logger.log('Total filas (encabezado incluido): ' + rows.length);

  var year = 2026, month = 6, startDay = 1, endDay = 15;

  for (var i = 1; i <= Math.min(rows.length - 1, 6); i++) {
    var r     = rows[i];
    var fecha = toFechaStr(r[1]);
    var parts = fecha ? fecha.split('-') : [];
    var rowY  = parseInt(parts[0] || 0);
    var rowM  = parseInt(parts[1] || 0);
    var rowD  = parseInt(parts[2] || 0);
    var pasa  = (rowY === year && rowM === month && rowD >= startDay && rowD <= endDay);
    Logger.log(
      'Fila ' + (i + 1) + ': ' +
      'raw="' + String(r[1]) + '" ' +
      'tipo=' + typeof r[1] + ' ' +
      'esDate=' + (r[1] instanceof Date) + ' ' +
      '→ parsed="' + fecha + '" ' +
      '(Y=' + rowY + ' M=' + rowM + ' D=' + rowD + ') ' +
      'PASA_FILTRO=' + pasa + ' | ' +
      'empleado="' + String(r[2]) + '"'
    );
  }
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
    'Prueba manual desde editor',
    '',
    ''
  ]);
  Logger.log('Fila de prueba insertada correctamente.');
}
