const SPREADSHEET_ID = '1hA3sn6BlJ_78hIwly5NCAqkeGwpaAiaWZir9DTWLH7s';
const SHEET_NAME = 'Hoja 1';
const ADMIN_EMAIL = 'DIEGODANTEDANIELOLIVARESTELLO@GMAIL.COM';

function doGet(event) {
  const params = event && event.parameter ? event.parameter : {};
  if (params.action === 'claim' && params.callback) {
    const callback = String(params.callback).replace(/[^\w$]/g, '');
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const result = JSON.parse(doPost({ parameter: params }).getContent());
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } finally {
      lock.releaseLock();
    }
  }
  return outputJson({ ok: true, service: 'raspadita-cristian' });
}

function doPost(event) {
  try {
    const data = parseRequest(event);
    validateRequest(data);

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('No existe la hoja "' + SHEET_NAME + '".');

    const columns = ensureHeaders(sheet);
    const cardsSheet = ensureCardsSheet();
    const row = Array(sheet.getLastColumn()).fill('');
    const values = {
      id: Utilities.getUuid(),
      nombre: data.nombre.trim(),
      dni: data.dni.replace(/\D/g, ''),
      correo: data.correo.trim().toLowerCase(),
      premio: data.premio || 'Sin premio',
      tarjeta: data.tarjeta || '',
      ganador: data.ganador === true || data.ganador === 'true' ? 'Sí' : 'No',
      fecha: new Date()
    };

    const existing = sheet.getDataRange().getValues();
    const dniColumn = columns.filter(function (column) { return column.key === 'dni'; })[0];
    const nameColumn = columns.filter(function (column) { return column.key === 'nombre'; })[0];
    const emailColumn = columns.filter(function (column) { return column.key === 'correo'; })[0];
    const existingIndex = existing.slice(1).findIndex(function (existingRow) {
      return String(existingRow[dniColumn.index]).replace(/\D/g, '') === values.dni &&
        String(existingRow[nameColumn.index]).trim().toLowerCase() === values.nombre.toLowerCase() &&
        String(existingRow[emailColumn.index]).trim().toLowerCase() === values.correo;
    });
    if (existingIndex >= 0 && data.stage !== 'resultado') {
      throw new Error('Esta persona ya participó.');
    }
    if (data.stage === 'registro' && values.tarjeta) {
      const cardRow = findCardRow(cardsSheet, values.tarjeta);
      if (!cardRow) throw new Error('La tarjeta indicada no existe.');
      if (String(cardRow[3]).toLowerCase() !== 'disponible') {
        throw new Error('Esa tarjeta ya fue elegida por otra persona.');
      }
      cardsSheet.getRange(cardRow[0], 3, 1, 6).setValues([[
        'Usada', values.nombre, values.dni, values.correo, 'No', new Date()
      ]]);
    }
    if (data.stage === 'resultado' && values.tarjeta) {
      const cardRow = findCardRow(cardsSheet, values.tarjeta);
      if (cardRow) cardsSheet.getRange(cardRow[0], 7).setValue(values.ganador);
    }

    columns.forEach(function (column) {
      if (Object.prototype.hasOwnProperty.call(values, column.key)) row[column.index] = values[column.key];
    });
    if (existingIndex >= 0) {
      sheet.getRange(existingIndex + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    if (values.ganador === 'Sí' && data.stage === 'resultado') {
      sendWinnerEmail(values);
      notifyAdmin(values);
    }

    return outputJson({ ok: true });
  } catch (error) {
    console.error(error);
    return outputJson({ ok: false, error: error.message });
  }
}

function parseRequest(event) {
  if (!event) {
    throw new Error('La solicitud no contiene datos.');
  }
  if (event.parameter && event.parameter.nombre) return event.parameter;
  if (!event.postData || !event.postData.contents) {
    throw new Error('La solicitud no contiene datos.');
  }
  const contents = event.postData.contents;
  try {
    return JSON.parse(contents);
  } catch (error) {
    return event.parameter || {};
  }
}

function validateRequest(data) {
  if (!data || !data.nombre || !data.dni || !data.correo) {
    throw new Error('Faltan nombre, DNI o correo.');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.correo))) {
    throw new Error('El correo no es válido.');
  }
}

function ensureHeaders(sheet) {
  const required = [
    { key: 'id', label: 'id' },
    { key: 'nombre', label: 'Nombre' },
    { key: 'dni', label: 'dni' },
    { key: 'correo', label: 'correo' },
    { key: 'premio', label: 'premio' },
    { key: 'tarjeta', label: 'tarjeta' },
    { key: 'ganador', label: 'ganador' },
    { key: 'fecha', label: 'fecha' }
  ];
  const width = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, width).getValues()[0];
  const normalized = existing.map(function (value) {
    return String(value).trim().toLowerCase();
  });

  required.forEach(function (column) {
    if (normalized.indexOf(column.label.toLowerCase()) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(column.label);
      normalized.push(column.label.toLowerCase());
    }

  });

  return required.map(function (column) {
    return { key: column.key, index: normalized.indexOf(column.label.toLowerCase()) };
  }).sort(function (a, b) {
    return a.index - b.index;
  });
}

function ensureCardsSheet() {
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = book.getSheetByName('Tarjetas');
  if (!sheet) {
    sheet = book.insertSheet('Tarjetas');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'numero', 'premio', 'estado', 'nombre', 'dni', 'correo', 'ganador', 'fecha'
    ]]);
    const prizes = {};
    [100, 88, 167, 25, 8].forEach(function (number) { prizes[number] = 'Factura'; });
    [36, 70, 147, 14, 51].forEach(function (number) { prizes[number] = 'Te llevo el pan a la mesa'; });
    [200, 201, 202, 203].forEach(function (number) { prizes[number] = '6 churros'; });
    [204, 205, 206, 207].forEach(function (number) { prizes[number] = '6 donas'; });
    prizes[39] = 'Premio mayor';
    const rows = [];
    for (let number = 1; number <= 500; number += 1) {
      rows.push([number, prizes[number] || 'Sin premio', 'Disponible', '', '', '', 'No', '']);
    }
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return sheet;
}

function findCardRow(sheet, number) {
  const values = sheet.getDataRange().getValues();
  const target = String(number).trim();
  for (let index = 1; index < values.length; index += 1) {
    if (String(values[index][0]).trim() === target) return [index + 1].concat(values[index]);
  }
  return null;
}

function sendWinnerEmail(values) {
  const name = escapeHtml(values.nombre);
  const prize = escapeHtml(values.premio);
  const card = escapeHtml(values.tarjeta);
  const htmlBody = [
    '<div style="margin:0;padding:32px 16px;background:#fff7ed;font-family:Arial,sans-serif;color:#3b2415;">',
    '<div style="max-width:620px;margin:0 auto;background:#ffffff;border-radius:22px;overflow:hidden;border:1px solid #fed7aa;box-shadow:0 8px 24px rgba(124,45,18,.12);">',
    '<div style="padding:28px 24px;text-align:center;background:linear-gradient(135deg,#f97316,#facc15);color:#ffffff;">',
    '<div style="font-size:42px;line-height:1;">🎉 🥳 🎉</div>',
    '<h1 style="margin:14px 0 6px;font-size:30px;">¡Tenemos una gran noticia!</h1>',
    '<p style="margin:0;font-size:17px;font-weight:bold;">¡Ganaste en Jugá con Cristian!</p>',
    '</div>',
    '<div style="padding:30px 28px;">',
    '<p style="margin:0 0 18px;font-size:20px;">¡Felicitaciones, <strong>' + name + '</strong>! 👏</p>',
    '<p style="margin:0 0 22px;font-size:16px;line-height:1.6;">La suerte estuvo de tu lado y tu raspadita tiene premio. Gracias por participar y compartir este momento con nosotros. 💛</p>',
    '<div style="margin:22px 0;padding:22px;border-radius:16px;background:#fff7ed;border:2px dashed #fb923c;text-align:center;">',
    '<div style="font-size:15px;color:#9a3412;text-transform:uppercase;letter-spacing:1px;font-weight:bold;">Tu premio</div>',
    '<div style="margin-top:8px;font-size:28px;color:#c2410c;font-weight:bold;">🏆 ' + prize + '</div>',
    '<div style="margin-top:12px;font-size:14px;color:#7c2d12;">Tarjeta: <strong>' + card + '</strong></div>',
    '</div>',
    '<p style="margin:22px 0 0;font-size:16px;line-height:1.6;">Comunicáte con <strong>Cristian</strong> para recibir tu premio: <strong>2614541378</strong>. Si tenés alguna consulta, respondé este correo. 📩</p>',
    '<p style="margin:26px 0 0;text-align:center;font-size:20px;font-weight:bold;color:#ea580c;">¡Gracias por jugar! 🍞✨</p>',
    '</div>',
    '<div style="padding:16px 24px;text-align:center;background:#ffedd5;color:#9a3412;font-size:12px;">Jugá con Cristian · Panadería Cristian</div>',
    '</div>',
    '</div>'
  ].join('');
  MailApp.sendEmail({
    to: values.correo,
    subject: '¡Ganaste en Jugá con Cristian!',
    body: '¡Felicitaciones, ' + values.nombre + '! Ganaste ' + values.premio +
      '. Comunicáte con Cristian para recibir tu premio: 2614541378.',
    htmlBody: htmlBody
  });
}

function notifyAdmin(values) {
  if (!ADMIN_EMAIL || ADMIN_EMAIL.indexOf('REEMPLAZAR_') === 0) return;
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: '¡Alguien ganó! Tarjeta ' + values.tarjeta,
    htmlBody: '<p><strong>' + escapeHtml(values.nombre) + '</strong> ganó <strong>' +
      escapeHtml(values.premio) + '</strong>.</p><p>DNI: ' +
      escapeHtml(values.dni) + '<br>Correo: ' + escapeHtml(values.correo) + '</p>'
  });
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, function (character) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character];
  });
}

function outputJson(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
