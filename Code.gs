function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.logoDataUri = getLogoDataUri();

  return template.evaluate()
    .setTitle('Teachers Evaluation');
}

function getLogoDataUri() {
  const folders = DriveApp.getFoldersByName('images');
  if (!folders.hasNext()) {
    throw new Error('Google Drive folder "images" was not found.');
  }

  const folder = folders.next();
  const files = folder.getFilesByName('logo.png');
  if (!files.hasNext()) {
    throw new Error('File "logo.png" was not found in the Google Drive folder "images".');
  }

  const blob = files.next().getBlob();
  const contentType = blob.getContentType();
  if (!/^image\/(png|jpeg|gif|webp|svg\+xml)$/i.test(contentType)) {
    throw new Error('The file "logo.png" in the "images" folder is not a supported image.');
  }

  return 'data:' + contentType + ';base64,' +
    Utilities.base64Encode(blob.getBytes());
}

function submitForm(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('No teacher evaluations were provided.');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('The response spreadsheet could not be opened.');
  }

  const headers = [
    'Submitted At',
    'Teacher',
    'Subject Knowledge',
    'Subject Knowledge Comment',
    'Teaching Skill',
    'Teaching Skill Comment',
    'Sincerity towards the Cadet',
    'Sincerity Comment',
    'Personality and Behaviour',
    'Personality and Behaviour Comment',
    'Class',
    'House',
    'Form',
    'Cadet Name',
    'Cadet Number'
  ];
  const sheet = getResponseSheet(spreadsheet, 'Responses', headers);

  const rows = data.map(function(response) {
    if (!response.teacher) {
      throw new Error('A teacher name is missing from an evaluation.');
    }

    return [
      new Date(),
      response.teacher || '',
      response.sk1 || '',
      response.comment1 || '',
      response.ts1 || '',
      response.comment2 || '',
      response.cm1 || '',
      response.comment3 || '',
      response.cs1 || '',
      response.comment4 || '',
      response.className || '',
      response.house || '',
      response.form || '',
      response.cadetName || '',
      response.cadetNumber || ''
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);

  rows.forEach(function(row) {
    const teacherSheetName = getTeacherSheetName(row[1]);
    const teacherSheet = getResponseSheet(spreadsheet, teacherSheetName, headers);
    teacherSheet.getRange(teacherSheet.getLastRow() + 1, 1, 1, row.length)
      .setValues([row]);
  });

  return 'All ' + rows.length + ' teacher evaluations have been saved to Responses and their individual teacher tabs.';
}

function getResponseSheet(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function getTeacherSheetName(teacherName) {
  const sheetName = String(teacherName)
    .replace(/[\[\]:*?\/\\]/g, '')
    .replace(/^'+|'+$/g, '')
    .trim()
    .substring(0, 100);

  if (!sheetName || sheetName === 'Responses') {
    throw new Error('The teacher name cannot be used as a sheet name.');
  }

  return sheetName;
}