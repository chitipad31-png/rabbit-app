function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บันทึกรับ-จ่าย');

    if (data.action === 'delete') {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 8).getValue().toString() === data.id) {
          sheet.deleteRow(i);
          break;
        }
      }
      return jsonOutput_({status: 'ok'});
    }

    if (data.action === 'add') {
      sheet.appendRow([
        data.date, data.category, data.name,
        data.amount, data.note, data.type,
        new Date(), data.id
      ]);

      addItemToMasterData_(ss, data.category, data.name);
      return jsonOutput_({status: 'ok'});
    }

    return jsonOutput_({status: 'error', message: 'Unknown action'});
  } catch (err) {
    return jsonOutput_({status: 'error', message: err.toString()});
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function addItemToMasterData_(ss, category, name) {
  const cleanName = String(name || '').trim();
  const cleanCategory = String(category || '').trim();
  if (!cleanName || !cleanCategory) return;

  const aliases = {
    'LINE MAN': 'ยอดขายLINE MAN'
  };
  const masterCategory = aliases[cleanCategory] || cleanCategory;
  const masterSheet = ss.getSheetByName('MasterData');
  if (!masterSheet) throw new Error('ไม่พบชีต MasterData');

  let lastColumn = Math.max(masterSheet.getLastColumn(), 1);
  let headers = masterSheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value).trim());

  let columnIndex = headers.findIndex(header => header === masterCategory) + 1;
  if (!columnIndex) {
    columnIndex = lastColumn + 1;
    masterSheet.getRange(1, columnIndex).setValue(masterCategory);
    lastColumn = columnIndex;
    headers.push(masterCategory);
  }

  const lastRow = Math.max(masterSheet.getLastRow(), 2);
  const values = masterSheet
    .getRange(2, columnIndex, Math.max(lastRow - 1, 1), 1)
    .getDisplayValues()
    .flat();

  const normalizedName = cleanName.toLowerCase();
  const alreadyExists = values.some(value =>
    String(value).trim().toLowerCase() === normalizedName
  );
  if (alreadyExists) return;

  const firstBlankIndex = values.findIndex(value => !String(value).trim());
  const targetRow = firstBlankIndex >= 0 ? firstBlankIndex + 2 : lastRow + 1;
  masterSheet.getRange(targetRow, columnIndex).setValue(cleanName);
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (e.parameter.action === 'getItemMaster') {
      const sheet = ss.getSheetByName('MasterData');
      let uniqueItems = [];

      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          uniqueItems = [...new Set(data.slice(1).flat().filter(String))];
        }
      }

      return jsonOutput_(uniqueItems);
    }

    if (e.parameter.action === 'getHistory') {
      const sheet = ss.getSheetByName('บันทึกรับ-จ่าย');
      const records = [];

      if (sheet) {
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          let dateVal = row[0];

          if (dateVal instanceof Date) {
            const y = dateVal.getFullYear();
            const m = String(dateVal.getMonth() + 1).padStart(2, '0');
            const d = String(dateVal.getDate()).padStart(2, '0');
            dateVal = y + '-' + m + '-' + d;
          }

          if (row[7]) {
            records.push({
              date: dateVal,
              category: row[1],
              name: row[2],
              amount: row[3],
              note: row[4],
              type: row[5],
              id: row[7].toString()
            });
          }
        }
      }

      return jsonOutput_(records);
    }

    return jsonOutput_({status: 'error', message: 'Unknown action'});
  } catch (err) {
    return jsonOutput_({status: 'error', message: err.toString()});
  }
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
