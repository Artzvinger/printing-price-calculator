function doPost(e) {
  try {
    if (!e || !e.postData) {
      return jsonResponse({ success: false, message: 'Нет данных в POST' });
    }

    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById('1dHH91iMv4_z-udSGFheLkQlzAXsQL7V79sLayiK3_VU');
    const sheet = ss.getSheetByName('Лист1');

    sheet.getRange('A8').setValue(data.companyName || '');
    sheet.getRange('D8').setValue(data.companyAddress || '');
    sheet.getRange('F8').setValue(data.companyContacts || '');

    sheet.getRange('A12').setValue(data.productName || '');
    sheet.getRange('C12').setValue(Number(data.quantity) || 0);
    sheet.getRange('D12').setValue(Number(data.perSheet) || 0);
    sheet.getRange('E12').setValue(data.notes || '');

    const materialText = data.materialType === 'cardboard'
      ? 'Картон 220 г/м^2'
      : 'Бумага_мелованная 150 г/м^2';
    sheet.getRange('B21').setValue(materialText);

    sheet.getRange('L13').setValue(Number(data.materialPrice) || 0);
    const currencyMap = { RUB: '₽', USD: '$', EUR: '€', '₽': '₽', '$': '$', '€': '€' };
    sheet.getRange('L12').setValue(currencyMap[data.materialCurrency] || '₽');

    sheet.getRange('F15').setValue(Number(data.printWidth) || 420);
    sheet.getRange('H15').setValue(Number(data.printHeight) || 297);
    sheet.getRange('F16').setValue(Number(data.purchaseWidth) || 420);
    sheet.getRange('H16').setValue(Number(data.purchaseHeight) || 297);

    let format = (data.formatSize || 'А3').toUpperCase();
    if (!['А2', 'А3', 'А4'].includes(format)) format = 'А3';
    sheet.getRange('I16').setValue(format);

    sheet.getRange('E3').setValue(Number(data.usdRate) || 90);
    sheet.getRange('F3').setValue(Number(data.eurRate) || 100);

    const operations = {
      'B22': data.opCutting,      // Подрезка
      'B26': data.opPrinting,     // Печать
      'B27': data.opLamination,   // Ламинация
      'B28': data.opUV,           // УФ-лак
      'B29': data.opCutting2,     // Резка
      'B30': data.opEmbossing1,   // Тиснение (1)
      'B31': data.opEmbossing2,   // Тиснение (2)
      'B32': data.opDieCutting,   // Вырубка
      'B34': data.opGluing,       // Склейка
      'B35': data.opBinding       // Брошюровка
    };

    for (const cell in operations) {
      const normalized = normalizeOperationValue(cell, operations[cell]);
      sheet.getRange(cell).setValue(normalized);
    }

    sheet.getRange('F8').setValue(data.shippingDate || '');

    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    const results = {
      success: true,
      message: 'Расчёт выполнен успешно',
      sheetsKg: sheet.getRange('C16').getDisplayValue(),
      circulation: sheet.getRange('D16').getDisplayValue(),
      total: sheet.getRange('H40').getDisplayValue(),
      vat: sheet.getRange('H41').getDisplayValue(),
      final: sheet.getRange('H42').getDisplayValue(),
    };

    return jsonResponse(results);

  } catch (err) {
    return jsonResponse({
      success: false,
      message: 'Ошибка обработки данных',
      error: err.message,
    });
  }
}

function normalizeOperationValue(cell, value) {
  if (!value) return 'нет';
  const v = value.trim().toLowerCase();

  switch (cell) {
    case 'B22': // Подрезка
      if (v.includes('а2')) return 'на формат А2';
      if (v.includes('а3')) return 'на формат А3';
      if (v.includes('а4')) return 'на формат А4';
      return 'нет';

    case 'B26': // Печать
      if (v.includes('1+0')) return '1+0';
      if (v.includes('2+0')) return '2+0';
      if (v.includes('3+0')) return '3+0';
      if (v.includes('4+0')) return '4+0';
      if (v.includes('5+0')) return '5+0';
      return 'нет';

    case 'B27': // Ламинация
    case 'B28': // УФ-лак
      if (v.includes('глянцевая 1+0')) return 'Глянцевая 1+0';
      if (v.includes('глянцевая 1+1')) return 'Глянцевая 1+1';
      return 'нет';

    case 'B30': // Тиснение 1
    case 'B31': // Тиснение 2
      if (v.includes('фольгой 1')) return 'фольгой 1 клише';
      if (v.includes('фольгой 2')) return 'фольгой 2 клише';
      if (v.includes('конгревное 1')) return 'конгревное 1 клише';
      if (v.includes('конгревное 2')) return 'конгревное 2 клише';
      return 'нет';

    case 'B29': // Резка
      if (v.includes('А3')) return 'на формат А3'; 
      if (v.includes('А4')) return 'на формат А4'; 
      return 'нет';

    case 'B32': // Вырубка
      if (v.includes('1 изделие')) return '1 изделие';
      if (v.includes('2 изделие')) return '2 изделие';
      if (v.includes('3 изделие')) return '3 изделие';
      return 'нет';

    case 'B34': // Склейка
      if (v.includes('1 точка')) return '1 точка';
      if (v.includes('2 точки')) return '2 точки';
      return 'нет';

    case 'B35': // Брошюровка
      if (v.includes('на клей')) return 'на клей';
      if (v.includes('на нитку')) return 'на нитку';
      if (v.includes('на пружину')) return 'на пружину';
      return 'нет';

    default:
      return 'нет';
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonResponse({ success: true, message: 'API работает. Используй POST.' });
}
