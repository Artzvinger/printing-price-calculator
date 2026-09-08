function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({
        success: false,
        message: 'Нет данных в POST'
      });
    }

    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById(
      '1dHH91iMv4_z-udSGFheLkQlzAXsQL7V79sLayiK3_VU'
    );

    const sheet = ss.getSheetByName('Лист1');

    if (!sheet) {
      return jsonResponse({
        success: false,
        message: 'Лист "Лист1" не найден'
      });
    }

    // ==========================================
    // ЗАКАЗЧИК
    // ==========================================

    sheet.getRange('A8').setValue(
      data.companyName || ''
    );

    sheet.getRange('D8').setValue(
      data.companyAddress || ''
    );

    sheet.getRange('F8').setValue(
      data.companyContacts || ''
    );


    // ==========================================
    // ИЗДЕЛИЕ
    // ==========================================

    sheet.getRange('A12').setValue(
      data.productName || ''
    );

    sheet.getRange('C12').setValue(
      Number(data.quantity) || 0
    );

    sheet.getRange('D12').setValue(
      Number(data.perSheet) || 0
    );

    sheet.getRange('E12').setValue(
      data.notes || ''
    );


    // ==========================================
    // МАТЕРИАЛ
    // ==========================================

    let materialText;

    if (data.materialType === 'cardboard') {
      materialText = 'Картон 220 г/м^2';
    } else {
      materialText = 'Бумага_мелованная 150 г/м^2';
    }

    sheet.getRange('B21').setValue(materialText);

    sheet.getRange('L13').setValue(
      Number(data.materialPrice) || 0
    );

    const currencyMap = {
      'RUB': '₽',
      'USD': '$',
      'EUR': '€',
      '₽': '₽',
      '$': '$',
      '€': '€'
    };

    sheet.getRange('L12').setValue(
      currencyMap[data.materialCurrency] || '₽'
    );


    // ==========================================
    // ФОРМАТЫ
    // ==========================================

    sheet.getRange('F15').setValue(
      Number(data.printWidth) || 420
    );

    sheet.getRange('H15').setValue(
      Number(data.printHeight) || 297
    );

    sheet.getRange('F16').setValue(
      Number(data.purchaseWidth) || 420
    );

    sheet.getRange('H16').setValue(
      Number(data.purchaseHeight) || 297
    );

    let format = String(
      data.formatSize || 'А3'
    ).toUpperCase();

    if (!['А2', 'А3', 'А4'].includes(format)) {
      format = 'А3';
    }

    sheet.getRange('I16').setValue(format);


    // ==========================================
    // КУРСЫ ВАЛЮТ
    // ==========================================

    sheet.getRange('E3').setValue(
      Number(data.usdRate) || 90
    );

    sheet.getRange('F3').setValue(
      Number(data.eurRate) || 100
    );


    // ==========================================
    // ОПЕРАЦИИ
    // ==========================================

    const operations = {
      'B22': data.cuttingFormat, // Подрезка
      'B26': data.printType,     // Печать
      'B27': data.lamination,    // Ламинация
      'B28': data.uvVarnish,     // УФ-лак
      'B29': data.cutting,       // Резка
      'B30': data.embossing1,    // Тиснение 1
      'B31': data.embossing2,    // Тиснение 2
      'B32': data.dieCutting,    // Вырубка
      'B34': data.gluing,        // Склейка
      'B35': data.binding        // Брошюровка
    };

    for (const cell in operations) {
      const normalized = normalizeOperationValue(
        cell,
        operations[cell]
      );

      sheet.getRange(cell).setValue(normalized);
    }


    // ==========================================
    // ДАТА ОТГРУЗКИ
    // ==========================================
    //
    // ВАЖНО:
    // Раньше здесь использовалась F8,
    // из-за чего дата перезаписывала контакты.
    //
    // Пока оставляем дату в G8.
    // Если в твоей таблице дата должна быть
    // в другой ячейке — поменяем только её.
    //

    sheet.getRange('G8').setValue(
      data.shippingDate || ''
    );


    // ==========================================
    // ПЕРЕСЧЁТ ФОРМУЛ
    // ==========================================

    SpreadsheetApp.flush();

    // Даём формулам таблицы время пересчитаться.
    Utilities.sleep(1000);

    SpreadsheetApp.flush();


    // ==========================================
    // ПОЛУЧЕНИЕ РЕЗУЛЬТАТОВ
    // ==========================================

    const results = {
      success: true,
      message: 'Расчёт выполнен успешно',

      sheetsKg: sheet
        .getRange('C16')
        .getDisplayValue(),

      circulation: sheet
        .getRange('D16')
        .getDisplayValue(),

      total: sheet
        .getRange('H40')
        .getDisplayValue(),

      vat: sheet
        .getRange('H41')
        .getDisplayValue(),

      final: sheet
        .getRange('H42')
        .getDisplayValue()
    };


    return jsonResponse(results);

  } catch (err) {

    console.error(err);

    return jsonResponse({
      success: false,
      message: 'Ошибка обработки данных',
      error: err.message || String(err)
    });
  }
}


// ==================================================
// НОРМАЛИЗАЦИЯ ОПЕРАЦИЙ
// ==================================================

function normalizeOperationValue(cell, value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return 'нет';
  }

  const v = String(value)
    .trim()
    .toLowerCase();


  switch (cell) {

    // ------------------------------------------
    // ПОДРЕЗКА
    // ------------------------------------------

    case 'B22':

      if (v.includes('а2')) {
        return 'на формат А2';
      }

      if (v.includes('а3')) {
        return 'на формат А3';
      }

      if (v.includes('а4')) {
        return 'на формат А4';
      }

      return 'нет';


    // ------------------------------------------
    // ПЕЧАТЬ
    // ------------------------------------------

    case 'B26':

      if (v.includes('1+0')) {
        return '1+0';
      }

      if (v.includes('2+0')) {
        return '2+0';
      }

      if (v.includes('3+0')) {
        return '3+0';
      }

      if (v.includes('4+0')) {
        return '4+0';
      }

      if (v.includes('5+0')) {
        return '5+0';
      }

      return 'нет';


    // ------------------------------------------
    // ЛАМИНАЦИЯ
    // ------------------------------------------

    case 'B27':

      if (v.includes('глянцевая 1+0')) {
        return 'Глянцевая 1+0';
      }

      if (v.includes('глянцевая 1+1')) {
        return 'Глянцевая 1+1';
      }

      return 'нет';


    // ------------------------------------------
    // УФ-ЛАК
    // ------------------------------------------

    case 'B28':

      if (v.includes('глянцевая 1+0')) {
        return 'Глянцевая 1+0';
      }

      if (v.includes('глянцевая 1+1')) {
        return 'Глянцевая 1+1';
      }

      return 'нет';


    // ------------------------------------------
    // РЕЗКА
    // ------------------------------------------

    case 'B29':

      if (v.includes('а3')) {
        return 'на формат А3';
      }

      if (v.includes('а4')) {
        return 'на формат А4';
      }

      return 'нет';


    // ------------------------------------------
    // ТИСНЕНИЕ
    // ------------------------------------------

    case 'B30':
    case 'B31':

      if (v.includes('фольгой 1')) {
        return 'фольгой 1 клише';
      }

      if (v.includes('фольгой 2')) {
        return 'фольгой 2 клише';
      }

      if (v.includes('конгревное 1')) {
        return 'конгревное 1 клише';
      }

      if (v.includes('конгревное 2')) {
        return 'конгревное 2 клише';
      }

      return 'нет';


    // ------------------------------------------
    // ВЫРУБКА
    // ------------------------------------------

    case 'B32':

      if (v.includes('1 изделие')) {
        return '1 изделие';
      }

      if (v.includes('2 изделие')) {
        return '2 изделие';
      }

      if (v.includes('3 изделие')) {
        return '3 изделие';
      }

      return 'нет';


    // ------------------------------------------
    // СКЛЕЙКА
    // ------------------------------------------

    case 'B34':

      if (v.includes('1 точка')) {
        return '1 точка';
      }

      if (v.includes('2 точки')) {
        return '2 точки';
      }

      return 'нет';


    // ------------------------------------------
    // БРОШЮРОВКА
    // ------------------------------------------

    case 'B35':

      if (v.includes('на клей')) {
        return 'на клей';
      }

      if (v.includes('на нитку')) {
        return 'на нитку';
      }

      if (v.includes('на пружину')) {
        return 'на пружину';
      }

      return 'нет';


    default:
      return 'нет';
  }
}


// ==================================================
// JSON RESPONSE
// ==================================================

function jsonResponse(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


// ==================================================
// GET — ПРОВЕРКА API
// ==================================================

function doGet() {

  return jsonResponse({
    success: true,
    message: 'API работает. Используй POST.'
  });
}