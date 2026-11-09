// Google Apps Script 後端程式碼
var SPREADSHEET_ID = "1RygbA0GM8Az7rB8STRQa_aZ9zIKKnz7BGoCuxb_9-A8";

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  if (action === 'getRecords') {
    return handleGetRecords(e.parameter.employeeName);
  } else if (action === 'clockIn' || action === 'clockOut' || action === 'manualRecord') {
    // 支援 GET 請求寫入，完全免疫瀏覽器跨網域 CORS 問題
    return handleDispatch(action, e.parameter);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Clock-in API is active', spreadsheetId: SPREADSHEET_ID }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }
    var action = data.action;
    return handleDispatch(action, data);
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function handleDispatch(action, data) {
  if (action === 'clockIn') {
    return handleClockIn(data);
  } else if (action === 'clockOut') {
    return handleClockOut(data);
  } else if (action === 'manualRecord') {
    return handleManualRecord(data);
  }
  return responseJSON({ status: 'error', message: '未知操作' });
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('打卡紀錄');
  if (!sheet) {
    sheet = ss.insertSheet('打卡紀錄');
    sheet.appendRow([
      '紀錄編號', '員工姓名', '日期', '上班時間', '下班時間', '總工時(小時)', '時薪(NTD)', '當日薪資(NTD)', '備註/型態', '打卡時間戳記'
    ]);
  }
  return sheet;
}

function handleClockIn(data) {
  var sheet = getSheet();
  var now = new Date();
  var dateStr = data.date || Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  var timeStr = data.time || Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm');
  var employee = data.employeeName || '員工';
  var hourlyWage = Number(data.hourlyWage) || 200;
  var note = data.note || '即時上班打卡';
  var recordId = 'REC-' + new Date().getTime();

  sheet.appendRow([
    recordId, employee, dateStr, timeStr, '', 0, hourlyWage, 0, note,
    Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
  ]);

  return responseJSON({
    status: 'success',
    message: '上班打卡成功！',
    record: { recordId: recordId, employee: employee, date: dateStr, time: timeStr }
  });
}

function handleClockOut(data) {
  var sheet = getSheet();
  var now = new Date();
  var dateStr = data.date || Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  var timeStr = data.time || Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm');
  var employee = data.employeeName || '員工';
  var note = data.note || '即時下班打卡';

  var values = sheet.getDataRange().getValues();
  var targetRow = -1;

  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][1] === employee && values[i][2] === dateStr && (!values[i][4] || values[i][4] === '')) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return responseJSON({
      status: 'error',
      message: '找不到您今日 (' + dateStr + ') 的上班打卡紀錄，請使用補登模式！'
    });
  }

  var rowData = values[targetRow - 1];
  var startTimeStr = rowData[3];
  var hourlyWage = Number(rowData[6]) || Number(data.hourlyWage) || 200;
  var hours = calculateHours(startTimeStr, timeStr);
  var salary = Math.round(hours * hourlyWage);

  sheet.getRange(targetRow, 5).setValue(timeStr);
  sheet.getRange(targetRow, 6).setValue(hours);
  sheet.getRange(targetRow, 8).setValue(salary);
  sheet.getRange(targetRow, 9).setValue(note);

  return responseJSON({
    status: 'success',
    message: '下班打卡成功！今日工時：' + hours + ' 小時，應得薪資：NT$ ' + salary,
    record: { employee: employee, date: dateStr, startTime: startTimeStr, endTime: timeStr, hours: hours, salary: salary }
  });
}

function handleManualRecord(data) {
  var sheet = getSheet();
  var employee = data.employeeName || '員工';
  var dateStr = data.date;
  var startTimeStr = data.startTime;
  var endTimeStr = data.endTime;
  var hourlyWage = Number(data.hourlyWage) || 200;
  var note = data.note || '自訂補登紀錄';

  var hours = calculateHours(startTimeStr, endTimeStr);
  var salary = Math.round(hours * hourlyWage);
  var now = new Date();
  var recordId = 'REC-' + new Date().getTime();

  sheet.appendRow([
    recordId, employee, dateStr, startTimeStr, endTimeStr, hours, hourlyWage, salary, note,
    Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
  ]);

  return responseJSON({
    status: 'success',
    message: '自訂打卡紀錄已儲存！工時：' + hours + ' 小時，薪資：NT$ ' + salary,
    record: { recordId: recordId, employee: employee, date: dateStr, hours: hours, salary: salary }
  });
}

function handleGetRecords(employeeName) {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var records = [];

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!employeeName || r[1] === employeeName) {
      records.push({
        id: r[0],
        employee: r[1],
        date: Utilities.formatDate(new Date(r[2]), 'Asia/Taipei', 'yyyy-MM-dd'),
        startTime: r[3],
        endTime: r[4] || '未打卡',
        hours: Number(r[5]) || 0,
        hourlyWage: Number(r[6]) || 0,
        salary: Number(r[7]) || 0,
        note: r[8] || ''
      });
    }
  }

  return responseJSON({
    status: 'success',
    records: records.reverse()
  });
}

function calculateHours(start, end) {
  if (!start || !end) return 0;
  var pStart = start.split(':');
  var pEnd = end.split(':');
  var sMinutes = parseInt(pStart[0], 10) * 60 + parseInt(pStart[1], 10);
  var eMinutes = parseInt(pEnd[0], 10) * 60 + parseInt(pEnd[1], 10);
  if (eMinutes < sMinutes) {
    eMinutes += 24 * 60;
  }
  var diffMinutes = eMinutes - sMinutes;
  return parseFloat((diffMinutes / 60).toFixed(2));
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
