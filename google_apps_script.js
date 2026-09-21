// ============================================================
// SCMP Pilot Survey — Google Apps Script 后端
// ============================================================
// 部署步骤：
// 1. 打开 Google Sheets，新建一个空白表格
// 2. 在表格底部新建 3 个工作表（sheet tab），分别命名为：
//    - pretest
//    - diary
//    - participants
// 3. 点击菜单 Extensions → Apps Script
// 4. 把下面全部代码粘贴进去，替换原有内容
// 5. 点击 Deploy → New deployment
//    - 类型选 "Web app"
//    - Execute as: Me
//    - Who has access: Anyone
// 6. 点击 Deploy，复制生成的 URL
// 7. 把 URL 粘贴到 survey HTML 文件顶部的 API_URL 变量里
// ============================================================

function doPost(e) {
  try {
    var raw = "";
    // Handle form submission (payload parameter)
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }
    // Handle direct JSON body
    else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.type === "pretest") {
      var sheet = ss.getSheetByName("pretest");
      if (!sheet) { sheet = ss.insertSheet("pretest"); }
      
      // 第一次写入时添加表头
      if (sheet.getLastRow() === 0) {
        var headers = ["timestamp", "pid", "group", 
          "P1", "P2", "P3", "P4", "P5", "P6", 
          "P7", "P8", "P9", "P10", "P11"];
        sheet.appendRow(headers);
      }
      
      var a = data.answers;
      var row = [
        new Date().toISOString(),
        data.pid,
        data.group,
        a.P1 || "", a.P2 || "", a.P3 || "", a.P4 || "",
        Array.isArray(a.P5) ? a.P5.join("; ") : (a.P5 || ""),
        Array.isArray(a.P6) ? a.P6.join("; ") : (a.P6 || ""),
        a.P7 || "", a.P8 || "", a.P9 || "", a.P10 || "", a.P11 || ""
      ];
      sheet.appendRow(row);
      
      // 更新 participants 表
      updateParticipant(ss, data.pid, data.group, "pretest");
      
    } else if (data.type === "diary") {
      var sheet = ss.getSheetByName("diary");
      if (!sheet) { sheet = ss.insertSheet("diary"); }
      
      if (sheet.getLastRow() === 0) {
        var headers = ["timestamp", "pid", "group", "day",
          "D1", "D2", "D3", "D4", "D5", "D6", "D7",
          "D8a", "D8b", "D8c", "D8d", "D8e",
          "D9", "D10", "D11", "D11_source",
          "N1", "N2", "N3", "N4", "N5", "N6"];
        sheet.appendRow(headers);
      }
      
      var a = data.answers;
      var row = [
        new Date().toISOString(),
        data.pid,
        data.group,
        data.day,
        a.D1 || "", a.D2 || "", a.D3 || "", a.D4 || "",
        a.D5 || "", a.D6 || "", a.D7 || "",
        a.D8a || "", a.D8b || "", a.D8c || "", a.D8d || "", a.D8e || "",
        a.D9 || "", a.D10 || "", a.D11 || "", a.D11_source || "",
        a.N1 || "", a.N2 || "", a.N3 || "", a.N4 || "", a.N5 || "", a.N6 || ""
      ];
      sheet.appendRow(row);
      
      updateParticipant(ss, data.pid, data.group, "day" + data.day);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e.parameter.action || "all";
    var result = {};
    
    if (action === "all" || action === "pretest") {
      result.pretest = sheetToJSON(ss.getSheetByName("pretest"));
    }
    if (action === "all" || action === "diary") {
      result.diary = sheetToJSON(ss.getSheetByName("diary"));
    }
    if (action === "all" || action === "participants") {
      result.participants = sheetToJSON(ss.getSheetByName("participants"));
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sheetToJSON(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

function updateParticipant(ss, pid, group, completed) {
  var sheet = ss.getSheetByName("participants");
  if (!sheet) {
    sheet = ss.insertSheet("participants");
    sheet.appendRow(["pid", "group", "completed", "lastActive"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === pid) {
      var existing = data[i][2] ? data[i][2].toString() : "";
      var items = existing ? existing.split(", ") : [];
      if (items.indexOf(completed) === -1) items.push(completed);
      sheet.getRange(i + 1, 3).setValue(items.join(", "));
      sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([pid, group, completed, new Date().toISOString()]);
  }
}
