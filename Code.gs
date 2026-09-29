/**
 * =========================================================================
 * 原資增能獎勵報到系統 (click_fjuirc) - Google Apps Script (GAS) 後端核心
 * 試算表 ID: 1up0RNU638zVvAtLQlmcW3zXDuNRl5aUxV1VVUcCSu5Y
 * 工作表 GID: 292620472
 * 
 * 欄位結構 (共 14 欄)：
 * A (1, idx 0):  學校名稱
 * B (2, idx 1):  姓名
 * C (3, idx 2):  職稱
 * D (4, idx 3):  身分證字號 (查詢與掃描比對鍵值)
 * E (5, idx 4):  聯絡電話
 * F (6, idx 5):  飲食習慣
 * G (7, idx 6):  去程
 * H (8, idx 7):  回程
 * I (9, idx 8):  備註欄
 * J (10, idx 9): 簽到 (簽到時間寫入欄位)
 * K (11, idx 10): 簽退 (簽退時間寫入欄位)
 * L (12, idx 11): 時數
 * M (13, idx 12): 備註
 * N (14, idx 13): 組別
 * 
 * 系統稽核：
 * 自動維護「系統Log表」，完整記錄簽到、簽退、重複掃描與臨時報到之軌跡
 * =========================================================================
 */

const SPREADSHEET_ID = "1up0RNU638zVvAtLQlmcW3zXDuNRl5aUxV1VVUcCSu5Y";
const TARGET_GID = 292620472;
const SHEET_LOG = "系統Log表";

/**
 * 取得目標工作表 (優先比對 GID，找不到則取第一張工作表)
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === TARGET_GID) {
      return sheets[i];
    }
  }
  return sheets[0];
}

/**
 * 寫入系統稽核 Log 表 (若工作表不存在自動建立)
 */
function writeLog(actionType, name, idNumber, school, groupName, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(SHEET_LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOG);
      logSheet.appendRow([
        "紀錄時間",
        "動作類型",
        "學員姓名",
        "身分證字號",
        "學校名稱",
        "組別",
        "詳細結果說明"
      ]);
      logSheet.setFrozenRows(1);
    }
    const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([
      nowStr,
      actionType || "",
      name || "-",
      idNumber || "-",
      school || "-",
      groupName || "-",
      details || ""
    ]);
  } catch (err) {
    console.warn("寫入 Log 異常: " + err.toString());
  }
}

/**
 * 處理 POST 請求 (主要 API 進入點)
 */
function doPost(e) {
  try {
    const contents = e.postData ? e.postData.contents : "{}";
    const request = JSON.parse(contents);
    const action = request.action;

    if (action === "processScan") {
      return jsonResponse(processScan(request.idNumber, request.scanType));
    } else if (action === "getAdminData") {
      return jsonResponse(getAdminData());
    } else if (action === "addTempUser") {
      return jsonResponse(addTempUser(request.data));
    }

    return jsonResponse({ success: false, error: "未知的請求動作 (Unknown action)" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

/**
 * 處理 GET 請求 (提供測試或備用取得資料)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === "getAdminData") {
      return jsonResponse(getAdminData());
    } else if (action === "processScan") {
      return jsonResponse(processScan(e.parameter.idNumber, e.parameter.scanType));
    }

    return jsonResponse({
      status: "online",
      message: "原資報到系統 (click_fjuirc) GAS API 正常運行中 (已啟用 Log 稽核記錄)",
      spreadsheetId: SPREADSHEET_ID,
      targetGid: TARGET_GID
    });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

/**
 * 時間日期格式化輔助函式
 */
function formatDateTime(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  }
  return val.toString().trim();
}

/**
 * 執行報到掃描 (支援 簽到 checkIn 與 簽退 checkOut，自動留存 Log)
 * @param {string} idNumber 身分證字號
 * @param {string} scanType 'checkIn' 或 'checkOut' (預設為 'checkIn')
 */
function processScan(idNumber, scanType) {
  if (!idNumber) {
    return { status: "not_found", message: "身分證字號不可為空" };
  }

  const type = (scanType === "checkOut") ? "checkOut" : "checkIn";
  const cleanId = idNumber.toString().trim().toUpperCase();
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();

  // 第一列為表頭 (Row 0)，從第二列開始搜尋 (i = 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // D 欄為身分證字號 (Index 3)
    const targetId = row[3] ? row[3].toString().trim().toUpperCase() : "";

    if (targetId === cleanId) {
      const school = row[0] ? row[0].toString().trim() : "";      // A: 學校名稱
      const name = row[1] ? row[1].toString().trim() : "";        // B: 姓名
      const title = row[2] ? row[2].toString().trim() : "";       // C: 職稱
      const phone = row[4] ? row[4].toString().trim() : "";       // E: 聯絡電話
      const diet = row[5] ? row[5].toString().trim() : "";        // F: 飲食習慣
      const depart = row[6] ? row[6].toString().trim() : "";      // G: 去程
      const returnTrip = row[7] ? row[7].toString().trim() : "";  // H: 回程
      const note1 = row[8] ? row[8].toString().trim() : "";       // I: 備註欄
      let checkInTime = formatDateTime(row[9]);                   // J: 簽到 (Col 10)
      let checkOutTime = formatDateTime(row[10]);                 // K: 簽退 (Col 11)
      const hours = row[11] ? row[11].toString().trim() : "";     // L: 時數
      const note2 = row[12] ? row[12].toString().trim() : "";     // M: 備註
      const groupName = row[13] ? row[13].toString().trim() : ""; // N: 組別

      const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

      if (type === "checkIn") {
        // 簽到處理
        if (checkInTime !== "") {
          writeLog("重複簽到", name, cleanId, school, groupName, `此學員已於 ${checkInTime} 簽到過`);
          return {
            status: "already",
            scanType: "checkIn",
            name: name,
            school: school,
            title: title,
            groupName: groupName,
            checkInTime: checkInTime,
            checkOutTime: checkOutTime,
            hours: hours
          };
        }

        // 首次簽到：寫入 J 欄 (第 10 欄，列號為 i + 1)
        sheet.getRange(i + 1, 10).setValue(nowStr);
        SpreadsheetApp.flush();
        writeLog("簽到成功", name, cleanId, school, groupName, "完成大會簽到手續");

        return {
          status: "success",
          scanType: "checkIn",
          name: name,
          school: school,
          title: title,
          groupName: groupName,
          checkInTime: nowStr,
          checkOutTime: checkOutTime,
          hours: hours
        };

      } else {
        // 簽退處理 (checkOut)
        if (checkOutTime !== "") {
          writeLog("重複簽退", name, cleanId, school, groupName, `此學員已於 ${checkOutTime} 簽退過`);
          return {
            status: "already",
            scanType: "checkOut",
            name: name,
            school: school,
            title: title,
            groupName: groupName,
            checkInTime: checkInTime,
            checkOutTime: checkOutTime,
            hours: hours
          };
        }

        // 首次簽退：寫入 K 欄 (第 11 欄，列號為 i + 1)
        sheet.getRange(i + 1, 11).setValue(nowStr);
        SpreadsheetApp.flush();
        writeLog("簽退成功", name, cleanId, school, groupName, "完成大會簽退手續");

        return {
          status: "success",
          scanType: "checkOut",
          name: name,
          school: school,
          title: title,
          groupName: groupName,
          checkInTime: checkInTime,
          checkOutTime: nowStr,
          hours: hours
        };
      }
    }
  }

  // 找不到該身分證
  writeLog("查無資料", "未知", cleanId, "-", "-", "名冊中查無此身分證字號");
  return { status: "not_found" };
}

/**
 * 新增現場臨時報到 (自動完成即時簽到，並留存 Log)
 */
function addTempUser(data) {
  if (!data || !data.idNumber || !data.name) {
    return { success: false, message: "「姓名」與「身分證字號」為必填欄位！" };
  }

  const cleanId = data.idNumber.toString().trim().toUpperCase();
  const sheet = getTargetSheet();
  const values = sheet.getDataRange().getValues();

  // 檢查是否已有名單存在此身分證
  for (let i = 1; i < values.length; i++) {
    const rowId = values[i][3] ? values[i][3].toString().trim().toUpperCase() : "";
    if (rowId === cleanId) {
      const existingName = values[i][1] ? values[i][1].toString().trim() : "";
      let checkInTime = formatDateTime(values[i][9]);
      if (checkInTime) {
        writeLog("臨時報到(重複)", existingName, cleanId, "-", "-", `已在名冊中且已於 ${checkInTime} 簽到過`);
        return {
          success: false,
          status: "already_checked_in",
          message: `此身分證已在名冊中（${existingName}），且已於 ${checkInTime} 完成簽到！`
        };
      } else {
        // 在名冊中但尚未簽到：直接為其補簽到
        const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 1, 10).setValue(nowStr);
        SpreadsheetApp.flush();
        writeLog("臨時報到(補簽)", existingName, cleanId, "-", "-", "原在名冊中尚未簽到，現場加簽補完成簽到");
        return {
          success: true,
          status: "existing_checked_in",
          message: `此身分證已在名冊中（${existingName}），已直接為其完成簽到！`,
          name: existingName,
          checkInTime: nowStr
        };
      }
    }
  }

  // 首次臨時報到：新增一列
  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  const school = (data.school || "").toString().trim();
  const name = (data.name || "").toString().trim();
  const title = (data.title || "").toString().trim();
  const phone = (data.phone || "").toString().trim();
  const diet = (data.diet || "").toString().trim();
  const depart = (data.depart || "").toString().trim();
  const returnTrip = (data.returnTrip || "").toString().trim();
  const note1 = (data.note1 || "").toString().trim();
  const checkIn = nowStr;
  const checkOut = "";
  const hours = (data.hours || "").toString().trim();
  let note2 = (data.note2 || "").toString().trim();
  note2 = note2 ? `[現場臨時報到] ${note2}` : "[現場臨時報到]";
  const groupName = (data.groupName || "臨時組").toString().trim();

  // 寫入 14 欄
  const newRow = [
    school,      // A (1): 學校名稱
    name,        // B (2): 姓名
    title,       // C (3): 職稱
    cleanId,     // D (4): 身分證字號
    phone,       // E (5): 聯絡電話
    diet,        // F (6): 飲食習慣
    depart,      // G (7): 去程
    returnTrip,  // H (8): 回程
    note1,       // I (9): 備註欄
    checkIn,     // J (10): 簽到
    checkOut,    // K (11): 簽退
    hours,       // L (12): 時數
    note2,       // M (13): 備註
    groupName    // N (14): 組別
  ];

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  writeLog("現場臨時報到", name, cleanId, school, groupName, "現場新增名額並立即簽到");

  return {
    success: true,
    status: "created",
    message: "現場臨時報到成功，並已完成即時簽到！",
    name: name,
    school: school,
    groupName: groupName,
    checkInTime: nowStr
  };
}

/**
 * 取得完整名單數據 (提供看板查詢)
 */
function getAdminData() {
  try {
    const sheet = getTargetSheet();
    const rawData = sheet.getDataRange().getValues();

    if (rawData.length <= 1) {
      return { success: false, error: "試算表中無資料 (只有標題列或為空)" };
    }

    const dataList = [];
    // 從第 2 列開始 (排除標題列)
    for (let i = 1; i < rawData.length; i++) {
      const r = rawData[i];
      // 忽略全空列 (判斷姓名與身分證皆空則略過)
      if (!r[1] && !r[3]) continue;

      dataList.push([
        r[0] ? r[0].toString().trim() : "",               // 0: 學校名稱
        r[1] ? r[1].toString().trim() : "",               // 1: 姓名
        r[2] ? r[2].toString().trim() : "",               // 2: 職稱
        r[3] ? r[3].toString().toUpperCase().trim() : "", // 3: 身分證字號
        r[4] ? r[4].toString().trim() : "",               // 4: 聯絡電話
        r[5] ? r[5].toString().trim() : "",               // 5: 飲食習慣
        r[6] ? r[6].toString().trim() : "",               // 6: 去程
        r[7] ? r[7].toString().trim() : "",               // 7: 回程
        r[8] ? r[8].toString().trim() : "",               // 8: 備註欄
        formatDateTime(r[9]),                             // 9: 簽到
        formatDateTime(r[10]),                            // 10: 簽退
        r[11] ? r[11].toString().trim() : "",             // 11: 時數
        r[12] ? r[12].toString().trim() : "",             // 12: 備註
        r[13] ? r[13].toString().trim() : ""              // 13: 組別
      ]);
    }

    return { success: true, data: dataList };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 輔助函式：產生 JSON 回應
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
