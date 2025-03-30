/**************************
 * 全域變數設定
 **************************/
const TARGET_SHEET_ID = "13YlIspcyWnmwtKu3aXdBq29wAvCvyOpOFlMNEAJNgpc"; // 派案總表 ID
const VISIT_RECORD_TEMPLATE_ID = "1jt8cvHDl66yOWN7SUjKODufsLqy72rHyXgOGoDhIITU"; // 訪視紀錄表模板 ID
const DRIVE_FOLDER_ID = "1r43CWOrbpY6q8_CNruz9lVSFOXgRulmE"; // 根資料夾 ID
const VISIT_HOURS_TEMPLATE_ID = "1LXt49lAOiAQNuSgnMvafYz2DqVzo5RxESOzEp9rjyhM"; // 訪視時數表模板 ID
const MODIFICATION_LOG_SHEET_ID = "1sOW3iKA_-P-rVlnBsM0jYydCPywIgASUbTXnK60raQk";
const TEMPLATE_SPREADSHEET_ID = "1xPGtnxoyCsqth6ETT2H1ib7UtYQHKjIIDqvr8yYHyIU";
const DATE_FORMAT = "yyyy年MM月dd日 EEEE a hh:mm"; // 確保全中文格式

/**************************
 * 初始化試算表結構
 **************************/
function initializeSpreadsheet() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  if (!spreadsheet) {
    Logger.log("🚨 無法開啟試算表，請檢查 TARGET_SHEET_ID 是否正確");
    return;
  }

  // 初始化「負責人基本資料」工作表
  let ownerSheet = spreadsheet.getSheetByName("負責人基本資料");
  if (!ownerSheet) {
    ownerSheet = spreadsheet.insertSheet("負責人基本資料");
    const headers = ["負責人姓名", "Email", "聯絡電話", "備註"];
    ownerSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#d9e8f5")
      .setHorizontalAlignment("center");
    Logger.log("ℍ 已創建「負責人基本資料」工作表");
  }

  // 初始化「派案總表」
  let caseSheet = spreadsheet.getSheetByName("派案總表");
  if (!caseSheet) {
    caseSheet = spreadsheet.insertSheet("派案總表");
    Logger.log("ℍ 已創建「派案總表」工作表");
  }
  setupCaseSheet(caseSheet);

  // 初始化「報告總表」並匯入歷年資料
  handleReportSheet(spreadsheet);
}

/**************************
 * 設定「派案總表」結構
 **************************/
function setupCaseSheet(caseSheet) {
  if (!caseSheet) {
    Logger.log("🚨 setupCaseSheet: caseSheet 未定義，無法設定結構");
    return;
  }

  const headers = caseSheet.getRange(1, 1, 1, caseSheet.getLastColumn()).getValues()[0];
  let continueServiceIndex = headers.indexOf("服務延續到下一年");
  let timestampIndex = headers.indexOf("延續勾選時間");

  if (continueServiceIndex === -1) {
    const lastCol = headers.length;
    caseSheet.getRange(1, lastCol + 1).setValue("服務延續到下一年")
      .setFontWeight("bold")
      .setBackground("#d9e8f5");
    continueServiceIndex = lastCol;
    Logger.log("ℍ 在「派案總表」新增「服務延續到下一年」欄");
  }
  if (timestampIndex === -1) {
    const lastCol = caseSheet.getLastColumn();
    caseSheet.getRange(1, lastCol + 1).setValue("延續勾選時間")
      .setFontWeight("bold")
      .setBackground("#d9e8f5");
    timestampIndex = lastCol;
    Logger.log("ℍ 在「派案總表」新增「延續勾選時間」欄");
  }

  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  if (!spreadsheet) {
    Logger.log("🚨 setupCaseSheet: 無法開啟試算表，請檢查 TARGET_SHEET_ID 是否正確");
    return;
  }
  handleCaseContinuation(spreadsheet, caseSheet, continueServiceIndex, timestampIndex);
}

/**************************
 * 處理跨年延續邏輯（派案總表）
 **************************/
function handleCaseContinuation(spreadsheet, caseSheet, continueServiceIndex, timestampIndex) {
  if (!spreadsheet) {
    Logger.log("🚨 handleCaseContinuation: spreadsheet 未定義，無法處理跨年延續");
    return;
  }
  if (!caseSheet) {
    Logger.log("🚨 handleCaseContinuation: caseSheet 未定義，無法處理跨年延續");
    return;
  }

  const currentYear = parseInt(Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy"));
  const data = caseSheet.getDataRange().getValues();
  const headers = data[0];
  const caseClosedIndex = headers.indexOf("結案");
  const caseNumberIndex = headers.indexOf("案號");

  if (caseClosedIndex === -1 || caseNumberIndex === -1) {
    Logger.log("🚨 「派案總表」缺少「結案」或「案號」欄位，無法處理跨年延續");
    return;
  }

  let casesToContinue = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const caseNumber = row[caseNumberIndex];
    const caseClosed = row[caseClosedIndex];
    const yearInCaseNumber = caseNumber ? parseInt(caseNumber.split("-")[0]) + 1911 : currentYear;

    if ((!caseClosed || caseClosed === "") && yearInCaseNumber < currentYear) {
      caseSheet.getRange(i + 1, continueServiceIndex + 1).setValue(true);
      if (!row[timestampIndex]) {
        caseSheet.getRange(i + 1, timestampIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT));
      }
      casesToContinue.push([...row]);
      Logger.log(`ℍ 標記案號 ${caseNumber} 為服務延續到下一年`);
    }
  }

  if (casesToContinue.length > 0) {
    Logger.log(`ℍ 準備轉移 ${casesToContinue.length} 筆案件到下一年度`);
    transferCasesToNextYear(spreadsheet, casesToContinue, headers);
  } else {
    Logger.log("ℍ 沒有需要轉移的案件");
  }
}

/**************************
 * 將未完成案件轉移到下一年度（使用模板生成新試算表）
 **************************/
function transferCasesToNextYear(spreadsheet, casesToContinue, headers) {
  if (!spreadsheet) {
    Logger.log("🚨 transferCasesToNextYear: spreadsheet 未定義，嘗試使用當前活動試算表");
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      Logger.log("🚨 transferCasesToNextYear: 無法獲取活動試算表，退出");
      return;
    }
  }

  if (!Array.isArray(headers) || headers.length === 0) {
    Logger.log("🚨 transferCasesToNextYear: headers 未定義或無效，無法繼續");
    return;
  }

  if (!Array.isArray(casesToContinue) || casesToContinue.length === 0) {
    Logger.log("🚨 transferCasesToNextYear: casesToContinue 未定義或無案件需要轉移");
    return;
  }

  const TEMPLATE_SPREADSHEET_ID = "1xPGtnxoyCsqth6ETT2H1ib7UtYQHKjIIDqvr8yYHyIU";
  const currentYear = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");
  const nextYear = (parseInt(currentYear) + 1).toString();
  const newSpreadsheetName = `${nextYear}派案總表`;

  try {
    // 獲取模板試算表的父資料夾
    const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
    const parentFolder = templateFile.getParents().next();
    const parentFolderId = parentFolder.getId();
    Logger.log(`ℍ 模板試算表父資料夾 ID: ${parentFolderId}`);

    // 檢查是否已存在同名試算表
    let newSpreadsheet;
    const existingFiles = DriveApp.getFilesByName(newSpreadsheetName);
    if (existingFiles.hasNext()) {
      newSpreadsheet = SpreadsheetApp.open(existingFiles.next());
      Logger.log(`📤 使用現有試算表: ${newSpreadsheetName}`);
    } else {
      // 複製模板並生成新試算表
      const newSpreadsheetId = copyScriptToNewSpreadsheet(TEMPLATE_SPREADSHEET_ID, newSpreadsheetName, parentFolderId);
      newSpreadsheet = SpreadsheetApp.openById(newSpreadsheetId);
      Logger.log(`📤 已創建新試算表: ${newSpreadsheetName}, ID: ${newSpreadsheetId}`);
    }

    const newCaseSheet = newSpreadsheet.getSheetByName("派案總表");
    if (!newCaseSheet) {
      Logger.log("🚨 新試算表中未找到「派案總表」工作表，退出");
      return;
    }
    newCaseSheet.clear();
    newCaseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const caseNumberIndex = headers.indexOf("案號");
    const continueServiceIndex = headers.indexOf("服務延續到下一年");

    if (caseNumberIndex === -1 || continueServiceIndex === -1) {
      Logger.log("🚨 transferCasesToNextYear: headers 缺少「案號」或「服務延續到下一年」欄位，無法處理案件");
      return;
    }

    let newRows = [];
    casesToContinue.forEach(row => {
      const caseNumber = row[caseNumberIndex];
      const rocYear = parseInt(nextYear) - 1911;
      const caseNumberParts = caseNumber.split("-");
      caseNumberParts[0] = rocYear.toString();
      row[caseNumberIndex] = caseNumberParts.join("-");
      row[continueServiceIndex] = false;
      newRows.push(row);
      Logger.log(`ℍ 將案號 ${caseNumber} 更新為 ${row[caseNumberIndex]} 並轉移到 ${newSpreadsheetName}`);
    });

    if (newRows.length > 0) {
      newCaseSheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
      Logger.log(`📤 已轉移 ${newRows.length} 筆未完成案件到新派案總表`);
    }

    // 設置觸發器
    setupReportTrigger(newSpreadsheet.getId());
    setupHourlyTrigger(newSpreadsheet.getId());

  } catch (error) {
    Logger.log(`🚨 transferCasesToNextYear 執行錯誤: ${error.message}`);
  }
}

/**************************
 * 將模板試算表複製到新試算表並設置腳本
 **************************/
function copyScriptToNewSpreadsheet(sourceTemplateId, targetSpreadsheetName, parentFolderId) {
  try {
    // 驗證 sourceTemplateId 是否有效
    if (!sourceTemplateId || typeof sourceTemplateId !== "string" || sourceTemplateId.length < 40) {
      throw new Error(`無效的模板試算表 ID: ${sourceTemplateId}`);
    }

    let sourceFile;
    try {
      sourceFile = DriveApp.getFileById(sourceTemplateId);
      Logger.log(`ℍ 成功獲取模板試算表: ${sourceFile.getName()}, ID: ${sourceTemplateId}`);
    } catch (e) {
      throw new Error(`無法存取模板試算表 ID ${sourceTemplateId}: ${e.message}`);
    }

    // 驗證父資料夾
    let parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(parentFolderId);
      Logger.log(`ℍ 成功獲取父資料夾: ${parentFolder.getName()}, ID: ${parentFolderId}`);
    } catch (e) {
      throw new Error(`無法存取父資料夾 ID ${parentFolderId}: ${e.message}`);
    }

    // 複製模板試算表
    const newFile = sourceFile.makeCopy(targetSpreadsheetName, parentFolder);
    const newSpreadsheetId = newFile.getId();
    Logger.log(`ℍ 已複製模板試算表到: ${targetSpreadsheetName}, ID: ${newSpreadsheetId}`);

    // 注意：腳本需手動複製
    Logger.log(`ℍ 注意：請手動將此腳本複製到新試算表 ${newSpreadsheetId} 的 Apps Script 編輯器中`);
    return newSpreadsheetId;
  } catch (error) {
    Logger.log(`🚨 copyScriptToNewSpreadsheet 執行錯誤: ${error.message}`);
    throw error;
  }
}

/**************************
 * 處理「報告總表」歷年資料匯入
 **************************/
function handleReportSheet(spreadsheet) {
  const currentYear = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");
  const nextYear = (parseInt(currentYear) + 1).toString();
  const nextYearReportSheetName = `報告總表_${nextYear}`;
  let nextYearReportSheet = spreadsheet.getSheetByName(nextYearReportSheetName);

  // 如果下一年度報告總表不存在，則創建
  if (!nextYearReportSheet) {
    nextYearReportSheet = spreadsheet.insertSheet(nextYearReportSheetName);
    Logger.log(`ℍ 創建下一年度報告總表: ${nextYearReportSheetName}`);
  }

  // 收集所有歷年報告總表的資料
  let allReportData = [];
  let unifiedHeaders = [];
  spreadsheet.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.startsWith("報告總表_") && sheetName !== nextYearReportSheetName) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      if (unifiedHeaders.length === 0) {
        unifiedHeaders = headers;
        nextYearReportSheet.getRange(1, 1, 1, headers.length).setValues([headers])
          .setFontWeight("bold")
          .setBackground("#d9e8f5");
      }
      allReportData = allReportData.concat(data.slice(1)); // 排除標題列
      Logger.log(`ℍ 收集 ${sheetName} 的資料，總計 ${data.length - 1} 筆`);
    }
  });

  // 將歷年資料寫入下一年度報告總表
  if (allReportData.length > 0) {
    const existingData = nextYearReportSheet.getDataRange().getValues();
    const startRow = existingData.length + 1;
    nextYearReportSheet.getRange(startRow, 1, allReportData.length, unifiedHeaders.length).setValues(allReportData);
    Logger.log(`ℍ 已將 ${allReportData.length} 筆歷年資料匯入 ${nextYearReportSheetName}`);
  }
}

/**************************
 * 合併後的 onEdit 函數
 **************************/
function onEdit(e) {
  Logger.log(`ℍ onEdit 已觸發，事件對象: ${JSON.stringify(e)}`);
  
  if (!e || !e.range) {
    Logger.log(`⚠ 事件對象無效，跳過執行`);
    return;
  }

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const value = e.value;

  Logger.log(`ℍ 編輯事件 - 工作表: ${sheetName}, 行: ${row}, 列: ${col}, 值: ${value}`);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(`ℍ 當前試算表 ID: ${spreadsheet.getId()}`);

  // 處理「派案總表」
  if (sheetName === "派案總表") {
    const caseSheet = spreadsheet.getSheetByName("派案總表");
    if (!caseSheet) {
      Logger.log("🚨 caseSheet 未找到，退出");
      return;
    }
    const headers = caseSheet.getRange(1, 1, 1, caseSheet.getLastColumn()).getValues()[0];
    Logger.log(`ℍ 標題列: ${headers}`);

    const continueServiceIndex = headers.indexOf("服務延續到下一年");
    const timestampIndex = headers.indexOf("延續勾選時間");
    const notifyIndex = headers.indexOf("已通知");
    const sentIndex = headers.indexOf("已寄送");

    Logger.log(`ℍ 欄位索引 - 已通知: ${notifyIndex}, 已寄送: ${sentIndex}, 服務延續: ${continueServiceIndex}`);

    // 處理「服務延續到下一年」
    if (col - 1 === continueServiceIndex && value === "TRUE" && row > 1) {
      Logger.log(`ℍ 檢測到「服務延續到下一年」勾選，行 ${row}`);
      const rowData = caseSheet.getRange(row, 1, 1, headers.length).getValues()[0];
      const timestampCell = caseSheet.getRange(row, timestampIndex + 1);
      if (!timestampCell.getValue()) {
        const timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT);
        timestampCell.setValue(timestamp);
        Logger.log(`ℍ 記錄延續勾選時間: ${timestamp}`);
      }
      transferCasesToNextYear(spreadsheet, [rowData], headers);
      Logger.log(`ℍ 完成行 ${row} 延續處理`);
    }

    // 處理「已通知」
    if (col - 1 === notifyIndex && value === "TRUE" && row > 1) {
      Logger.log(`ℍ 檢測到「已通知」勾選，行 ${row}`);
      const rowData = caseSheet.getRange(row, 1, 1, headers.length).getValues()[0];
      const checkTimeIndex = headers.indexOf("勾選時間");
      if (checkTimeIndex !== -1 && !caseSheet.getRange(row, checkTimeIndex + 1).getValue()) {
        const checkTime = Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT);
        caseSheet.getRange(row, checkTimeIndex + 1).setValue(checkTime);
        Logger.log(`ℍ 記錄勾選時間: ${checkTime}`);
      }
      processRow(row - 1, caseSheet, headers, rowData, false); // 明確傳遞 isSent = false
      Logger.log(`ℍ 完成行 ${row} 已通知處理`);
    } else if (col - 1 === notifyIndex) {
      Logger.log(`⚠ 「已通知」條件未滿足 - 值: ${value}, 行: ${row}`);
    }

    // 處理「已寄送」
    if (col - 1 === sentIndex && value === "TRUE" && row > 1) {
      Logger.log(`ℍ 檢測到「已寄送」勾選，行 ${row}`);
      const rowData = caseSheet.getRange(row, 1, 1, headers.length).getValues()[0];
      const timestampIndexLocal = headers.indexOf("寄送時間");
      if (timestampIndexLocal !== -1 && !caseSheet.getRange(row, timestampIndexLocal + 1).getValue()) {
        const sentTime = Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT);
        caseSheet.getRange(row, timestampIndexLocal + 1).setValue(sentTime);
        Logger.log(`ℍ 記錄寄送時間: ${sentTime}`);
      }
      processRow(row - 1, caseSheet, headers, rowData, true); // 明確傳遞 isSent = true
      Logger.log(`ℍ 完成行 ${row} 已寄送處理`);
    }
  }

  // 處理「報告總表」下拉選項變更
  if (sheetName === "報告總表") {
    Logger.log(`ℍ 檢測到「報告總表」編輯`);
    if (row === 2 && col >= 2 && col <= 6) {
      Logger.log(`ℍ 篩選條件變更 (行 ${row}, 列 ${col})，開始更新報告總表`);
      updateReportSummarySheet(e);
    } else {
      Logger.log(`⚠ 編輯不在篩選條件範圍 (行 ${row}, 列 ${col})，跳過更新`);
    }
  }
}

/**************************
 * onOpen 觸發器
 **************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('報告工具')
    .addItem('更新報告', 'manualUpdateReport')
    .addItem('初始化試算表', 'initializeSpreadsheet')
    .addItem('匯入歷年報告', 'handleReportSheet')
    .addToUi();
  initializeSpreadsheet();
}

function processRow(rowIndex, sheet, headers, rowData, isSent = false) {
  const indices = {
    owner: headers.findIndex(h => h.includes("負責人")),
    email: headers.findIndex(h => h.includes("Email")),
    remuneration: headers.findIndex(h => h.includes("業務報酬") && h.includes("單次")),
    caseType: headers.findIndex(h => h.includes("個案類型")),
    totalVisits: headers.findIndex(h => h.includes("總共要訪視次數")),
    caseNumber: headers.findIndex(h => h.includes("案號")),
    caseName: headers.findIndex(h => h.includes("個案姓名")),
    casePhone: headers.findIndex(h => h.includes("個案電話")),
    caseAddress: headers.findIndex(h => h.includes("個案住址")),
    transport: headers.findIndex(h => h.includes("交通費補助")),
    status: headers.findIndex(h => h.includes("狀態")),
    serviceDate: headers.findIndex(h => h.includes("已預約初訪日期及時間")),
    notify: headers.findIndex(h => h.includes("已通知")),
    timestamp: headers.findIndex(h => h.includes("寄送時間")),
    sent: headers.findIndex(h => h.includes("已寄送")),
    plannerLink: headers.findIndex(h => h.includes("規畫師雲端")),
    visitHours: headers.findIndex(h => h.includes("訪視時數表")),
    caseDelivery: headers.findIndex(h => h.includes("訪視記錄表")),
    checkTime: headers.findIndex(h => h.includes("勾選時間")),
    alreadyVisited: headers.findIndex(h => h.includes("已訪視次數")), // AB 欄，第 28 欄 (索引 27)
    ownerCaseCount: headers.findIndex(h => h.includes("負責人派案數")), // 現在在 H 欄
    remainingVisits: headers.findIndex(h => h.includes("剩餘訪視次數")) // AC 欄，第 29 欄 (索引 28)
  };

  Logger.log(`ℍ processRow 開始處理行 ${rowIndex + 1}, isSent: ${isSent}, rowData: ${rowData}`);

  // 若為「已通知」觸發，檢查是否已處理過
  if (!isSent) {
    const timestampNotEmpty = rowData[indices.timestamp];
    const alreadySent = rowData[indices.sent] === "已寄送" || rowData[indices.sent].toString().includes("已寄送");
    if (timestampNotEmpty || alreadySent) {
      Logger.log(`⚠ 行 ${rowIndex + 1} 已寄送過，跳過「已通知」處理`);
      return;
    }
  }

  const checkTime = new Date();
  if (!isSent && !rowData[indices.checkTime]) {
    sheet.getRange(rowIndex + 1, indices.checkTime + 1).setValue(checkTime);
    Logger.log(`ℍ 勾選時間記錄: ${Utilities.formatDate(checkTime, "Asia/Taipei", "yyyy年MM月dd日 HH:mm")}`);
  }

  const ownerName = String(rowData[indices.owner] || "").trim();
  const caseType = rowData[indices.caseType] || "";
  let caseNumber = rowData[indices.caseNumber];
  const caseName = rowData[indices.caseName] || "未提供姓名";
  const email = rowData[indices.email];
  const totalVisits = Number(rowData[indices.totalVisits]) || 0;
  const serviceDate = rowData[indices.serviceDate] ? new Date(rowData[indices.serviceDate]) : null;

  if (!ownerName || !caseName) {
    const errorMsg = `負責人或個案姓名缺失 (ownerName: ${ownerName}, caseName: ${caseName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日 HH:mm")})`);
    return;
  }

  let ownerCaseCount = 0;
  if (!caseNumber && ownerName) {
    const currentYear = "2025"; // 使用西元年 2025
    const ownerCode = ownerName.split("-")[0];
    const totalCases = sheet.getDataRange().getValues().slice(1, rowIndex + 1).filter(r => r[indices.owner] === ownerName).length;
    ownerCaseCount = totalCases;
    const caseSeq = String(ownerCaseCount).padStart(2, "0");
    const typeCode = caseType ? caseType.split("-")[0] : "";
    caseNumber = `${currentYear}-${ownerCode}-${caseSeq}${typeCode}`; // 案號使用 2025
    sheet.getRange(rowIndex + 1, indices.caseNumber + 1).setValue(caseNumber);
    Logger.log(`ℍ 自動生成案號: ${caseNumber}`);
  }

  if (ownerName && indices.ownerCaseCount !== -1) {
    sheet.getRange(rowIndex + 1, indices.ownerCaseCount + 1).setValue(ownerCaseCount); // 更新到 H 欄
    Logger.log(`ℍ 更新負責人派案數: ${ownerName} - ${ownerCaseCount}`);
  }

  let subFolderName;
  switch (caseType) {
    case "i":
      subFolderName = `${ownerName}-i機構`;
      break;
    case "if":
      subFolderName = `${ownerName}-if機構家屬`;
      break;
    default:
      subFolderName = `${ownerName}-${caseType || "未分類"}`;
      break;
  }

  // 負責人資料夾
  let ownerFolder = getDriveFolder(ownerName) || DriveApp.getFolderById(DRIVE_FOLDER_ID).createFolder(ownerName);
  if (!ownerFolder || !ownerFolder.getId) {
    const errorMsg = `ownerFolder 創建失敗 (ownerName: ${ownerName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日 HH:mm")})`);
    return;
  }

  // 新增年度派案資料層：2025派案資料
  const yearFolderName = "2025派案資料";
  let yearFolder = getDriveFolder(yearFolderName, ownerFolder) || ownerFolder.createFolder(yearFolderName);
  if (!yearFolder || !yearFolder.getId) {
    const errorMsg = `yearFolder 創建失敗 (yearFolderName: ${yearFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日 HH:mm")})`);
    return;
  }

  // 個案類型資料夾移至年度資料夾下
  let typeFolder = getDriveFolder(subFolderName, yearFolder) || yearFolder.createFolder(subFolderName);
  if (!typeFolder || !typeFolder.getId) {
    const errorMsg = `typeFolder 創建失敗 (subFolderName: ${subFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日 HH:mm")})`);
    return;
  }

  const caseFolderName = `${caseName}-${caseNumber}-${ownerName}`;
  let caseFolder = getDriveFolder(caseFolderName, typeFolder) || typeFolder.createFolder(caseFolderName);
  if (!caseFolder || !caseFolder.getId) {
    const errorMsg = `caseFolder 創建失敗 (caseFolderName: ${caseFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy年MM月dd日 HH:mm")})`);
    return;
  }
  Logger.log(`ℍ caseFolder 成功創建或獲取: ${caseFolder.getName()} (ID: ${caseFolder.getId()})`);

  const plannerLink = getDriveFolderUrl(ownerName, null);
  if (plannerLink !== "⚠ 找不到對應的資料夾" && plannerLink !== "⚠ 分享資料夾時發生錯誤") {
    sheet.getRange(rowIndex + 1, indices.plannerLink + 1).setFormula(`=HYPERLINK("${plannerLink}", "${ownerName}")`);
  }

  const serviceDateStr = serviceDate ? Utilities.formatDate(serviceDate, "Asia/Taipei", "yyyy年MM月dd日 HH:mm") : "無資料"; // 使用西元年格式

  const visitRecordLink = createAndShareCopy(ownerName, caseNumber, caseName, null, caseFolder, totalVisits, rowData, headers);
  const visitRecordId = getFileIdFromLink(visitRecordLink);
  sheet.getRange(rowIndex + 1, indices.caseDelivery + 1).setFormula(`=HYPERLINK("${visitRecordLink}", "${caseFolderName}")`);

  const visitHoursLink = createAndShareVisitHoursSheet(ownerName, null, caseNumber, caseName, totalVisits, visitRecordLink, caseFolder, rowData[indices.remuneration], rowData[indices.transport], rowIndex, null, visitRecordId);
  sheet.getRange(rowIndex + 1, indices.visitHours + 1).setFormula(`=HYPERLINK("${visitHoursLink}", "2025${ownerName}訪視總表")`); // 修改訪視時數表名稱

  const pdfFile = generatePDF(headers, rowData, serviceDateStr, plannerLink, ownerName, caseNumber, caseName, visitRecordLink, visitHoursLink);
  const pdfUrl = savePdfToDrive(pdfFile, ownerName, caseNumber, caseName, null, caseFolder);

  const visitHoursSheet = SpreadsheetApp.openByUrl(visitHoursLink);
  const caseSheet = visitHoursSheet.getSheetByName(`${caseName}-${caseNumber}`);
  if (caseSheet) {
    const totalVisitsNum = totalVisits > 0 ? totalVisits : 0;
    for (let j = 0; j < totalVisitsNum; j++) {
      caseSheet.getRange(j + 2, 11).setFormula(`=HYPERLINK("${pdfUrl}", "${ownerName}-${caseNumber}-案件報告.pdf")`);
    }
  }

  try {
    sendSummaryEmail({
      email: email,
      ownerName: ownerName,
      cases: [{
        caseNumber,
        caseName,
        pdfUrl,
        plannerLink,
        visitRecordLink,
        visitHoursLink,
      }]
    });
    const sentTime = new Date();
    sheet.getRange(rowIndex + 1, indices.timestamp + 1).setValue(sentTime);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`已寄送：${Utilities.formatDate(sentTime, "Asia/Taipei", "yyyy年MM月dd日 HH:mm")}`); // 使用西元年格式
    Logger.log(`📩 已發送總結 Email 給 ${email}，寄送時間: ${Utilities.formatDate(sentTime, "Asia/Taipei", "yyyy年MM月dd日 HH:mm")}`);
  } catch (error) {
    const errorTime = new Date();
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${error.message}（${Utilities.formatDate(errorTime, "Asia/Taipei", "yyyy年MM月dd日 HH:mm")}`); // 使用西元年格式
    Logger.log(`🚨 Email 寄送失敗 (${email}): ${error.message}，錯誤時間: ${Utilities.formatDate(errorTime, "Asia/Taipei", "yyyy年MM月dd日 HH:mm")}`);
  }
}

function testProcessRow() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("派案總表");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(2, 1, 1, headers.length).getValues()[0]; // 測試第 2 行
  processRow(1, sheet, headers, rowData, false); // rowIndex 從 0 開始
}

/**************************
 * 處理單行資料
 **************************/
function processRow(rowIndex, sheet, headers, rowData, isSent = false) {
  const indices = {
    owner: headers.findIndex(h => h.includes("負責人")),
    email: headers.findIndex(h => h.includes("Email")),
    remuneration: headers.findIndex(h => h.includes("業務報酬") && h.includes("單次")),
    caseType: headers.findIndex(h => h.includes("個案類型")),
    totalVisits: headers.findIndex(h => h.includes("總共要訪視次數")),
    caseNumber: headers.findIndex(h => h.includes("案號")),
    caseName: headers.findIndex(h => h.includes("個案姓名")),
    casePhone: headers.findIndex(h => h.includes("個案電話")),
    caseAddress: headers.findIndex(h => h.includes("個案住址")),
    transport: headers.findIndex(h => h.includes("交通費補助")),
    status: headers.findIndex(h => h.includes("狀態")),
    serviceDate: headers.findIndex(h => h.includes("已預約初訪日期及時間")),
    notify: headers.findIndex(h => h.includes("已通知")),
    timestamp: headers.findIndex(h => h.includes("寄送時間")),
    sent: headers.findIndex(h => h.includes("已寄送")),
    plannerLink: headers.findIndex(h => h.includes("規畫師雲端")),
    visitHours: headers.findIndex(h => h.includes("訪視時數表")),
    caseDelivery: headers.findIndex(h => h.includes("訪視記錄表")),
    checkTime: headers.findIndex(h => h.includes("勾選時間")),
    alreadyVisited: headers.findIndex(h => h.includes("已訪視次數")),
    ownerCaseCount: headers.findIndex(h => h.includes("負責人派案數")),
    remainingVisits: headers.findIndex(h => h.includes("剩餘訪視次數"))
  };

  Logger.log(`ℍ processRow 開始處理行 ${rowIndex + 1}, isSent: ${isSent}, rowData: ${rowData}`);

  // 若為「已通知」觸發，檢查是否已處理過
  if (!isSent) {
    const timestampNotEmpty = rowData[indices.timestamp];
    const alreadySent = rowData[indices.sent] === "已寄送" || rowData[indices.sent].toString().includes("已寄送");
    if (timestampNotEmpty || alreadySent) {
      Logger.log(`⚠ 行 ${rowIndex + 1} 已寄送過，跳過「已通知」處理`);
      return;
    }
  }

  const checkTime = new Date();
  if (!isSent && !rowData[indices.checkTime]) {
    sheet.getRange(rowIndex + 1, indices.checkTime + 1).setValue(checkTime);
    Logger.log(`ℍ 勾選時間記錄: ${Utilities.formatDate(checkTime, "Asia/Taipei", DATE_FORMAT)}`);
  }

  const ownerName = String(rowData[indices.owner] || "").trim();
  const caseType = rowData[indices.caseType] || "";
  let caseNumber = rowData[indices.caseNumber];
  const caseName = rowData[indices.caseName] || "未提供姓名";
  const email = rowData[indices.email];
  const totalVisits = Number(rowData[indices.totalVisits]) || 0;
  const serviceDate = rowData[indices.serviceDate] ? new Date(rowData[indices.serviceDate]) : null;

  if (!ownerName || !caseName) {
    const errorMsg = `負責人或個案姓名缺失 (ownerName: ${ownerName}, caseName: ${caseName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  let ownerCaseCount = 0;
  if (!caseNumber && ownerName) {
    const currentYear = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");
    const rocYear = String(parseInt(currentYear) - 1911);
    const ownerCode = ownerName.split("-")[0];
    const totalCases = sheet.getDataRange().getValues().slice(1, rowIndex + 1).filter(r => r[indices.owner] === ownerName).length;
    ownerCaseCount = totalCases;
    const caseSeq = String(ownerCaseCount).padStart(2, "0");
    const typeCode = caseType ? caseType.split("-")[0] : "";
    caseNumber = `${rocYear}-${ownerCode}-${caseSeq}${typeCode}`;
    sheet.getRange(rowIndex + 1, indices.caseNumber + 1).setValue(caseNumber);
    Logger.log(`ℍ 自動生成案號: ${caseNumber}`);
  }

  if (ownerName && indices.ownerCaseCount !== -1) {
    sheet.getRange(rowIndex + 1, indices.ownerCaseCount + 1).setValue(ownerCaseCount);
    Logger.log(`ℍ 更新負責人派案數: ${ownerName} - ${ownerCaseCount}`);
  }

  let subFolderName;
  switch (caseType) {
    case "i":
      subFolderName = `${ownerName}-i機構`;
      break;
    case "if":
      subFolderName = `${ownerName}-if機構家屬`;
      break;
    default:
      subFolderName = `${ownerName}-${caseType || "未分類"}`;
      break;
  }

  let ownerFolder = getDriveFolder(ownerName) || DriveApp.getFolderById(DRIVE_FOLDER_ID).createFolder(ownerName);
  if (!ownerFolder || !ownerFolder.getId) {
    const errorMsg = `ownerFolder 創建失敗 (ownerName: ${ownerName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  let typeFolder = getDriveFolder(subFolderName) || ownerFolder.createFolder(subFolderName);
  if (!typeFolder || !typeFolder.getId) {
    const errorMsg = `typeFolder 創建失敗 (subFolderName: ${subFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  const caseFolderName = `${caseName}-${caseNumber}-${ownerName}`;
  let caseFolder = getDriveFolder(caseFolderName) || typeFolder.createFolder(caseFolderName);
  if (!caseFolder || !caseFolder.getId) {
    const errorMsg = `caseFolder 創建失敗 (caseFolderName: ${caseFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }
  Logger.log(`ℍ caseFolder 成功創建或獲取: ${caseFolder.getName()} (ID: ${caseFolder.getId()})`);

  const plannerLink = getDriveFolderUrl(ownerName, null);
  if (plannerLink !== "⚠ 找不到對應的資料夾" && plannerLink !== "⚠ 分享資料夾時發生錯誤") {
    sheet.getRange(rowIndex + 1, indices.plannerLink + 1).setFormula(`=HYPERLINK("${plannerLink}", "${ownerName}")`);
  }

  const serviceDateStr = serviceDate ? Utilities.formatDate(serviceDate, "Asia/Taipei", DATE_FORMAT) : "無資料";

  const visitRecordLink = createAndShareCopy(ownerName, caseNumber, caseName, null, caseFolder, totalVisits, rowData, headers);
  const visitRecordId = getFileIdFromLink(visitRecordLink);
  sheet.getRange(rowIndex + 1, indices.caseDelivery + 1).setFormula(`=HYPERLINK("${visitRecordLink}", "${caseFolderName}")`);

  const visitHoursLink = createAndShareVisitHoursSheet(ownerName, null, caseNumber, caseName, totalVisits, visitRecordLink, caseFolder, rowData[indices.remuneration], rowData[indices.transport], rowIndex, null, visitRecordId);
  sheet.getRange(rowIndex + 1, indices.visitHours + 1).setFormula(`=HYPERLINK("${visitHoursLink}", "${ownerName} 訪視時數表")`);

  const pdfFile = generatePDF(headers, rowData, serviceDateStr, plannerLink, ownerName, caseNumber, caseName, visitRecordLink, visitHoursLink);
  const pdfUrl = savePdfToDrive(pdfFile, ownerName, caseNumber, caseName, null, caseFolder);

  const visitHoursSheet = SpreadsheetApp.openByUrl(visitHoursLink);
  const caseSheet = visitHoursSheet.getSheetByName(`${caseName}-${caseNumber}`);
  if (caseSheet) {
    const totalVisitsNum = totalVisits > 0 ? totalVisits : 0;
    for (let j = 0; j < totalVisitsNum; j++) {
      caseSheet.getRange(j + 2, 11).setFormula(`=HYPERLINK("${pdfUrl}", "${ownerName}-${caseNumber}-案件報告.pdf")`);
    }
  }

  try {
    sendSummaryEmail({
      email: email,
      ownerName: ownerName,
      cases: [{
        caseNumber,
        caseName,
        pdfUrl,
        plannerLink,
        visitRecordLink,
        visitHoursLink,
      }]
    });
    const sentTime = new Date();
    sheet.getRange(rowIndex + 1, indices.timestamp + 1).setValue(sentTime);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`已寄送：${Utilities.formatDate(sentTime, "Asia/Taipei", DATE_FORMAT)}`);
    Logger.log(`📩 已發送總結 Email 給 ${email}，寄送時間: ${Utilities.formatDate(sentTime, "Asia/Taipei", DATE_FORMAT)}`);
  } catch (error) {
    const errorTime = new Date();
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${error.message}（${Utilities.formatDate(errorTime, "Asia/Taipei", DATE_FORMAT)}）`);
    Logger.log(`🚨 Email 寄送失敗 (${email}): ${error.message}，錯誤時間: ${Utilities.formatDate(errorTime, "Asia/Taipei", DATE_FORMAT)}`);
  }
}

/**************************
 * 測試函數
 **************************/
function testInitialization() {
  initializeSpreadsheet();
  Logger.log("ℍ 測試初始化完成");
}

/**************************
 * 測試函數：驗證試算表存取
 **************************/
function testSpreadsheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById("13YlIspcyWnmwtKu3aXdBq29wAvCvyOpOFlMNEAJNgpc");
    Logger.log(`✅ 成功存取試算表: ${spreadsheet.getName()}`);
  } catch (error) {
    Logger.log(`🚨 存取試算表失敗: ${error.message}`);
  }
}

/**************************
 * 測試函數：檢查設置
 **************************/
function verifySetup() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  Logger.log(`試算表名稱: ${spreadsheet.getName()}`);
  const sheet = spreadsheet.getSheetByName("派案總表");
  Logger.log(`工作表狀態: ${sheet ? "存在" : "不存在"}`);
}

/**************************
 * delayedCheckAndSend：延遲處理函數
 **************************/
function delayedCheckAndSend() {
  Logger.log(`ℍ delayedCheckAndSend 開始執行`);
  Logger.log(`ℍ TARGET_SHEET_ID: ${TARGET_SHEET_ID}`);
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`ℍ 當前觸發器數量: ${triggers.length}`);
  for (let i = 0; i < triggers.length; i++) {
    const triggerId = triggers[i].getUniqueId();
    const row = PropertiesService.getScriptProperties().getProperty(triggerId);
    if (row) {
      const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
      if (!spreadsheet) {
        Logger.log(`🚨 無法打開試算表，TARGET_SHEET_ID 可能無效: ${TARGET_SHEET_ID}`);
        return;
      }
      const sheet = spreadsheet.getSheetByName("派案總表");
      if (!sheet) {
        Logger.log(`🚨 無法獲取工作表 '派案總表'，請確認工作表名稱`);
        return;
      }
      Logger.log(`ℍ 成功獲取工作表: ${sheet.getName()}`);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      checkAndSendEmailsWithPDFForRow(parseInt(row, 10) - 1, sheet, headers);
      ScriptApp.deleteTrigger(triggers[i]);
      PropertiesService.getScriptProperties().deleteProperty(triggerId);
      Logger.log(`ℍ 處理行 ${row} 並清理觸發器: ${triggerId}`);
      break;
    } else {
      Logger.log(`⚠ 觸發器 ${triggerId} 無對應的 row 屬性，跳過`);
    }
  }
}

function checkAndSendEmailsWithPDFForRow(rowIndex, sheet, headers) {
  if (!sheet) {
    Logger.log(`🚨 sheet 參數未定義，無法繼續執行`);
    return;
  }
  if (!headers) {
    Logger.log(`🚨 headers 參數未定義，無法繼續執行`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const row = data[rowIndex];
  const indices = {
    owner: headers.findIndex(h => h.includes("負責人")),
    email: headers.findIndex(h => h.includes("Email")),
    remuneration: headers.findIndex(h => h.includes("業務報酬") && h.includes("單次")),
    caseType: headers.findIndex(h => h.includes("個案類型")),
    totalVisits: headers.findIndex(h => h.includes("總共要訪視次數")),
    caseNumber: headers.findIndex(h => h.includes("案號")),
    caseName: headers.findIndex(h => h.includes("個案姓名")),
    casePhone: headers.findIndex(h => h.includes("個案電話")),
    caseAddress: headers.findIndex(h => h.includes("個案住址")),
    transport: headers.findIndex(h => h.includes("交通費補助")),
    status: headers.findIndex(h => h.includes("狀態")),
    serviceDate: headers.findIndex(h => h.includes("已預約初訪日期及時間")),
    notify: headers.findIndex(h => h.includes("已通知")),
    timestamp: headers.findIndex(h => h.includes("寄送時間")),
    sent: headers.findIndex(h => h.includes("已寄送")),
    plannerLink: headers.findIndex(h => h.includes("規畫師雲端")),
    visitHours: headers.findIndex(h => h.includes("訪視時數表")),
    caseDelivery: headers.findIndex(h => h.includes("訪視記錄表")),
    checkTime: headers.findIndex(h => h.includes("勾選時間")),
    alreadyVisited: headers.findIndex(h => h.includes("已訪視次數")),
    ownerCaseCount: headers.findIndex(h => h.includes("負責人派案數")), // 現在在 H 欄
  };

  let sentCount = 0;
  const emailsToSend = {};

  const isChecked = row[indices.notify] === true || row[indices.notify] === "TRUE";
  if (!isChecked) {
    Logger.log(`⚠ 行 ${rowIndex + 1} 的「已通知」未勾選，跳過`);
    return;
  }

  const checkTimeNotEmpty = row[indices.checkTime];
  const timestampNotEmpty = row[indices.timestamp];
  const alreadySent = row[indices.sent];
  if (checkTimeNotEmpty || timestampNotEmpty || alreadySent) {
    Logger.log(`⚠ 行 ${rowIndex + 1} 已處理過，跳過`);
    return;
  }

  const checkTime = new Date();
  sheet.getRange(rowIndex + 1, indices.checkTime + 1).setValue(checkTime);
  Logger.log(`ℍ 勾選時間記錄: ${Utilities.formatDate(checkTime, "Asia/Taipei", DATE_FORMAT)}`);

  const ownerName = String(row[indices.owner] || "").trim();
  const caseType = row[indices.caseType] || "";
  let caseNumber = row[indices.caseNumber];
  const caseName = row[indices.caseName] || "未提供姓名";
  const email = row[indices.email];
  const totalVisits = Number(row[indices.totalVisits]) || 0;
  const serviceDate = row[indices.serviceDate] ? new Date(row[indices.serviceDate]) : null;

  // 檢查必要數據
  if (!ownerName || !caseName) {
    const errorMsg = `負責人或個案姓名缺失 (ownerName: ${ownerName}, caseName: ${caseName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  let ownerCaseCount = 0;
  if (!caseNumber && ownerName) {
    const currentYear = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");
    const rocYear = String(parseInt(currentYear) - 1911);
    const ownerCode = ownerName.split("-")[0];
    const totalCases = data.slice(1, rowIndex + 1).filter(r => r[indices.owner] === ownerName).length;
    ownerCaseCount = totalCases;
    const caseSeq = String(ownerCaseCount).padStart(2, "0");
    const typeCode = caseType ? caseType.split("-")[0] : "";
    caseNumber = `${rocYear}-${ownerCode}-${caseSeq}${typeCode}`;
    sheet.getRange(rowIndex + 1, indices.caseNumber + 1).setValue(caseNumber);
    Logger.log(`ℍ 自動生成案號: ${caseNumber}`);
  }

  if (ownerName && indices.ownerCaseCount !== -1) {
    sheet.getRange(rowIndex + 1, indices.ownerCaseCount + 1).setValue(ownerCaseCount);
    Logger.log(`ℍ 更新負責人派案數: ${ownerName} - ${ownerCaseCount}`);
  }

  let subFolderName;
  switch (caseType) {
    case "i":
      subFolderName = `${ownerName}-i機構`;
      break;
    case "if":
      subFolderName = `${ownerName}-if機構家屬`;
      break;
    default:
      subFolderName = `${ownerName}-${caseType || "未分類"}`;
      break;
  }

  let ownerFolder = getDriveFolder(ownerName) || DriveApp.getFolderById(DRIVE_FOLDER_ID).createFolder(ownerName);
  if (!ownerFolder || !ownerFolder.getId) {
    const errorMsg = `ownerFolder 創建失敗 (ownerName: ${ownerName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  let typeFolder = getDriveFolder(subFolderName) || ownerFolder.createFolder(subFolderName);
  if (!typeFolder || !typeFolder.getId) {
    const errorMsg = `typeFolder 創建失敗 (subFolderName: ${subFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }

  const caseFolderName = `${caseName}-${caseNumber}-${ownerName}`;
  let caseFolder = getDriveFolder(caseFolderName) || typeFolder.createFolder(caseFolderName);
  if (!caseFolder || !caseFolder.getId) {
    const errorMsg = `caseFolder 創建失敗 (caseFolderName: ${caseFolderName})`;
    Logger.log(`🚨 行 ${rowIndex + 1} ${errorMsg}`);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${errorMsg}（${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}）`);
    return;
  }
  Logger.log(`ℍ caseFolder 成功創建或獲取: ${caseFolder.getName()} (ID: ${caseFolder.getId()})`);

  const plannerLink = getDriveFolderUrl(ownerName, null);
  if (plannerLink !== "⚠ 找不到對應的資料夾" && plannerLink !== "⚠ 分享資料夾時發生錯誤") {
    sheet.getRange(rowIndex + 1, indices.plannerLink + 1).setFormula(`=HYPERLINK("${plannerLink}", "${ownerName}")`);
  }

  const serviceDateStr = serviceDate ? Utilities.formatDate(serviceDate, "Asia/Taipei", DATE_FORMAT) : "無資料";

  try {
    const visitRecordLink = createAndShareCopy(ownerName, caseNumber, caseName, null, caseFolder, totalVisits, row, headers);
    const visitRecordId = getFileIdFromLink(visitRecordLink);
    sheet.getRange(rowIndex + 1, indices.caseDelivery + 1).setFormula(`=HYPERLINK("${visitRecordLink}", "${caseFolderName}")`);

    const visitHoursLink = createAndShareVisitHoursSheet(ownerName, null, caseNumber, caseName, totalVisits, visitRecordLink, caseFolder, row[indices.remuneration], row[indices.transport], rowIndex, null, visitRecordId);
    sheet.getRange(rowIndex + 1, indices.visitHours + 1).setFormula(`=HYPERLINK("${visitHoursLink}", "${ownerName} 訪視時數表")`);

    const pdfFile = generatePDF(headers, row, serviceDateStr, plannerLink, ownerName, caseNumber, caseName, visitRecordLink, visitHoursLink);
    const pdfUrl = savePdfToDrive(pdfFile, ownerName, caseNumber, caseName, null, caseFolder);

    const visitHoursSheet = SpreadsheetApp.openByUrl(visitHoursLink);
    const caseSheet = visitHoursSheet.getSheetByName(`${caseName}-${caseNumber}`);
    if (caseSheet) {
      const totalVisitsNum = totalVisits > 0 ? totalVisits : 0;
      for (let j = 0; j < totalVisitsNum; j++) {
        caseSheet.getRange(j + 2, 11).setFormula(`=HYPERLINK("${pdfUrl}", "${ownerName}-${caseNumber}-案件報告.pdf")`);
      }
    }

    emailsToSend[ownerName] = emailsToSend[ownerName] || { email, ownerName, cases: [] };
    emailsToSend[ownerName].cases.push({
      caseNumber,
      caseName,
      pdfUrl,
      plannerLink,
      visitRecordLink,
      visitHoursLink,
    });

    sendSummaryEmail(emailsToSend[ownerName]);
    const sentTime = new Date();
    sheet.getRange(rowIndex + 1, indices.timestamp + 1).setValue(sentTime);
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`已寄送：${Utilities.formatDate(sentTime, "Asia/Taipei", DATE_FORMAT)}`);
    Logger.log(`✅ 行 ${rowIndex + 1} 處理完成，已寄送 (案號 ${caseNumber})，寄送時間: ${Utilities.formatDate(sentTime, "Asia/Taipei", DATE_FORMAT)}`);
    sentCount++;
  } catch (error) {
    const errorTime = new Date();
    sheet.getRange(rowIndex + 1, indices.sent + 1).setValue(`錯誤：${error.message}（${Utilities.formatDate(errorTime, "Asia/Taipei", DATE_FORMAT)}）`);
    Logger.log(`🚨 行 ${rowIndex + 1} 處理失敗 (案號 ${caseNumber})：${error.message}，錯誤時間: ${Utilities.formatDate(errorTime, "Asia/Taipei", DATE_FORMAT)}`);
  }

  if (sentCount > 0) {
    Logger.log("📩 已發送總結 Email 給所有負責人");
  }
}

function checkAndSendEmailsWithPDF() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sheet = spreadsheet.getSheetByName("派案總表");
  if (!sheet) {
    Logger.log("⚠ 找不到「派案總表」，請確認名稱是否正確");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  Logger.log(`📌 找到的欄位名稱：${headers.join(", ")}`);

  const indices = {
    owner: headers.findIndex(h => h.includes("負責人")),
    email: headers.findIndex(h => h.includes("Email")),
    remuneration: headers.findIndex(h => h.includes("業務報酬") && h.includes("單次")),
    caseType: headers.findIndex(h => h.includes("個案類型")),
    totalVisits: headers.findIndex(h => h.includes("總共要訪視次數")),
    caseNumber: headers.findIndex(h => h.includes("案號")),
    caseName: headers.findIndex(h => h.includes("個案姓名")),
    casePhone: headers.findIndex(h => h.includes("個案電話")),
    caseAddress: headers.findIndex(h => h.includes("個案住址")),
    transport: headers.findIndex(h => h.includes("交通費補助")),
    status: headers.findIndex(h => h.includes("狀態")),
    serviceDate: headers.findIndex(h => h.includes("已預約初訪日期及時間")),
    notify: headers.findIndex(h => h.includes("已通知")),
    timestamp: headers.findIndex(h => h.includes("寄送時間")),
    sent: headers.findIndex(h => h.includes("已寄送")),
    plannerLink: headers.findIndex(h => h.includes("規畫師雲端")),
    visitHours: headers.findIndex(h => h.includes("訪視時數表")),
    caseDelivery: headers.findIndex(h => h.includes("訪視記錄表")),
    checkTime: headers.findIndex(h => h.includes("勾選時間")),
    alreadyVisited: headers.findIndex(h => h.includes("已訪視次數")),
    ownerCaseCount: headers.findIndex(h => h.includes("負責人派案數")), // 現在在 H 欄
  };

  let sentCount = 0;
  const emailsToSend = {};
  const currentTime = new Date();

  for (let i = 1; i < data.length && sentCount < 5; i++) {
    const row = data[i];
    const isChecked = row[indices.notify] === true || row[indices.notify] === "TRUE";
    if (!isChecked) continue;

    const checkTimeNotEmpty = row[indices.checkTime];
    const timestampNotEmpty = row[indices.timestamp];
    const alreadySent = row[indices.sent];
    if (checkTimeNotEmpty || timestampNotEmpty || alreadySent) continue;

    const ownerName = String(row[indices.owner] || "").trim();
    const caseType = row[indices.caseType] || "";
    let caseNumber = row[indices.caseNumber];
    const caseName = row[indices.caseName] || "未提供姓名";
    const email = row[indices.email];
    const totalVisits = Number(row[indices.totalVisits]) || 0;
    const serviceDate = row[indices.serviceDate] ? new Date(row[indices.serviceDate]) : null;

    let ownerCaseCount = 0;
    if (!caseNumber && ownerName) {
      const currentYear = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");
      const rocYear = String(parseInt(currentYear) - 1911);
      const ownerCode = ownerName.split("-")[0];
      const totalCases = data.slice(1, i + 1).filter(r => r[indices.owner] === ownerName).length;
      ownerCaseCount = totalCases;
      const caseSeq = String(ownerCaseCount).padStart(2, "0");
      const typeCode = caseType ? caseType.split("-")[0] : "";
      caseNumber = `${rocYear}-${ownerCode}-${caseSeq}${typeCode}`;
      sheet.getRange(i + 1, indices.caseNumber + 1).setValue(caseNumber);
      Logger.log(`ℍ 自動生成案號: ${caseNumber}`);
    }

    if (ownerName && indices.ownerCaseCount !== -1) {
      sheet.getRange(i + 1, indices.ownerCaseCount + 1).setValue(ownerCaseCount); // 更新到 H 欄
      Logger.log(`ℍ 更新負責人派案數: ${ownerName} - ${ownerCaseCount}`);
    }

    let subFolderName;
    switch (caseType) {
      case "i":
        subFolderName = `${ownerName}-i機構`;
        break;
      case "if":
        subFolderName = `${ownerName}-if機構家屬`;
        break;
      default:
        subFolderName = `${ownerName}-${caseType || "未分類"}`;
        break;
    }

    let ownerFolder = getDriveFolder(ownerName) || DriveApp.getFolderById(DRIVE_FOLDER_ID).createFolder(ownerName);
    let typeFolder = getDriveFolder(subFolderName) || ownerFolder.createFolder(subFolderName);
    const caseFolderName = `${caseName}-${caseNumber}-${ownerName}`;
    let caseFolder = getDriveFolder(caseFolderName) || typeFolder.createFolder(caseFolderName);

    const plannerLink = getDriveFolderUrl(ownerName, null);
    if (plannerLink !== "⚠ 找不到對應的資料夾" && plannerLink !== "⚠ 分享資料夾時發生錯誤") {
      sheet.getRange(i + 1, indices.plannerLink + 1).setFormula(`=HYPERLINK("${plannerLink}", "${ownerName}")`);
    }

    const serviceDateStr = serviceDate ? Utilities.formatDate(serviceDate, "Asia/Taipei", DATE_FORMAT) : "無資料";

    try {
      const visitRecordLink = createAndShareCopy(ownerName, caseNumber, caseName, null, caseFolder, totalVisits, row, headers);
      const visitRecordId = getFileIdFromLink(visitRecordLink);
      sheet.getRange(i + 1, indices.caseDelivery + 1).setFormula(`=HYPERLINK("${visitRecordLink}", "${caseFolderName}")`);

      const visitHoursLink = createAndShareVisitHoursSheet(ownerName, null, caseNumber, caseName, totalVisits, visitRecordLink, caseFolder, row[indices.remuneration], row[indices.transport], i - 1, null, visitRecordId); // remuneration 從 Z 欄取值
      sheet.getRange(i + 1, indices.visitHours + 1).setFormula(`=HYPERLINK("${visitHoursLink}", "${ownerName} 訪視時數表")`);

      const pdfFile = generatePDF(headers, row, serviceDateStr, plannerLink, ownerName, caseNumber, caseName, visitRecordLink, visitHoursLink);
      const pdfUrl = savePdfToDrive(pdfFile, ownerName, caseNumber, caseName, null, caseFolder);

      const visitHoursSheet = SpreadsheetApp.openByUrl(visitHoursLink);
      const caseSheet = visitHoursSheet.getSheetByName(`${caseName}-${caseNumber}`);
      if (caseSheet) {
        const totalVisitsNum = totalVisits > 0 ? totalVisits : 0;
        for (let j = 0; j < totalVisitsNum; j++) {
          caseSheet.getRange(j + 2, 12).setFormula(`=HYPERLINK("${pdfUrl}", "${ownerName}-${caseNumber}-案件報告.pdf")`); // 調整為第 12 欄 (L 欄)
        }
      }

      emailsToSend[ownerName] = emailsToSend[ownerName] || { email, ownerName, cases: [] };
      emailsToSend[ownerName].cases.push({
        caseNumber,
        caseName,
        pdfUrl,
        plannerLink,
        visitRecordLink,
        visitHoursLink,
      });

      const sentTime = Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT);
      sheet.getRange(i + 1, indices.timestamp + 1).setValue(new Date());
      sheet.getRange(i + 1, indices.sent + 1).setValue(`已寄送：${sentTime}`);
      sheet.getRange(i + 1, indices.checkTime + 1).setValue(currentTime);
      sentCount++;
      Logger.log(`✅ 行 ${i + 1} 處理完成，已寄送 (案號 ${caseNumber})`);
    } catch (error) {
      sheet.getRange(i + 1, indices.sent + 1).setValue(`錯誤：${error.message}`);
      Logger.log(`🚨 行 ${i + 1} 處理失敗 (案號 ${caseNumber})：${error.message}`);
    }
  }

  if (sentCount > 0) {
    Object.values(emailsToSend).forEach(sendSummaryEmail);
    Logger.log("📩 已發送總結 Email 給所有負責人");
    generateMonthlyRemunerationSheets();
    Logger.log("📊 已生成所有月份的報酬表");
  } else {
    Logger.log("⚠ 沒有需要發送的案件");
    generateMonthlyRemunerationSheets();
    Logger.log("📊 已更新所有月份的報酬表");
  }
}

/**************************
 * 發送總結 Email 給負責人
 **************************/
function sendSummaryEmail(ownerInfo) {
  const subject = `您的最新派案通知總結 - ${Utilities.formatDate(new Date(), "Asia/Taipei", DATE_FORMAT)}`;
  let htmlBody = `<p>親愛的 ${ownerInfo.ownerName}，</p><p>以下是您的新案件總結：</p><ul>`;
  ownerInfo.cases.forEach(caseInfo => {
    htmlBody += `<li>案件 (案號：${caseInfo.caseNumber})：
      <ul>
        <li><a href="${caseInfo.pdfUrl}">${ownerInfo.ownerName}-${caseInfo.caseNumber}-案件報告.pdf</a></li>
        <li><a href="${caseInfo.plannerLink}">${ownerInfo.ownerName}</a>（規畫師雲端）</li>
        <li><a href="${caseInfo.visitRecordLink}">${caseInfo.caseName}-${caseInfo.caseNumber}</a>（訪視記錄表）</li>
        <li><a href="${caseInfo.visitHoursLink}">${ownerInfo.ownerName} 訪視時數表</a></li>
      </ul>
    </li>`;
  });
  htmlBody += `</ul><p>請妥善保管上述個案資料，確保符合個人資料保護法規，感謝配合！</p><p>賽親派案系統</p>`;

  try {
    MailApp.sendEmail({ to: ownerInfo.email, subject, htmlBody });
    Logger.log(`📩 Email 已發送至 ${ownerInfo.email}`);
  } catch (error) {
    Logger.log(`🚨 發送 Email 失敗 (${ownerInfo.email}): ${error.message}`);
  }
}

/**************************
 * 負責人訪視時數表生成函數（單一試算表，個案分工作表）
 **************************/
function createAndShareVisitHoursSheet(ownerName, email, caseNumber, caseName, totalVisits, visitRecordLink, parentFolder, remuneration, transport, rowIndex, pdfUrl, visitRecordId) {
  Logger.log(`ℍ 開始處理 ${ownerName} 的訪視時數表，案號: ${caseNumber}`);

  if (!parentFolder || !parentFolder.getId) {
    Logger.log(`🚨 parentFolder 無效 (ID: ${parentFolder ? parentFolder.getId : 'null'}), 無法繼續`);
    return null;
  }

  let caseFolder = parentFolder;
  let typeFolder, ownerFolder;

  try {
    const parentIterator = caseFolder.getParents();
    if (parentIterator.hasNext()) {
      typeFolder = parentIterator.next();
      const ownerIterator = typeFolder.getParents();
      if (ownerIterator.hasNext()) {
        ownerFolder = ownerIterator.next();
      } else {
        Logger.log(`🚨 無法找到 ${typeFolder.getName()} 的父資料夾 (ownerFolder)，使用預設根資料夾`);
        ownerFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      }
    } else {
      Logger.log(`🚨 無法找到 ${caseFolder.getName()} 的父資料夾 (typeFolder)，使用預設根資料夾`);
      typeFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      ownerFolder = typeFolder;
    }
    Logger.log(`ℍ 成功獲取資料夾結構: 個案資料夾 (${caseFolder.getName()}) -> 類型資料夾 (${typeFolder.getName()}) -> 負責人資料夾 (${ownerFolder.getName()})`);
  } catch (error) {
    Logger.log(`🚨 獲取資料夾結構失敗: ${error.message}`);
    return null;
  }

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
    Logger.log(`ℍ 成功開啟試算表: ${spreadsheet.getName()}`);
  } catch (error) {
    Logger.log(`🚨 無法開啟試算表 (ID: ${TARGET_SHEET_ID}): ${error.message}`);
    return null;
  }

  const mainSheet = spreadsheet.getSheetByName("派案總表");
  if (!mainSheet) {
    Logger.log(`🚨 找不到派案總表`);
    return null;
  }

  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const caseTypeIndex = headers.findIndex(h => h.includes("個案類型"));
  const caseNumberIndex = headers.findIndex(h => h.includes("案號"));

  let effectiveRowIndex = rowIndex;
  if (rowIndex === null || isNaN(rowIndex)) {
    const dataRange = mainSheet.getDataRange().getValues();
    for (let i = 1; i < dataRange.length; i++) {
      const currentCaseNumber = dataRange[i][caseNumberIndex] || "";
      if (currentCaseNumber === caseNumber) {
        effectiveRowIndex = i - 1;
        Logger.log(`🚨 rowIndex 為 null，根據案號 ${caseNumber} 找到 effectiveRowIndex: ${effectiveRowIndex}`);
        break;
      }
    }
    if (effectiveRowIndex === null) {
      Logger.log(`🚨 無法根據案號 ${caseNumber} 找到有效 rowIndex，設為 0`);
      effectiveRowIndex = 0;
    }
  }

  const rowData = mainSheet.getRange(effectiveRowIndex + 1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  let caseType = caseTypeIndex !== -1 ? String(rowData[caseTypeIndex] || "未分類") : "未分類";
  caseType = caseType.trim();
  Logger.log(`ℍ 提取到的 caseType: ${caseType}`);

  const visitHoursSheetName = `${ownerName} 訪視時數表`;
  let totalSs;
  let existingSheetFile = null;

  try {
    const files = ownerFolder.getFilesByName(visitHoursSheetName);
    if (files.hasNext()) {
      existingSheetFile = files.next();
      totalSs = SpreadsheetApp.openById(existingSheetFile.getId());
      Logger.log(`ℍ 找到現有的訪視時數表: ${totalSs.getName()} (ID: ${totalSs.getId()})`);
    }
  } catch (error) {
    Logger.log(`🚨 檢查訪視時數表是否存在時失敗: ${error.message}`);
  }

  if (!totalSs) {
    try {
      const templateSheet = SpreadsheetApp.openById(VISIT_HOURS_TEMPLATE_ID);
      totalSs = templateSheet.copy(visitHoursSheetName);
      Logger.log(`ℍ 成功創建訪視時數表: ${totalSs.getName()}`);
    } catch (error) {
      Logger.log(`🚨 無法創建訪視時數表 (模板 ID: ${VISIT_HOURS_TEMPLATE_ID}): ${error.message}`);
      return null;
    }

    try {
      const file = DriveApp.getFileById(totalSs.getId());
      ownerFolder.addFile(file);
      const currentParents = file.getParents();
      while (currentParents.hasNext()) {
        const parent = currentParents.next();
        if (parent.getId() !== ownerFolder.getId()) {
          parent.removeFile(file);
          Logger.log(`ℍ 移除訪視時數表從舊資料夾: ${parent.getName()}`);
        }
      }
      Logger.log(`ℍ 成功移動訪視時數表到資料夾: ${ownerFolder.getName()}`);
    } catch (error) {
      Logger.log(`🚨 移動訪視時數表到資料夾失敗: ${error.message}`);
      return null;
    }
  }

  let descriptionSheet = totalSs.getSheetByName("說明");
  if (!descriptionSheet) {
    descriptionSheet = totalSs.insertSheet("說明", 0);
    descriptionSheet.getRange("A1").setValue("此為負責人訪視時數表說明頁面。\n- 總表：記錄所有個案的概要資訊。\n- 個案工作表：記錄具體訪視記錄。");
    descriptionSheet.setFrozenRows(1);
    Logger.log(`ℍ 創建「說明」工作表: ${descriptionSheet.getName()}`);
  }

  let overviewSheet = totalSs.getSheetByName("總表");
  if (!overviewSheet) {
    overviewSheet = totalSs.insertSheet("總表", 1);
    const overviewHeaders = ["個案連結", "總共訪視次數", "已訪視次數", "剩餘訪視次數", "備註", "結案"];
    overviewSheet.getRange(1, 1, 1, overviewHeaders.length).setValues([overviewHeaders])
      .setFontWeight("bold")
      .setBackground("#d9e8f5");
    Logger.log(`ℍ 創建「總表」工作表: ${overviewSheet.getName()}`);
  }

  totalSs.setActiveSheet(descriptionSheet);
  totalSs.moveActiveSheet(1);
  totalSs.setActiveSheet(overviewSheet);
  totalSs.moveActiveSheet(2);
  Logger.log(`ℍ 調整工作表順序: 說明 -> 總表 -> 個案工作表`);

  const caseSheetName = `${caseName}-${caseNumber}`;
  let totalCaseSheet;
  try {
    totalCaseSheet = totalSs.getSheetByName(caseSheetName);
    if (totalCaseSheet) {
      Logger.log(`ℍ 個案工作表已存在: ${caseSheetName}，將更新資料`);
    } else {
      totalCaseSheet = totalSs.insertSheet(caseSheetName);
      const headers = ["已訪視次數", "案號", "個案姓名", "服務日期", "服務時間幾點開始", "服務時間幾點結束", "訪視次數", "總共訪視次數", "剩餘訪視次數", "備註", "PDF 連結", "訪視記錄表連結", "結案"];
      totalCaseSheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold")
        .setBackground("#d9e8f5");
      Logger.log(`ℍ 創建個案工作表: ${caseSheetName}`);
    }
  } catch (error) {
    Logger.log(`🚨 創建或更新個案工作表失敗: ${error.message}`);
    return null;
  }

  const totalVisitsNum = totalVisits > 0 ? totalVisits : 0;
  const startRow = 2;
  for (let j = 0; j < totalVisitsNum; j++) {
    const row = startRow + j;
    const visitNumber = j + 1;
    let visitSheetName = (j === 0) ? "第1次初訪" : (j + 1 === totalVisitsNum) ? `第${j + 1}次結案` : `第${j + 1}次`;
    try {
      totalCaseSheet.getRange(row, 1).setFormula(`=COUNTIF(D${startRow}:D${startRow + totalVisitsNum - 1},"<>")`);
      totalCaseSheet.getRange(row, 2).setValue(caseNumber);
      totalCaseSheet.getRange(row, 3).setValue(caseName);
      totalCaseSheet.getRange(row, 7).setValue(1);
      totalCaseSheet.getRange(row, 8).setValue(totalVisitsNum);
      totalCaseSheet.getRange(row, 9).setFormula(`=H${row}-A${row}`);
      totalCaseSheet.getRange(row, 10).setValue(""); // 備註欄留空，允許手動輸入
      totalCaseSheet.getRange(row, 11).setFormula(`=HYPERLINK("${pdfUrl || ""}", "${ownerName}-${caseNumber}-案件報告.pdf")`);
      Logger.log(`ℍ ${caseSheetName} 第 ${row} 行 K 欄設置 PDF 連結`);
      totalCaseSheet.getRange(row, 12).setFormula(`=HYPERLINK("${visitRecordLink}", "${caseName}-${caseNumber}-${ownerName}")`);
      Logger.log(`ℍ ${caseSheetName} 第 ${row} 行 L 欄設置訪視記錄表連結`);

      // 結案欄 (M 欄) 從 visitRecordId 動態抓取，對應次數
      const closeCaseFormula = `=IFERROR(IF(ISDATE(IMPORTRANGE("${visitRecordId}", "'${visitSheetName}'!N21")), TEXT(IMPORTRANGE("${visitRecordId}", "'${visitSheetName}'!N21"), "yyyy年MM月dd日"), ""), "")`;
      totalCaseSheet.getRange(row, 13).setFormula(closeCaseFormula);
      Logger.log(`ℍ ${caseSheetName} 第 ${row} 行 M 欄設置結案公式，來源: ${visitRecordId}, 工作表: ${visitSheetName}, 單元格: N21`);
    } catch (error) {
      Logger.log(`🚨 填入個案資料失敗 (行 ${row}): ${error.message}`);
      continue;
    }
  }

  const totalOverviewSheet = totalSs.getSheetByName("總表");
  const totalOverviewData = totalOverviewSheet.getDataRange().getValues();
  let caseExistsInOverview = false;
  let overviewRowIndex = -1;

  for (let i = 1; i < totalOverviewData.length; i++) {
    const caseLink = totalOverviewData[i][0];
    if (caseLink && caseLink.includes(`${caseName}-${caseNumber}`)) {
      caseExistsInOverview = true;
      overviewRowIndex = i + 1;
      break;
    }
  }

  const totalSheetUrl = totalSs.getUrl();
  const sanitizedCaseSheetName = caseSheetName.replace(/'/g, "\\'").replace(/"/g, '\\"').replace(/ /g, "\\ ");
  Logger.log(`ℍ 設置總表，工作表名稱: ${sanitizedCaseSheetName}`);

  try {
    if (!caseExistsInOverview) {
      const totalOverviewLastRow = totalOverviewSheet.getLastRow();
      totalOverviewSheet.getRange(totalOverviewLastRow + 1, 1, 1, 6).setValues([[
        `=HYPERLINK("${totalSheetUrl}#gid=${totalCaseSheet.getSheetId()}", "${caseName}-${caseNumber}")`,
        totalVisits,
        `=IFERROR(INDIRECT("'${sanitizedCaseSheetName}'!A2"), 0)`,
        `=B${totalOverviewLastRow + 1}-C${totalOverviewLastRow + 1}`,
        `=IFERROR(INDEX(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1)<>""), ROWS(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1)<>""))), "")`,
        `=IFERROR(INDEX(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1)<>""), ROWS(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1)<>""))), "")`
      ]]);
      Logger.log(`📈 總表新增記錄成功: ${totalSs.getName()}, 案號 ${caseNumber}`);
    } else {
      totalOverviewSheet.getRange(overviewRowIndex, 1, 1, 6).setValues([[
        `=HYPERLINK("${totalSheetUrl}#gid=${totalCaseSheet.getSheetId()}", "${caseName}-${caseNumber}")`,
        totalVisits,
        `=IFERROR(INDIRECT("'${sanitizedCaseSheetName}'!A2"), 0)`,
        `=B${overviewRowIndex}-C${overviewRowIndex}`,
        `=IFERROR(INDEX(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1)<>""), ROWS(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!J2:J" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!J2:J")) + 1)<>""))), "")`,
        `=IFERROR(INDEX(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M 個案工作表") + 1), INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1)<>""), ROWS(FILTER(INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1), INDIRECT("'${sanitizedCaseSheetName}'!M2:M" & ROWS(INDIRECT("'${sanitizedCaseSheetName}'!M2:M")) + 1)<>""))), "")`
      ]]);
      Logger.log(`📈 總表更新記錄成功: ${totalSs.getName()}, 案號 ${caseNumber}, 行 ${overviewRowIndex}`);
    }
    SpreadsheetApp.flush();
    Logger.log(`ℍ 總表數據刷新完成`);
  } catch (error) {
    Logger.log(`🚨 總表更新失敗: ${error.message}`);
  }

  const timeOptions = [];
  for (let h = 8; h <= 21; h++) {
    for (let m = 0; m < 60; m += 15) {
      timeOptions.push(`${h < 10 ? "0" + h : h}:${m === 0 ? "00" : m < 10 ? "0" + m : m}`);
    }
  }
  try {
    totalCaseSheet.getRange(startRow, 4, totalVisitsNum, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setHelpText("請選擇日期").build())
      .setNumberFormat("yyyy年MM月dd日");
    totalCaseSheet.getRange(startRow, 5, totalVisitsNum, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(timeOptions, true).setAllowInvalid(false).build());
    totalCaseSheet.getRange(startRow, 6, totalVisitsNum, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(timeOptions, true).setAllowInvalid(false).build());
    Logger.log(`ℍ 設置日期和時間驗證成功`);
  } catch (error) {
    Logger.log(`🚨 設置驗證失敗: ${error.message}`);
  }

  let summarySheet = spreadsheet.getSheetByName("訪視總表");
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("訪視總表");
    const headersFull = ["次數", "負責人", "個案連結", "服務日期", "訪視次數", "總共訪視次數", "剩餘訪視次數", "完成訪視記錄", "備註", "結案", "業務報酬（單次）", "交通費補助", "共計報酬", "總計報酬"];
    summarySheet.getRange(1, 1, 1, headersFull.length).setValues([headersFull])
      .setFontWeight("bold")
      .setBackground("#d9e8f5")
      .setHorizontalAlignment("center");
    Logger.log(`📊 創建訪視總表成功`);
  }

  const summaryData = summarySheet.getDataRange().getValues();
  const existingEntries = new Set();
  for (let i = 1; i < summaryData.length; i++) {
    const visitNumber = summaryData[i][0];
    const caseKey = summaryData[i][2];
    existingEntries.add(`${caseKey}-${visitNumber}`);
  }

  let lastSummaryRow = summarySheet.getLastRow();
  for (let j = 0; j < totalVisitsNum; j++) {
    const row = startRow + j;
    const summaryRow = lastSummaryRow + j + 1;
    const caseKey = `=HYPERLINK("${totalSs.getUrl()}#gid=${totalCaseSheet.getSheetId()}", "${caseName}-${caseNumber}")`;
    const visitNumber = j + 1;
    const entryKey = `${caseKey}-${visitNumber}`;

    if (existingEntries.has(entryKey)) {
      Logger.log(`ℍ 訪視總表已存在記錄: ${entryKey}，跳過`);
      continue;
    }

    try {
      summarySheet.getRange(summaryRow, 1).setValue(visitNumber);
      summarySheet.getRange(summaryRow, 2).setValue(ownerName);
      summarySheet.getRange(summaryRow, 3).setFormula(caseKey);
      summarySheet.getRange(summaryRow, 4).setFormula(`=IFERROR(IF(ISDATE(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!D${row}")), TEXT(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!D${row}"), "yyyy年MM月dd日"), ""), "")`);
      summarySheet.getRange(summaryRow, 5).setFormula(`=IFERROR(IF(ISDATE(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!D${row}")), IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!G${row}"), ""), "")`);
      summarySheet.getRange(summaryRow, 6).setFormula(`=IFERROR(IF(ISDATE(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!D${row}")), IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!H${row}"), ""), "")`);
      summarySheet.getRange(summaryRow, 7).setFormula(`=IFERROR(IF(ISDATE(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!D${row}")), IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!I${row}"), ""), "")`);

      let visitSheetName = (j === 0) ? "第1次初訪" : (j + 1 === totalVisitsNum) ? `第${j + 1}次結案` : `第${j + 1}次`;
      summarySheet.getRange(summaryRow, 8).setFormula(`=IFERROR(IF(IMPORTRANGE("${visitRecordId}", "'${visitSheetName}'!H21")<>"", TEXT(IMPORTRANGE("${visitRecordId}", "'${visitSheetName}'!H21"), "yyyy年MM月dd日"), ""), "")`);
      summarySheet.getRange(summaryRow, 9).setFormula(`=IFERROR(IF(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!J${row}")<>"", IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!J${row}"), ""), "")`);
      summarySheet.getRange(summaryRow, 10).setFormula(`=IFERROR(IF(IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!M${row}")<>"", IMPORTRANGE("${totalSs.getUrl()}", "'${caseSheetName}'!M${row}"), ""), "")`);
      summarySheet.getRange(summaryRow, 11).setValue(remuneration || 0);
      summarySheet.getRange(summaryRow, 12).setValue(transport || 0);
      summarySheet.getRange(summaryRow, 13).setFormula(`=IFERROR(IF(AND(ISNUMBER(E${summaryRow}), ISNUMBER(K${summaryRow}), ISNUMBER(L${summaryRow})), E${summaryRow}*(K${summaryRow}+L${summaryRow}), 0), 0)`);
      summarySheet.getRange(summaryRow, 14).setFormula(`=IFERROR(IF(ROW()=2, M2, IF(MID(D${summaryRow}, 6, 2)<>MID(D${summaryRow-1}, 6, 2), M${summaryRow}, N${summaryRow-1}+M${summaryRow})), 0)`);

      existingEntries.add(entryKey);
      Logger.log(`ℍ 新增訪視總表記錄: ${entryKey}`);
    } catch (error) {
      Logger.log(`🚨 填入訪視總表記錄失敗 (行 ${summaryRow}): ${error.message}`);
      continue;
    }
  }

  // 提前宣告 totalRows，並使用 let 關鍵字
  let totalRows = summarySheet.getLastRow();

  // 設置「訪視總表」中的數字欄位格式和對齊
  const summaryHeaders = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
  const remunerationCol = summaryHeaders.indexOf("業務報酬（單次）") + 1; // 第 11 欄 (K)
  const transportCol = summaryHeaders.indexOf("交通費補助") + 1; // 第 12 欄 (L)
  const totalRemunerationCol = summaryHeaders.indexOf("共計報酬") + 1; // 第 13 欄 (M)
  const totalCompensationCol = summaryHeaders.indexOf("總計報酬") + 1; // 第 14 欄 (N)

  if (totalRows > 1) {
    if (remunerationCol > 0) {
      summarySheet.getRange(2, remunerationCol, totalRows - 1, 1)
        .setNumberFormat("#,##0") // 設置千位分隔符
        .setHorizontalAlignment("right"); // 靠右對齊
      Logger.log(`ℍ 訪視總表 - 業務報酬（單次）欄位 (K) 設置為數字格式並靠右對齊`);
    }
    if (transportCol > 0) {
      summarySheet.getRange(2, transportCol, totalRows - 1, 1)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");
      Logger.log(`ℍ 訪視總表 - 交通費補助欄位 (L) 設置為數字格式並靠右對齊`);
    }
    if (totalRemunerationCol > 0) {
      summarySheet.getRange(2, totalRemunerationCol, totalRows - 1, 1)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");
      Logger.log(`ℍ 訪視總表 - 共計報酬欄位 (M) 設置為數字格式並靠右對齊`);
    }
    if (totalCompensationCol > 0) {
      summarySheet.getRange(2, totalCompensationCol, totalRows - 1, 1)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");
      Logger.log(`ℍ 訪視總表 - 總計報酬欄位 (N) 設置為數字格式並靠右對齊`);
    }
  }

  // 設置派案總表 Y 欄（結案）從「訪視總表」J 欄（結案，第 10 欄）抓取最新數據
  if (effectiveRowIndex !== null && effectiveRowIndex >= 0) {
    try {
      const caseNumberCol = headers.findIndex(h => h.includes("案號")) + 1;
      const caseNumberValue = mainSheet.getRange(effectiveRowIndex + 1, caseNumberCol).getValue().trim();
      const caseNameCol = headers.findIndex(h => h.includes("個案姓名")) + 1;
      const caseNameValue = mainSheet.getRange(effectiveRowIndex + 1, caseNameCol).getValue().trim();
      const caseKey = `${caseNameValue}-${caseNumberValue}`; // 鍵值為 caseName-caseNumber 格式，例如 "xxx-114-B020-03c"

      Logger.log(`ℍ 設置派案總表 Y 欄，案號: ${caseNumberValue}, 個案姓名: ${caseNameValue}, 鍵值: ${caseKey}`);

      // 從「訪視總表」J 欄（結案，第 10 欄）抓取最新數據
      const closeCaseFormula = `=IFERROR(INDEX(FILTER('訪視總表'!J:J, '訪視總表'!C:C="${caseKey}", '訪視總表'!J:J<>""), ROWS(FILTER('訪視總表'!J:J, '訪視總表'!C:C="${caseKey}", '訪視總表'!J:J<>""))), "")`;
      mainSheet.getRange(effectiveRowIndex + 1, 25).setFormula(closeCaseFormula); // Y 欄 (第 25 欄)
      Logger.log(`ℍ 派案總表 Y 欄設置成功，行 ${effectiveRowIndex + 1}, 公式: ${closeCaseFormula}`);
      SpreadsheetApp.flush();
      const setFormula = mainSheet.getRange(effectiveRowIndex + 1, 25).getFormula();
      Logger.log(`ℍ 檢查 Y 欄公式是否被覆蓋，實際公式: ${setFormula}`);
    } catch (error) {
      Logger.log(`🚨 設置派案總表 Y 欄失敗: ${error.message}`);
      mainSheet.getRange(effectiveRowIndex + 1, 25).setValue(`錯誤：${error.message}`);
    }
  }

  // 設置派案總表 AA 欄（備註）從「訪視總表」I 欄（備註，第 9 欄）抓取最新數據
  if (effectiveRowIndex !== null && effectiveRowIndex >= 0) {
    try {
      const caseNumberCol = headers.findIndex(h => h.includes("案號")) + 1;
      const caseNumberValue = mainSheet.getRange(effectiveRowIndex + 1, caseNumberCol).getValue().trim();
      const caseNameCol = headers.findIndex(h => h.includes("個案姓名")) + 1;
      const caseNameValue = mainSheet.getRange(effectiveRowIndex + 1, caseNameCol).getValue().trim();
      const caseKey = `${caseNameValue}-${caseNumberValue}`; // 鍵值為 caseName-caseNumber 格式，例如 "xxx-114-B020-03c"

      Logger.log(`ℍ 設置派案總表 AA 欄，案號: ${caseNumberValue}, 個案姓名: ${caseNameValue}, 鍵值: ${caseKey}`);

      // 從「訪視總表」I 欄（備註，第 9 欄）抓取最新數據
      const remarkFormula = `=IFERROR(INDEX(FILTER('訪視總表'!I:I, '訪視總表'!C:C="${caseKey}", '訪視總表'!I:I<>""), ROWS(FILTER('訪視總表'!I:I, '訪視總表'!C:C="${caseKey}", '訪視總表'!I:I<>""))), "")`;
      mainSheet.getRange(effectiveRowIndex + 1, 27).setFormula(remarkFormula); // AA 欄 (第 27 欄)
      Logger.log(`ℍ 派案總表 AA 欄設置成功，行 ${effectiveRowIndex + 1}, 公式: ${remarkFormula}`);
      SpreadsheetApp.flush();
      const setFormula = mainSheet.getRange(effectiveRowIndex + 1, 27).getFormula();
      Logger.log(`ℍ 檢查 AA 欄公式是否被覆蓋，實際公式: ${setFormula}`);
    } catch (error) {
      Logger.log(`🚨 設置派案總表 AA 欄失敗: ${error.message}`);
      mainSheet.getRange(effectiveRowIndex + 1, 27).setValue(`錯誤：${error.message}`);
    }
  }

  // 設置派案總表 AB 欄（已訪視次數）和 AC 欄（剩餘訪視次數）從「訪視總表」抓取數據
  if (effectiveRowIndex !== null && effectiveRowIndex >= 0) {
    try {
      const caseNumberCol = headers.findIndex(h => h.includes("案號")) + 1;
      const caseNumberValue = mainSheet.getRange(effectiveRowIndex + 1, caseNumberCol).getValue().trim();
      const caseNameCol = headers.findIndex(h => h.includes("個案姓名")) + 1;
      const caseNameValue = mainSheet.getRange(effectiveRowIndex + 1, caseNameCol).getValue().trim();
      const caseKey = `${caseNameValue}-${caseNumberValue}`; // 鍵值為 caseName-caseNumber 格式，例如 "xxx-114-B020-03c"

      // AB 欄（已訪視次數）：計算「訪視總表」D 欄中有效日期的數量，考慮超連結格式
      Logger.log(`ℍ 設置派案總表 AB 欄（已訪視次數），案號: ${caseNumberValue}, 個案姓名: ${caseNameValue}, 鍵值: ${caseKey}`);
      const alreadyVisitedFormula = `=IFERROR(COUNT(FILTER('訪視總表'!D:D, REGEXMATCH('訪視總表'!C:C, "${caseKey}"), ISDATE('訪視總表'!D:D))), 0)`;
      mainSheet.getRange(effectiveRowIndex + 1, 28).setFormula(alreadyVisitedFormula); // AB 欄 (第 28 欄)
      Logger.log(`ℍ 派案總表 AB 欄設置成功，行 ${effectiveRowIndex + 1}, 公式: ${alreadyVisitedFormula}`);
      SpreadsheetApp.flush();
      const setFormulaAB = mainSheet.getRange(effectiveRowIndex + 1, 28).getFormula();
      Logger.log(`ℍ 檢查 AB 欄公式是否被覆蓋，實際公式: ${setFormulaAB}`);

      // AC 欄（剩餘訪視次數）：從「訪視總表」G 欄（剩餘訪視次數）抓取最新數據
      Logger.log(`ℍ 設置派案總表 AC 欄（剩餘訪視次數），案號: ${caseNumberValue}, 個案姓名: ${caseNameValue}, 鍵值: ${caseKey}`);
      const remainingVisitsFormula = `=IFERROR(INDEX(FILTER('訪視總表'!G:G, '訪視總表'!C:C="${caseKey}", '訪視總表'!G:G<>""), ROWS(FILTER('訪視總表'!G:G, '訪視總表'!C:C="${caseKey}", '訪視總表'!G:G<>""))), "")`;
      mainSheet.getRange(effectiveRowIndex + 1, 29).setFormula(remainingVisitsFormula); // AC 欄 (第 29 欄)
      Logger.log(`ℍ 派案總表 AC 欄設置成功，行 ${effectiveRowIndex + 1}, 公式: ${remainingVisitsFormula}`);
      SpreadsheetApp.flush();
      const setFormulaAC = mainSheet.getRange(effectiveRowIndex + 1, 29).getFormula();
      Logger.log(`ℍ 檢查 AC 欄公式是否被覆蓋，實際公式: ${setFormulaAC}`);
    } catch (error) {
      Logger.log(`🚨 設置派案總表 AB/AC 欄失敗: ${error.message}`);
      mainSheet.getRange(effectiveRowIndex + 1, 28).setValue(`錯誤：${error.message}`); // AB 欄
      mainSheet.getRange(effectiveRowIndex + 1, 29).setValue(`錯誤：${error.message}`); // AC 欄
    }
  }

  // 設置負責人訪視總表的分享權限為編輯
  try {
    DriveApp.getFileById(totalSs.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    Logger.log(`ℍ 已設置負責人訪視時數表分享權限為任何人可編輯`);
  } catch (error) {
    Logger.log(`🚨 設置負責人訪視時數表分享失敗: ${error.message}`);
  }

// 更新派案總表 AB 欄（已訪視次數）
  updateVisitedCount();

  Logger.log(`📊 訪視總表總行數: ${totalRows}`);
  return totalSs.getUrl();
}

function updateVisitedCount() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID); // 替換為您的試算表 ID
  const mainSheet = spreadsheet.getSheetByName("派案總表");
  const summarySheet = spreadsheet.getSheetByName("訪視總表");

  if (!mainSheet || !summarySheet) {
    Logger.log("🚨 找不到 '派案總表' 或 '訪視總表'");
    return;
  }

  // 獲取派案總表數據
  const mainData = mainSheet.getDataRange().getValues();
  const headers = mainData[0];
  const caseNumberCol = headers.indexOf("案號") + 1;
  const caseNameCol = headers.indexOf("個案姓名") + 1;
  const visitedCol = 28; // AB 欄（第 28 欄）

  if (caseNumberCol === 0 || caseNameCol === 0) {
    Logger.log("🚨 '派案總表' 表頭中缺少 '案號' 或 '個案姓名'");
    return;
  }

  // 獲取訪視總表數據
  const summaryData = summarySheet.getDataRange().getValues();
  Logger.log(`ℍ 訪視總表總行數: ${summaryData.length - 1}`);

  // 建立案號對應的服務日期計數
  const visitCountMap = new Map();
  for (let i = 1; i < summaryData.length; i++) {
    const caseLink = summaryData[i][2]; // C 欄（個案連結）
    const serviceDate = summaryData[i][3]; // D 欄（服務日期）
    Logger.log(`ℍ 檢查訪視總表第 ${i + 1} 行 - C 欄: ${caseLink}, D 欄: ${serviceDate}`);

    if (caseLink && serviceDate !== "" && serviceDate !== null && serviceDate !== undefined) { // 只要 D 欄非空就算一次
      // 提取 caseKey，處理超連結或純文字
      let caseKey = "";
      if (typeof caseLink === "string" && caseLink.includes("HYPERLINK")) {
        const match = caseLink.match(/"([^"]+)"\)$/);
        caseKey = match ? match[1] : "";
      } else {
        caseKey = caseLink.toString().trim();
      }

      if (caseKey) {
        visitCountMap.set(caseKey, (visitCountMap.get(caseKey) || 0) + 1);
        Logger.log(`ℍ 有效記錄 - 案號: ${caseKey}, 已訪視次數: ${visitCountMap.get(caseKey)}`);
      } else {
        Logger.log(`🚨 無法提取 caseKey - C 欄: ${caseLink}`);
      }
    } else {
      Logger.log(`🚨 D 欄無資料或 C 欄無效 - C 欄: ${caseLink}, D 欄: ${serviceDate}`);
    }
  }

  // 更新派案總表 AB 欄
  for (let i = 1; i < mainData.length; i++) {
    const caseNumber = mainData[i][caseNumberCol - 1].toString().trim();
    const caseName = mainData[i][caseNameCol - 1].toString().trim();
    const caseKey = `${caseName}-${caseNumber}`;
    const visitedCount = visitCountMap.get(caseKey) || 0;
    mainSheet.getRange(i + 1, visitedCol).setValue(visitedCount);
    Logger.log(`ℍ 更新派案總表第 ${i + 1} 行 - 案號: ${caseKey}, 已訪視次數: ${visitedCount}`);
  }
}

/**************************
 * 訪視記錄表生成函數
 **************************/
function createAndShareCopy(ownerName, caseNumber, caseName, email, parentFolder, totalVisits, caseData, headers) {
  const folder = parentFolder; // 使用個案資料夾
  const targetSheet = SpreadsheetApp.openById(VISIT_RECORD_TEMPLATE_ID);
  const copy = targetSheet.copy(`${caseName}-${caseNumber}-${ownerName}`);
  const file = DriveApp.getFileById(copy.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const copiedSheet = SpreadsheetApp.openById(copy.getId());
  const templateSheet = copiedSheet.getSheets()[0];
  const additionalSheets = copiedSheet.getSheets().slice(1);
  additionalSheets.forEach(sheet => copiedSheet.deleteSheet(sheet));

  headers = headers || [];
  caseData = caseData || [];
  const totalVisitsIndex = headers.findIndex(h => h === "總共要訪視次數");
  const caseNumberIndex = headers.findIndex(h => h === "案號");
  const ageIndex = headers.findIndex(h => h === "年齡");
  const genderIndex = headers.findIndex(h => h === "性別");

  const totalVisitsNum = totalVisitsIndex !== -1 && caseData[totalVisitsIndex] ? Number(caseData[totalVisitsIndex]) : (totalVisits > 0 ? totalVisits : 1);
  const caseNumberValue = caseNumberIndex !== -1 && caseData[caseNumberIndex] ? String(caseData[caseNumberIndex]) : (caseNumber || "");
  const ageValue = ageIndex !== -1 && caseData[ageIndex] ? String(caseData[ageIndex]) : "";
  const genderValue = genderIndex !== -1 && caseData[genderIndex] ? String(caseData[genderIndex]) : "";

  for (let i = 1; i <= totalVisitsNum; i++) {
    let sheetName;
    if (i === 1) {
      sheetName = "第1次初訪";
    } else if (i === totalVisitsNum) {
      sheetName = `第${i}次結案`;
    } else {
      sheetName = `第${i}次`;
    }

    let newSheet;
    if (i === 1) {
      newSheet = templateSheet;
      newSheet.setName(sheetName);
    } else {
      newSheet = templateSheet.copyTo(copiedSheet);
      newSheet.setName(sheetName);
    }

    newSheet.getRange("Z5").setValue(sheetName);

    const headersInSheet = newSheet.getRange("4:4").getValues()[0];
    let caseNumberIdx = headersInSheet.indexOf("案號");
    let ageIdx = headersInSheet.indexOf("年齡");
    let genderIdx = headersInSheet.indexOf("性別");

    const baseRow = 4;
    if (caseNumberIdx !== -1) {
      newSheet.getRange(baseRow, caseNumberIdx + 2).setValue(caseNumberValue);
      newSheet.getRange(baseRow, caseNumberIdx + 3).setValue(caseNumberValue);
      newSheet.getRange(baseRow, caseNumberIdx + 4).setValue(caseNumberValue);
    }
    if (ageIdx !== -1) {
      newSheet.getRange(baseRow, ageIdx + 3).setValue(ageValue);
      newSheet.getRange(baseRow, ageIdx + 4).setValue(ageValue);
    }
    if (genderIdx !== -1) {
      newSheet.getRange(baseRow, genderIdx + 3).setValue(genderValue);
      newSheet.getRange(baseRow, genderIdx + 4).setValue(genderValue);
    }
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    Logger.log(`🚨 設置訪視記錄表分享失敗: ${error.message}`);
  }
  Logger.log(`ℍ 嘗試分享訪視記錄表: ${file.getUrl()}`);
  return file.getUrl();
}

/**************************
 * PDF 生成函數
 **************************/
function generatePDF(headers, caseData, serviceDate, plannerLink, ownerName, caseNumber, caseName, visitRecordLink, visitHoursLink) {
  try {
    const doc = DocumentApp.create(`TempDoc_${caseNumber}`);
    const body = doc.getBody();

    const headerStyle = {
      [DocumentApp.Attribute.FONT_SIZE]: 20,
      [DocumentApp.Attribute.FONT_FAMILY]: "Arial",
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: "#000000",
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
    };
    const tableStyle = {
      [DocumentApp.Attribute.FONT_SIZE]: 14,
      [DocumentApp.Attribute.FONT_FAMILY]: "Arial",
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.PADDING_TOP]: 4,
      [DocumentApp.Attribute.PADDING_BOTTOM]: 4
    };
    const noteStyle = {
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.FONT_FAMILY]: "Arial",
      [DocumentApp.Attribute.FOREGROUND_COLOR]: "#666666",
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT
    };

    const title = body.appendParagraph("📌 案件報告");
    title.setAttributes(headerStyle);
    body.appendParagraph("");

    let formattedServiceDate = "無資料";
    const serviceDateIndex = headers.findIndex(h => h === "已預約初訪日期及時間");
    if (serviceDateIndex !== -1 && caseData[serviceDateIndex]) {
      let serviceDateObj;
      if (typeof caseData[serviceDateIndex] === "string") {
        serviceDateObj = new Date(caseData[serviceDateIndex]);
      } else if (caseData[serviceDateIndex] instanceof Date) {
        serviceDateObj = caseData[serviceDateIndex];
      }
      if (serviceDateObj && !isNaN(serviceDateObj.getTime())) {
        const hours = serviceDateObj.getHours();
        const period = hours < 12 ? "上午" : "下午";
        const adjustedHours = hours % 12 || 12;
        const minutes = String(serviceDateObj.getMinutes()).padStart(2, "0");
        const weekdays = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
        const weekday = weekdays[serviceDateObj.getDay()];
        formattedServiceDate = Utilities.formatDate(serviceDateObj, "Asia/Taipei", "yyyy年MM月dd日") + 
                              ` ${weekday} ${period} ${adjustedHours}:${minutes}`;
      }
    }

    const genderIndex = headers.findIndex(h => h === "性別");
    const ageIndex = headers.findIndex(h => h === "年齡");
    const gender = genderIndex !== -1 ? String(caseData[genderIndex] || "未提供") : "未提供";
    const age = ageIndex !== -1 ? String(caseData[ageIndex] || "未提供") : "未提供";
    const genderIcon = gender === "男" ? "👨" : gender === "女" ? "👩" : "❓";

    const tableData = [
      ["項目", "內容"],
      ["👤 負責人", ownerName || ""],
      ["📧 Email", String(caseData[headers.findIndex(h => h === "Email")] || "")],
      ["🔢 總共要訪視次數", String(caseData[headers.findIndex(h => h === "總共要訪視次數")] || "")],
      ["📅 已預約初訪日期及時間", formattedServiceDate],
      ["📝 狀態", String(caseData[headers.findIndex(h => h === "狀態")] || "")],
      ["🆔 案號", caseNumber || ""],
      ["📋 個案類型", String(caseData[headers.findIndex(h => h === "個案類型")] || "")],
      ["👤 個案姓名", caseName || ""],
      [`${genderIcon} 性別`, gender],
      ["🎂 年齡", age],
      ["📞 個案電話", String(caseData[headers.findIndex(h => h === "個案電話")] || "")],
      ["🏠 個案住址", String(caseData[headers.findIndex(h => h === "個案住址")] || "")],
      ["📂 規畫師雲端", ownerName || ""],
      ["📋 訪視記錄表", `${caseName}-${caseNumber}-${ownerName}`],
      ["⏰ 訪視時數表", `${ownerName} 訪視時數表`],
    ];

    const table = body.appendTable(tableData);
    table.setBorderWidth(1).setBorderColor("#000000");
    const totalWidth = 550, itemWidth = 200, contentWidth = totalWidth - itemWidth;

    tableData.forEach((rowData, i) => {
      const row = table.getRow(i);
      row.getCell(0).setWidth(itemWidth).setBackgroundColor(i === 0 ? "#d9e8f5" : "#f9f9f9").editAsText().setText(rowData[0]).setAttributes(tableStyle);
      row.getCell(1).setWidth(contentWidth).setBackgroundColor(i === 0 ? "#d9e8f5" : "#ffffff").editAsText().setText(rowData[1]).setAttributes(tableStyle);
      if (i === 0) {
        row.getCell(0).setBold(true);
        row.getCell(1).setBold(true);
      }
      if (i === 13) row.getCell(1).editAsText().setLinkUrl(plannerLink || "");
      if (i === 14) row.getCell(1).editAsText().setLinkUrl(visitRecordLink || "");
      if (i === 15) row.getCell(1).editAsText().setLinkUrl(visitHoursLink || "");
    });

    const note = body.appendParagraph("備註：請妥善保管上述個案資料，確保符合個人資料保護法規，賽親感謝您的配合！");
    note.setAttributes(noteStyle);

    body.setAttributes({ [DocumentApp.Attribute.MARGIN_LEFT]: 22.5, [DocumentApp.Attribute.MARGIN_RIGHT]: 22.5 });
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(doc.getId()).getAs("application/pdf").setName(`${ownerName}-${caseNumber}-案件報告.pdf`);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    Logger.log(`ℍ PDF 生成成功: ${ownerName}-${caseNumber}-案件報告.pdf`);
    return pdfBlob;
  } catch (error) {
    Logger.log(`🚨 generatePDF 錯誤：${error.message}`);
    throw error;
  }
}

/**************************
 * 更新訪視總表（可選，若不使用公式）
 **************************/
function updateVisitSummaryFromAllHours() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  let summarySheet = spreadsheet.getSheetByName("訪視總表");

  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("訪視總表");
    const headers = ["負責人", "個案姓名+案號", "服務日期", "訪視次數", "總共訪視次數", "剩餘訪視次數"];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#d9e8f5")
      .setBorder(true, true, true, true, true, true);
    Logger.log("ℍ 已創建訪視總表並初始化表頭");
  }

  const mainSheet = spreadsheet.getSheetByName("派案總表");
  const mainData = mainSheet.getDataRange().getValues();
  const headers = mainData[0];
  const indices = {
    owner: headers.findIndex(h => h.includes("負責人")),
    caseNumber: headers.findIndex(h => h.includes("案號")),
    caseName: headers.findIndex(h => h.includes("個案姓名")),
    visitHours: headers.findIndex(h => h.includes("訪視時數表")),
  };

  const summaryData = summarySheet.getDataRange().getValues();
  const existingEntries = new Set();
  for (let i = 1; i < summaryData.length; i++) {
    existingEntries.add(`${summaryData[i][1]}-${summaryData[i][2]}`);
  }

  for (let i = 1; i < mainData.length; i++) {
    const row = mainData[i];
    const ownerName = row[indices.owner];
    const caseNumber = row[indices.caseNumber];
    const caseName = row[indices.caseName];
    let visitHoursLink = row[indices.visitHours];

    if (!visitHoursLink || !caseNumber || !caseName) continue;

    const hyperlinkMatch = visitHoursLink.match(/HYPERLINK\("(.*?)"/);
    if (hyperlinkMatch) visitHoursLink = hyperlinkMatch[1];

    try {
      const visitHoursSpreadsheet = SpreadsheetApp.openByUrl(visitHoursLink);
      const caseSheet = visitHoursSpreadsheet.getSheetByName(`${caseName}-${caseNumber}`);
      if (!caseSheet) continue;

      const caseData = caseSheet.getDataRange().getValues();
      const caseHeaders = caseData[0];
      const serviceDateIdx = caseHeaders.indexOf("服務日期");
      const totalVisitsIdx = caseHeaders.indexOf("總共訪視次數");
      if (serviceDateIdx === -1 || totalVisitsIdx === -1) continue;

      const totalVisits = caseData[1][totalVisitsIdx] || 0;
      const sheetUrl = visitHoursSpreadsheet.getUrl() + "#gid=" + caseSheet.getSheetId();
      const caseKey = `=HYPERLINK("${sheetUrl}", "${caseName}-${caseNumber}")`;

      for (let j = 1; j < caseData.length; j++) {
        const serviceDate = caseData[j][serviceDateIdx];
        if (serviceDate && typeof serviceDate === "object" && !isNaN(serviceDate.getTime())) {
          const formattedDate = Utilities.formatDate(serviceDate, "zh_TW", DATE_FORMAT);
          const visitCount = j;
          const remainingVisits = totalVisits - visitCount;
          const entryKey = `${caseKey}-${formattedDate}`;

          if (!existingEntries.has(entryKey)) {
            const lastRow = summarySheet.getLastRow();
            summarySheet.getRange(lastRow + 1, 1, 1, 6).setValues([
              [ownerName, caseKey, formattedDate, visitCount, totalVisits, remainingVisits]
            ]);
            existingEntries.add(entryKey);
            Logger.log(`📈 新增訪視總表記錄: ${caseName}-${caseNumber}, 服務日期: ${formattedDate}`);
          }
        }
      }
    } catch (error) {
      Logger.log(`🚨 更新訪視總表失敗 (${caseName}-${caseNumber}): ${error.message}`);
    }
  }
}

/**************************
 * 按月份生成報酬表
 **************************/
function generateMonthlyRemunerationSheets() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const summarySheet = spreadsheet.getSheetByName("訪視總表");
  if (!summarySheet) {
    Logger.log("⚠ 找不到訪視總表，無法生成報酬表");
    return;
  }

  const dataRange = summarySheet.getDataRange();
  const data = dataRange.getValues();
  const formulas = dataRange.getFormulas();
  if (data.length <= 1) {
    Logger.log("⚠ 訪視總表沒有資料，無法生成報酬表");
    return;
  }

  Logger.log(`📊 訪視總表總行數: ${data.length - 1}`);

  const monthlyData = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let serviceDate = row[3]; // 服務日期位於第 4 欄 (索引 3)
    const owner = row[1];     // 負責人位於第 2 欄 (索引 1)
    const caseFormula = formulas[i][2] || ""; // 個案連結公式位於第 3 欄 (索引 2)

    if (!serviceDate || serviceDate === "") {
      Logger.log(`Row ${i + 1} - 服務日期為空，跳過`);
      continue;
    }

    // 處理完整的日期格式
    let parsedDate;
    if (typeof serviceDate === "string") {
      try {
        // 移除星期和時間部分，只保留日期
        const dateStr = serviceDate.replace(/ EEEE .*$/, "").replace("上午", "").replace("下午", "").trim();
        const cleanedDateStr = dateStr.replace("年", "-").replace("月", "-").replace("日", "");
        parsedDate = new Date(cleanedDateStr);
        if (isNaN(parsedDate.getTime())) {
          Logger.log(`Row ${i + 1} - 日期解析失敗: ${serviceDate}, 清理後: ${cleanedDateStr}`);
          continue;
        }
      } catch (error) {
        Logger.log(`Row ${i + 1} - 日期解析錯誤: ${serviceDate}, 錯誤: ${error.message}`);
        continue;
      }
    } else if (serviceDate instanceof Date) {
      parsedDate = serviceDate;
    } else {
      Logger.log(`Row ${i + 1} - 服務日期類型無效: ${typeof serviceDate}, 值: ${serviceDate}`);
      continue;
    }

    if (parsedDate && !isNaN(parsedDate.getTime())) {
      const year = Utilities.formatDate(parsedDate, "Asia/Taipei", "yyyy"); // 加入年份
      const month = Utilities.formatDate(parsedDate, "Asia/Taipei", "MM");  // 月份
      const monthKey = `${year}-${month}`; // 例如 "2025-01"
      const monthName = `${month}月份`;    // 例如 "1月份"
      monthlyData[monthKey] = monthlyData[monthKey] || {};
      monthlyData[monthKey][owner] = monthlyData[monthKey][owner] || [];
      monthlyData[monthKey][owner].push({ row: [...row], formula: caseFormula });
      Logger.log(`Row ${i + 1} - 分類到 ${monthName} (${monthKey}), 負責人: ${owner}, 日期: ${Utilities.formatDate(parsedDate, "Asia/Taipei", DATE_FORMAT)}`);
    } else {
      Logger.log(`Row ${i + 1} - 服務日期無效: ${row[3]}，跳過`);
    }
  }

  if (Object.keys(monthlyData).length === 0) {
    Logger.log("⚠ 沒有有效的月份資料可生成報酬表");
    return;
  }

  for (const monthKey in monthlyData) {
    const [year, month] = monthKey.split("-");
    const monthName = `${month}月份報酬表`; // 例如 "01月份報酬表"
    let monthSheet = spreadsheet.getSheetByName(monthName);
    
    if (!monthSheet) {
      monthSheet = spreadsheet.insertSheet(monthName);
      const headers = summarySheet.getRange(1, 1, 1, 14).getValues()[0]; // 調整為 14 欄
      monthSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#d9e8f5")
        .setHorizontalAlignment("center");
      Logger.log(`📊 創建新的報酬表: ${monthName}`);
    } else {
      monthSheet.clear();
      const headers = summarySheet.getRange(1, 1, 1, 14).getValues()[0]; // 調整為 14 欄
      monthSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#d9e8f5")
        .setHorizontalAlignment("center");
      Logger.log(`📊 清空並重置報酬表: ${monthName}`);
    }

    let currentRow = 2;
    const owners = Object.keys(monthlyData[monthKey]).sort();
    for (const owner of owners) {
      const ownerData = monthlyData[monthKey][owner];
      const values = ownerData.map(item => {
        const row = [...item.row];
        // 將金額欄位轉為數字
        row[10] = Number(row[10]) || 0; // 業務報酬（單次）
        row[11] = Number(row[11]) || 0; // 交通費補助
        row[12] = Number(row[12]) || 0; // 共計報酬
        row[13] = Number(row[13]) || 0; // 總計報酬
        return row;
      });
      monthSheet.getRange(currentRow, 1, ownerData.length, 14).setValues(values); // 調整為 14 欄
      for (let i = 0; i < ownerData.length; i++) {
        const formula = ownerData[i].formula;
        if (formula) {
          monthSheet.getRange(currentRow + i, 3).setFormula(formula); // 恢復個案連結公式
        }
      }
      currentRow += ownerData.length;
      Logger.log(`📊 為 ${monthName} 添加 ${owner} 的 ${ownerData.length} 筆資料`);
    }

    // 設置報酬表中的數字欄位格式和對齊
    const monthHeaders = monthSheet.getRange(1, 1, 1, monthSheet.getLastColumn()).getValues()[0];
    const remunerationCol = monthHeaders.indexOf("業務報酬（單次）") + 1; // 第 11 欄 (K)
    const transportCol = monthHeaders.indexOf("交通費補助") + 1; // 第 12 欄 (L)
    const totalRemunerationCol = monthHeaders.indexOf("共計報酬") + 1; // 第 13 欄 (M)
    const totalCompensationCol = monthHeaders.indexOf("總計報酬") + 1; // 第 14 欄 (N)

    const totalRows = monthSheet.getLastRow();
    if (totalRows > 1) {
      if (remunerationCol > 0) {
        monthSheet.getRange(2, remunerationCol, totalRows - 1, 1)
          .setNumberFormat("#,##0")
          .setHorizontalAlignment("right");
        Logger.log(`ℍ ${monthName} - 業務報酬（單次）欄位 (K) 設置為數字格式並靠右對齊`);
      }
      if (transportCol > 0) {
        monthSheet.getRange(2, transportCol, totalRows - 1, 1)
          .setNumberFormat("#,##0")
          .setHorizontalAlignment("right");
        Logger.log(`ℍ ${monthName} - 交通費補助欄位 (L) 設置為數字格式並靠右對齊`);
      }
      if (totalRemunerationCol > 0) {
        monthSheet.getRange(2, totalRemunerationCol, totalRows - 1, 1)
          .setNumberFormat("#,##0")
          .setHorizontalAlignment("right");
        Logger.log(`ℍ ${monthName} - 共計報酬欄位 (M) 設置為數字格式並靠右對齊`);
      }
      if (totalCompensationCol > 0) {
        monthSheet.getRange(2, totalCompensationCol, totalRows - 1, 1)
          .setNumberFormat("#,##0")
          .setHorizontalAlignment("right");
        Logger.log(`ℍ ${monthName} - 總計報酬欄位 (N) 設置為數字格式並靠右對齊`);
      }
    }
  }

  // 調整工作表順序
  const fixedSheets = ["負責人基本資料", "報告總表", "派案總表", "訪視總表"];
  const allSheets = spreadsheet.getSheets();
  const monthSheets = [];

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (!fixedSheets.includes(sheetName) && sheetName.match(/^\d+月份報酬表$/)) {
      const monthNum = parseInt(sheetName.replace("月份報酬表", ""));
      monthSheets.push({ sheet, monthNum });
    }
  });

  monthSheets.sort((a, b) => b.monthNum - a.monthNum);

  let targetIndex = 0;

  // 按照指定的固定順序排列
  fixedSheets.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.setActiveSheet(sheet);
      spreadsheet.moveActiveSheet(targetIndex + 1);
      targetIndex++;
      Logger.log(`ℍ 移動工作表 ${sheetName} 到位置 ${targetIndex}`);
    } else {
      Logger.log(`⚠ 找不到工作表 ${sheetName}，跳過移動`);
    }
  });

  // 排列報酬表（按月份倒序）
  monthSheets.forEach(({ sheet }) => {
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(targetIndex + 1);
    targetIndex++;
    Logger.log(`ℍ 移動報酬表 ${sheet.getName()} 到位置 ${targetIndex}`);
  });

  Logger.log("ℍ 已調整工作表順序：負責人基本資料 -> 報告總表 -> 派案總表 -> 訪視總表 -> 報酬表（按月份倒序）");
}

/**************************
 * 輔助函數：查找現有訪視時數表
 **************************/
function findExistingVisitHoursSheet(ownerName, email, parentFolder) {
  const folder = parentFolder; // 此處應為 ownerFolder
  if (!folder) return null;

  const files = folder.getFilesByName(`${ownerName} 訪視時數表`);
  if (files.hasNext()) {
    const file = files.next();
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (error) {
      Logger.log(`🚨 設置現有訪視時數表分享失敗: ${error.message}`);
    }
    Logger.log(`ℍ 嘗試分享現有訪視時數表: ${file.getUrl()}`);
    return file.getUrl();
  }
  return null;
}

/**************************
 * 輔助函數：獲取 Google Drive 資料夾
 **************************/
function getDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

/**************************
 * 輔助函數：獲取資料夾 URL 並設置分享權限
 **************************/
function getDriveFolderUrl(folderName, email) {
  const folder = getDriveFolder(folderName);
  if (!folder) {
    Logger.log(`⚠ 未找到規畫師雲端資料夾: ${folderName}`);
    return "⚠ 找不到對應的資料夾";
  }
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log(`ℍ 嘗試分享資料夾: ${folder.getUrl()}`);
    return folder.getUrl();
  } catch (error) {
    Logger.log(`🚨 分享資料夾失敗 (${folderName}): ${error.message}`);
    return "⚠ 分享資料夾時發生錯誤";
  }
}

/**************************
 * 輔助函數：將 PDF 儲存到 Google Drive
 **************************/
function savePdfToDrive(pdfFile, ownerName, caseNumber, caseName, email, parentFolder) {
  const folder = parentFolder; // 使用個案資料夾
  const savedPdf = folder.createFile(pdfFile.setName(`${ownerName}-${caseNumber}-案件報告.pdf`));
  try {
    savedPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    Logger.log(`🚨 設置 PDF 分享失敗: ${error.message}`);
  }
  Logger.log(`ℍ 嘗試分享 PDF: ${savedPdf.getUrl()}`);
  return savedPdf.getUrl();
}

/**************************
 * 輔助函數：從連結中提取文件 ID
 **************************/
function getFileIdFromLink(link) {
  if (!link) return null;
  const matches = link.match(/[-\w]{25,}/);
  return matches ? matches[0] : null;
}

/**************************
 * 設置每小時觸發器
 **************************/
function setupHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === "generateMonthlyRemunerationSheets") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger("generateMonthlyRemunerationSheets")
    .timeBased()
    .everyHours(1)
    .inTimezone("Asia/Taipei")
    .create();
  Logger.log("ℍ 已設置每小時自動掃描訪視總表並更新報酬表");
}

/**************************
 * createSummaryStatsSheet：訪視總表初始化
 **************************/
function createSummaryStatsSheet() {
  const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
  let summarySheet = spreadsheet.getSheetByName("訪視總表");
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("訪視總表");
    Logger.log("ℍ 已創建「訪視總表」");
  }
}

/**************************
 * 報告總表生成與更新函數（支持指定試算表）
 **************************/
function initializeReportSummarySheet(spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID)) {
  let reportSheet = spreadsheet.getSheetByName("報告總表");
  
  const headers = [
    "篩選條件", "負責人", "年份", "月份", "個案類型", "交通補助", "次數", "個案連結", 
    "服務日期", "訪視次數", "總共訪視次數", "剩餘訪視次數", "完成訪視記錄", "備註", 
    "結案", "業務報酬（單次）", "交通費補助", "共計報酬", "總計報酬"
  ];

  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet("報告總表");
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#d9e8f5")
      .setHorizontalAlignment("center");
    
    setupFilterDropdowns(reportSheet);
    Logger.log("📊 創建報告總表成功，初始狀態為空白");
  } else {
    reportSheet.clear();
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#d9e8f5")
      .setHorizontalAlignment("center");
    
    setupFilterDropdowns(reportSheet);
    Logger.log("📊 重置報告總表成功，保持空白");
  }

  // 調整工作表順序
  const fixedSheets = ["負責人基本資料", "報告總表", "派案總表", "訪視總表"];
  let targetIndex = 0;

  fixedSheets.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.setActiveSheet(sheet);
      spreadsheet.moveActiveSheet(targetIndex + 1);
      targetIndex++;
      Logger.log(`ℍ 移動工作表 ${sheetName} 到位置 ${targetIndex}`);
    } else {
      Logger.log(`⚠ 找不到工作表 ${sheetName}，跳過移動`);
    }
  });

  const allSheets = spreadsheet.getSheets();
  const monthSheets = [];
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (!fixedSheets.includes(sheetName) && sheetName.match(/^\d+月份報酬表$/)) {
      const monthNum = parseInt(sheetName.replace("月份報酬表", ""));
      monthSheets.push({ sheet, monthNum });
    }
  });

  monthSheets.sort((a, b) => b.monthNum - a.monthNum);
  monthSheets.forEach(({ sheet }) => {
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(targetIndex + 1);
    targetIndex++;
    Logger.log(`ℍ 移動報酬表 ${sheet.getName()} 到位置 ${targetIndex}`);
  });

  Logger.log("ℍ 工作表順序調整完成：負責人基本資料 -> 報告總表 -> 派案總表 -> 訪視總表 -> 報酬表（按月份倒序）");
}

/**************************
 * 設置篩選條件下拉選單
 **************************/
function setupFilterDropdowns(reportSheet) {
  try {
    const visitSummarySheet = SpreadsheetApp.openById(TARGET_SHEET_ID).getSheetByName("訪視總表");
    if (!visitSummarySheet) {
      Logger.log("⚠ 找不到訪視總表，無法設置下拉選單");
      return;
    }
    const visitData = visitSummarySheet.getDataRange().getValues();
    const visitHeaders = visitData[0];
    Logger.log(`📋 訪視總表欄位: ${visitHeaders.join(", ")}`);

    // 負責人選項
    const ownerIndex = visitHeaders.indexOf("負責人");
    if (ownerIndex === -1) {
      Logger.log("⚠ 訪視總表中找不到 '負責人' 欄位");
      return;
    }
    const owners = [""].concat([...new Set(visitData.slice(1).map(row => row[ownerIndex]).filter(Boolean))]);
    Logger.log(`📋 負責人選項: ${owners.join(", ")}`);

    // 年份選項
    const dateIndex = visitHeaders.indexOf("服務日期");
    if (dateIndex === -1) {
      Logger.log("⚠ 訪視總表中找不到 '服務日期' 欄位");
      return;
    }
    const years = [""].concat([...new Set(visitData.slice(1).map(row => {
      const date = row[dateIndex];
      if (date && typeof date === "string") {
        const yearMatch = date.match(/(\d{4})|(\d{2})$/);
        return yearMatch ? (yearMatch[1] || (yearMatch[2] && `20${yearMatch[2]}`) || "") : "";
      } else if (date instanceof Date) {
        return date.getFullYear().toString();
      }
      return "";
    }).filter(Boolean))].sort());
    Logger.log(`📋 年份選項: ${years.join(", ")}`);

    // 月份選項
    const months = [""].concat([...new Set(visitData.slice(1).map(row => {
      const date = row[dateIndex];
      if (date && typeof date === "string") {
        const monthMatch = date.match(/年(\d{1,2})月/) || 
                         date.match(/(\d{1,2})月/) || 
                         date.match(/(\d{1,2})[-/]/) || 
                         date.match(/(\d{2})(\d{2})/) || 
                         date.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/) || 
                         date.match(/(\d{4})(\d{2})/) || 
                         date.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
        return monthMatch ? (monthMatch[1] || monthMatch[2]).padStart(2, "0") : "";
      } else if (date instanceof Date) {
        return (date.getMonth() + 1).toString().padStart(2, "0");
      }
      return "";
    }).filter(Boolean))].sort());
    Logger.log(`📋 月份選項: ${months.join(", ")}`);

    // 個案類型選項
    const caseLinkIndex = visitHeaders.indexOf("個案連結");
    if (caseLinkIndex === -1) {
      Logger.log("⚠ 訪視總表中找不到 '個案連結' 欄位");
      return;
    }
    const types = [""].concat([...new Set(visitData.slice(1).map(row => {
      const caseLink = row[caseLinkIndex] || "";
      const caseNumberMatch = caseLink.match(/B020-\d{2,3}[A-Za-z]{1,2}/) || caseLink.match(/-(\d{2,3}[A-Za-z]{1,2})$/) || caseLink.match(/\d{3}[A-Za-z]{1,2}/);
      const caseNumber = caseNumberMatch ? (caseNumberMatch[1] || caseNumberMatch[0]) : "";
      let caseType = "未分類";
      if (caseNumber) {
        const lastTwoChars = caseNumber.slice(-2).toLowerCase();
        const lastChar = caseNumber.slice(-1).toLowerCase();
        if (lastTwoChars === "if") {
          caseType = "if";
        } else if (["p", "c", "i"].includes(lastChar)) {
          caseType = lastChar;
        }
      }
      return caseType !== "未分類" ? caseType : "";
    }).filter(Boolean))].sort());
    Logger.log(`📋 個案類型選項: ${types.join(", ")}`);

    // 交通補助選項（提取所有非空值，包括 0）
    const transportAllowanceIndex = visitHeaders.indexOf("交通費補助");
    if (transportAllowanceIndex === -1) {
      Logger.log("⚠ 訪視總表中找不到 '交通費補助' 欄位");
      return;
    }
    const transportAllowanceValues = visitData.slice(1).map(row => {
      const value = row[transportAllowanceIndex];
      return (value !== null && value !== undefined && String(value).trim() !== "") ? String(value).trim() : null;
    }).filter(Boolean);
    const uniqueTransportAllowances = ["", "有", "無"].concat([...new Set(transportAllowanceValues)].sort((a, b) => parseFloat(a) - parseFloat(b)));
    Logger.log(`📋 交通補助選項: ${uniqueTransportAllowances.join(", ")}`);

    // 設置下拉選單（包含空選項）
    reportSheet.getRange(2, 2).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(owners)
      .setAllowInvalid(false)
      .build());
    reportSheet.getRange(2, 3).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(years)
      .setAllowInvalid(false)
      .build());
    reportSheet.getRange(2, 4).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(months)
      .setAllowInvalid(false)
      .build());
    reportSheet.getRange(2, 5).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(types)
      .setAllowInvalid(false)
      .build());
    reportSheet.getRange(2, 6).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueTransportAllowances)
      .setAllowInvalid(false)
      .build());

    Logger.log("📋 設置篩選條件下拉選單成功");
  } catch (error) {
    Logger.log(`⚠ setupFilterDropdowns 執行錯誤: ${error.message}`);
  }
}

/**************************
 * 根據篩選條件更新報告總表（修訂版）
 **************************/
function updateReportSummarySheet(e) {
  Logger.log("📋 開始執行 updateReportSummarySheet");

  try {
    const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const reportSheet = spreadsheet.getSheetByName("報告總表");
    if (!reportSheet) {
      Logger.log("⚠ 找不到報告總表，無法更新");
      return;
    }

    // 檢查報告總表的行數和列數
    const lastRow = reportSheet.getLastRow();
    const lastColumn = reportSheet.getLastColumn();
    Logger.log(`📋 報告總表 - 最後一行: ${lastRow}, 最後一列: ${lastColumn}`);

    // 如果是 onEdit 觸發，檢查編輯位置
    if (e && e.range) {
      const range = e.range;
      const row = range.getRow();
      const col = range.getColumn();
      if (row !== 2 || col < 2 || col > 6) {
        Logger.log(`⚠ 編輯不在篩選條件行 (行 ${row}, 列 ${col})，跳過更新`);
        return;
      }
    }

    // 獲取篩選條件
    const ownerFilter = reportSheet.getRange(2, 2).getValue() || "";
    const yearFilter = reportSheet.getRange(2, 3).getValue() || "";
    const monthFilter = reportSheet.getRange(2, 4).getValue() || "";
    const typeFilter = reportSheet.getRange(2, 5).getValue() || "";
    const transportAllowanceFilter = reportSheet.getRange(2, 6).getValue() || "";
    Logger.log(`📋 篩選條件 - 負責人: ${ownerFilter}, 年份: ${yearFilter}, 月份: ${monthFilter}, 個案類型: ${typeFilter}, 交通補助: ${transportAllowanceFilter}`);

    // 檢查是否至少有一個篩選條件被選擇
    const hasFilter = ownerFilter !== "" || yearFilter !== "" || monthFilter !== "" || typeFilter !== "" || transportAllowanceFilter !== "";
    Logger.log(`📋 是否有篩選條件: ${hasFilter}`);

    // 清空舊資料（從第 3 行開始）
    const rowsToClear = lastRow - 2;
    if (rowsToClear > 0) {
      reportSheet.getRange(3, 1, rowsToClear, lastColumn).clear();
      Logger.log(`📋 已清除報告總表第 3 行開始的 ${rowsToClear} 行數據`);
    } else {
      Logger.log("📋 報告總表無需清除數據（行數少於 3）");
    }

    // 如果沒有篩選條件，直接結束
    if (!hasFilter) {
      Logger.log("📋 所有篩選條件均為空，已清空報告總表資料，不進行資料填充");
      return;
    }

    const visitSummarySheet = spreadsheet.getSheetByName("訪視總表");
    if (!visitSummarySheet) {
      Logger.log("⚠ 找不到訪視總表，無法更新");
      reportSheet.getRange(3, 1).setValue("找不到訪視總表");
      return;
    }
    const visitData = visitSummarySheet.getDataRange().getValues();
    if (visitData.length <= 1) {
      Logger.log("⚠ 訪視總表無資料");
      reportSheet.getRange(3, 1).setValue("訪視總表無資料");
      return;
    }
    const visitHeaders = visitData[0];
    Logger.log(`📋 訪視總表欄位: ${visitHeaders.join(", ")}`);

    // 動態查找欄位索引
    const ownerIndex = visitHeaders.indexOf("負責人");
    const caseLinkIndex = visitHeaders.indexOf("個案連結");
    const dateIndex = visitHeaders.indexOf("服務日期");
    const frequencyIndex = visitHeaders.indexOf("次數");
    const visitCountIndex = visitHeaders.indexOf("訪視次數");
    const totalVisitCountIndex = visitHeaders.indexOf("總共訪視次數");
    const remainingVisitCountIndex = visitHeaders.indexOf("剩餘訪視次數");
    const recordCompleteIndex = visitHeaders.indexOf("完成訪視記錄");
    const noteIndex = visitHeaders.indexOf("備註");
    const caseClosedIndex = visitHeaders.indexOf("結案") !== -1 ? visitHeaders.indexOf("結案") : visitHeaders.indexOf("是否結案");
    const remunerationIndex = visitHeaders.indexOf("業務報酬（單次）") !== -1 ? visitHeaders.indexOf("業務報酬（單次）") : visitHeaders.indexOf("單次報酬");
    const transportAllowanceIndex = visitHeaders.indexOf("交通費補助");
    const totalRemunerationIndex = visitHeaders.indexOf("共計報酬");
    const totalCompensationIndex = visitHeaders.indexOf("總計報酬");

    // 檢查必要欄位
    if (ownerIndex === -1 || caseLinkIndex === -1 || dateIndex === -1 || transportAllowanceIndex === -1) {
      Logger.log("⚠ 訪視總表必要欄位缺失");
      reportSheet.getRange(3, 1).setValue("訪視總表缺少必要欄位");
      return;
    }
    Logger.log(`📋 欄位索引 - 負責人: ${ownerIndex}, 個案連結: ${caseLinkIndex}, 服務日期: ${dateIndex}, 次數: ${frequencyIndex}, 訪視次數: ${visitCountIndex}, 總共訪視次數: ${totalVisitCountIndex}, 剩餘訪視次數: ${remainingVisitCountIndex}, 完成訪視記錄: ${recordCompleteIndex}, 備註: ${noteIndex}, 結案: ${caseClosedIndex}, 業務報酬（單次）: ${remunerationIndex}, 交通費補助: ${transportAllowanceIndex}, 共計報酬: ${totalRemunerationIndex}, 總計報酬: ${totalCompensationIndex}`);

    let filteredData = [];
    visitData.slice(1).forEach((row, index) => {
      const owner = row[ownerIndex] || "";
      const caseLink = row[caseLinkIndex] || "";
      const serviceDate = row[dateIndex] || "";
      const transportAllowance = String(row[transportAllowanceIndex] || "");

      // 提取案號和個案類型
      const caseNumberMatch = caseLink.match(/B020-\d{2,3}[A-Za-z]{1,2}/) || caseLink.match(/-(\d{2,3}[A-Za-z]{1,2})$/) || caseLink.match(/\d{3}[A-Za-z]{1,2}/);
      const caseNumber = caseNumberMatch ? (caseNumberMatch[1] || caseNumberMatch[0]) : "";
      let caseType = "未分類";
      if (caseNumber) {
        const lastTwoChars = caseNumber.slice(-2).toLowerCase();
        const lastChar = caseNumber.slice(-1).toLowerCase();
        if (lastTwoChars === "if") caseType = "if";
        else if (["p", "c", "i"].includes(lastChar)) caseType = lastChar;
      }

      // 提取年份和月份（更穩健的解析）
      let year = "", month = "";
      if (serviceDate instanceof Date) {
        year = serviceDate.getFullYear().toString();
        month = (serviceDate.getMonth() + 1).toString().padStart(2, "0");
      } else if (typeof serviceDate === "string") {
        const parts = serviceDate.match(/(\d{4})年(\d{1,2})月/) || serviceDate.match(/(\d{4})-(\d{1,2})/);
        if (parts) {
          year = parts[1];
          month = parts[2].padStart(2, "0");
        }
      }
      Logger.log(`📋 行 ${index + 2}: 服務日期=${serviceDate}, 年份=${year}, 月份=${month}, 個案類型=${caseType}`);

      // 交通補助判斷
      const transportValue = Number(transportAllowance) || 0;
      const hasTransportAllowance = transportValue > 0;

      // 篩選條件匹配
      const matchesOwner = !ownerFilter || owner === ownerFilter;
      const matchesYear = !yearFilter || year === yearFilter.toString();
      const matchesMonth = !monthFilter || month === monthFilter.toString().padStart(2, "0");
      const matchesType = !typeFilter || caseType.toLowerCase() === typeFilter.toLowerCase();
      const matchesTransport = !transportAllowanceFilter || 
        (transportAllowanceFilter === "有" && hasTransportAllowance) || 
        (transportAllowanceFilter === "無" && !hasTransportAllowance) || 
        transportValue.toString() === transportAllowanceFilter;

      if (matchesOwner && matchesYear && matchesMonth && matchesType && matchesTransport) {
        const remuneration = remunerationIndex !== -1 ? Number(row[remunerationIndex]) || 0 : 0;
        const totalRemuneration = totalRemunerationIndex !== -1 ? Number(row[totalRemunerationIndex]) || 0 : 0;
        const totalCompensation = totalCompensationIndex !== -1 ? Number(row[totalCompensationIndex]) || 0 : 0;

        filteredData.push([
          "數據", owner, year, month, caseType, transportValue,
          frequencyIndex !== -1 ? row[frequencyIndex] || "" : "",
          caseLink, serviceDate,
          visitCountIndex !== -1 ? row[visitCountIndex] || "" : "",
          totalVisitCountIndex !== -1 ? row[totalVisitCountIndex] || "" : "",
          remainingVisitCountIndex !== -1 ? row[remainingVisitCountIndex] || "" : "",
          recordCompleteIndex !== -1 ? row[recordCompleteIndex] || "" : "",
          noteIndex !== -1 ? row[noteIndex] || "" : "",
          caseClosedIndex !== -1 ? row[caseClosedIndex] || "" : "",
          remuneration, transportValue, totalRemuneration, totalCompensation
        ]);
      }
    });

    const startRow = 3;
    if (filteredData.length > 0) {
      reportSheet.getRange(startRow, 1, filteredData.length, 19).setValues(filteredData);
      // 格式化數字欄位
      reportSheet.getRange(startRow, 16, filteredData.length, 4)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");

      // 總計行
      const summaryRow = startRow + filteredData.length;
      const totalVisitCount = filteredData.reduce((sum, row) => sum + (Number(row[9]) || 0), 0);
      const totalRemuneration = filteredData.reduce((sum, row) => sum + (row[15] || 0), 0);
      const totalTransport = filteredData.reduce((sum, row) => sum + (row[16] || 0), 0);
      const totalCompensation = filteredData.reduce((sum, row) => sum + (row[18] || 0), 0);
      reportSheet.getRange(summaryRow, 1, 1, 19).setValues([[
        "總計", "", "", "", "", "", "", "", "", totalVisitCount, "", "", "", "", "",
        totalRemuneration, totalTransport, totalCompensation, totalCompensation
      ]]).setFontWeight("bold").setBackground("#e6f3ff");
      reportSheet.getRange(summaryRow, 16, 1, 4).setNumberFormat("#,##0").setHorizontalAlignment("right");

      Logger.log(`📊 已填充 ${filteredData.length} 筆資料並生成總計行`);
      createChart(reportSheet, summaryRow + 1, ownerFilter || "所有");
    } else {
      reportSheet.getRange(startRow, 1).setValue("無符合篩選條件的數據");
      Logger.log("⚠ 無符合篩選條件的數據");
    }
  } catch (error) {
    Logger.log(`🚨 updateReportSummarySheet 執行錯誤: ${error.message}`);
    reportSheet.getRange(3, 1).setValue(`錯誤: ${error.message}`);
  }
}

/**************************
 * 手動更新報告總表（選單觸發）
 **************************/
function manualUpdateReport() {
  updateReportSummarySheet(null);
}

/**************************
 * 創建圖表
 **************************/
function createChart(reportSheet, startRow, ownerFilter) {
  try {
    const charts = reportSheet.getCharts();
    charts.forEach(chart => reportSheet.removeChart(chart));

    const dataRange = reportSheet.getRange(3, 1, reportSheet.getLastRow() - 3, 19);
    const dataValues = dataRange.getValues();
    if (dataValues.length > 0 && dataValues[0].some(cell => cell !== "")) {
      const chart = reportSheet.newChart()
        .addRange(reportSheet.getRange(reportSheet.getLastRow(), 9, 1, 3)) // J 列到 L 列
        .setChartType(Charts.ChartType.COLUMN)
        .setPosition(5, 1, 0, 0)
        .setOption('title', `負責人 ${ownerFilter || '所有'} 數據概覽`)
        .setOption('hAxis.title', '項目')
        .setOption('vAxis.title', '數值')
        .setOption('legend', { position: 'right' })
        .build();

      reportSheet.insertChart(chart);
      Logger.log("📊 圖表創建成功");
    } else {
      Logger.log("⚠ 無數據，跳過圖表創建");
    }
  } catch (error) {
    Logger.log(`⚠ createChart 執行錯誤: ${error.message}`);
  }
}

/**************************
 * 設置觸發器（支持指定試算表）
 **************************/
function setupReportTrigger(spreadsheetId = TARGET_SHEET_ID) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === "onEdit" && t.getTriggerSourceId() === spreadsheetId);
  if (!existingTrigger) {
    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(spreadsheetId)
      .onEdit()
      .create();
    Logger.log(`ℍ 已為試算表 ${spreadsheetId} 設置 onEdit 觸發器`);
  }
}

function setupHourlyTrigger(spreadsheetId = TARGET_SHEET_ID) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === "generateMonthlyRemunerationSheets" && trigger.getTriggerSourceId() === spreadsheetId) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger("generateMonthlyRemunerationSheets")
    .timeBased()
    .everyHours(1)
    .inTimezone("Asia/Taipei")
    .create();
  Logger.log(`ℍ 已為試算表 ${spreadsheetId} 設置每小時自動掃描觸發器`);
}

/**************************
 * 測試函數
 **************************/
function testReportInitialization() {
  initializeReportSummarySheet();
  setupReportTrigger();
}

/**************************
 * 清理觸發器（僅為安全起見）
 **************************/
function cleanTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`ℍ 當前觸發器數量: ${triggers.length}`);
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log(`ℍ 已刪除觸發器: ${trigger.getUniqueId()}`);
  });
  Logger.log(`ℍ 清理完成，剩餘觸發器數量: ${ScriptApp.getProjectTriggers().length}`);
}

function testSetupTrigger() {
  setupReportTrigger(TARGET_SHEET_ID);
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`ℍ 當前觸發器: ${triggers.map(t => `${t.getHandlerFunction()} - ${t.getTriggerSourceId()}`).join(", ")}`);
}
