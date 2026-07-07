/**
 * Code.gs
 * 🏭 鍋爐設備安裝工程自動化管理系統 - 後端入口點
 * Google Apps Script V8 Engine
 * PM 專案管理功能
 * 現場派工管理
 */

const SCRIPT_VERSION = "v3.1.0";

/**
 * HTTP GET 請求入口 (用以渲染介面或取得資料)
 */
function doGet(e) {
  // 如果帶有 action 參數，當作 API 呼叫
  if (e.parameter.action) {
    try {
      const result = handleRoute(e.parameter);
      return ContentService.createTextOutput(
        JSON.stringify(result),
      ).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(
        JSON.stringify({
          success: false,
          error: error.message || error.toString(),
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  let webAppUrl = "";
  try {
    webAppUrl = ScriptApp.getService().getUrl() || "";
  } catch (err) {
    console.warn("無法取得 ScriptApp.getService().getUrl()", err);
  }

  let titleSuffix = "";
  const last3 = webAppUrl.slice(-3).toLowerCase();
  if (last3 === "dev" || webAppUrl.toLowerCase().indexOf("dev") !== -1) {
    titleSuffix = " (測試版)";
  }

  // 檢查是否要求 PM 專案管理介面
  if (e.parameter.page === "pm") {
    return HtmlService.createHtmlOutputFromFile("PM")
      .setTitle("📊" + titleSuffix + " BISMS 專案管理系統")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // 允許 iframe 嵌入
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 否則回傳 UI 介面
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("🏭" + titleSuffix + " BISMS 現場管理系統")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // 允許 iframe 嵌入
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * HTTP POST 請求入口 (用以新增、更新資料或上傳照片)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }

    const result = handleRoute(payload);

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
      ContentService.MimeType.JSON,
    );
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        error: error.message || error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 統一路由處理 (API Gateway)
 */
function handleRoute(params) {
  const { action } = params;

  const requiresLock = [
    "submitResult",
    "updateTaskProcessStatus",
    "createTask",
    "batchCreateTasks",
    "generateReport",
    "generateTrackingReport",
    "deletePhoto",
    "renameComponent",
  ].includes(action);
  const lock = requiresLock ? LockService.getScriptLock() : null;

  if (lock) {
    lock.waitLock(15000); // 15秒超時
  }

  try {
    let result;
    switch (action) {
      case "getInitData":
        result = apiGetInitData(params);
        break;
      case "getPmInitData":
        result = apiGetPmInitData(params);
        break;
      case "updatePmRow":
        result = apiUpdatePmRow(params);
        break;
      case "batchUpdatePmRows":
        result = apiBatchUpdatePmRows(params);
        break;
      case "submitResult":
        result = apiSubmitResult(params);
        break;
      case "renameComponent":
        result = apiRenameComponent(params);
        break;
      case "finishTask":
        result = apiFinishTask(params);
        break;
      case "generateReport":
        result = apiGenerateReport(params);
        break;
      case "deletePhoto":
        result = apiDeletePhoto(params);
        break;
      case "approveTask":
        result = apiApproveTask(params);
        break;
      case "rejectTask":
        result = apiRejectTask(params);
        break;
      case "closeTask":
        result = apiCloseTask(params);
        break;
      case "updateTaskProcessStatus":
        result = apiUpdateTaskProcessStatus(params);
        break;
      case "createTask":
        result = apiCreateTask(params);
        break;
      case "batchCreateTasks":
        result = apiBatchCreateTasks(params);
        break;
      case "updateTaskFields":
        result = apiUpdateTaskFields(params);
        break;
      case "generateTrackingReport":
        result = apiGenerateTrackingReport(params);
        break;
      default:
        throw new Error("不支援的 Action: " + action);
    }
    console.log(action, "執行完成");
    return {
      success: true,
      data: result,
    };
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

// ============================================
// API Controllers
// ============================================

function sanitizeText(str) {
  if (typeof str !== "string") return str;
  return str.replace(/[&<>"']/g, function (m) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    }[m];
  });
}

/**
 * [輔助函式] 取得或建立指定的 Google Drive 資料夾，並寫入屬性快取
 */
function getOrCreateDriveFolder(folderName, propertyKey) {
  const scriptProps = PropertiesService.getScriptProperties();
  let folderId = propertyKey ? scriptProps.getProperty(propertyKey) : null;
  let folder;

  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      folderId = null; // ID 失效
    }
  }

  if (!folderId) {
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(currentSpreadsheet.getId()).getParents().next();
    const folders = parentFolder.getFoldersByName(folderName);

    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parentFolder.createFolder(folderName);
      try {
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        console.error("無法設定目錄權限", e);
      }
    }
    if (propertyKey) {
      scriptProps.setProperty(propertyKey, folder.getId());
    }
  }
  return folder;
}

/**
 * 取得 PM_ 工作表初始化所需的所有元數據與行
 */
function apiGetPmInitData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const pmSheets = [];
  sheets.forEach(s => {
    const name = s.getName();
    if (name.indexOf("PM_") === 0) {
      pmSheets.push(name);
    }
  });

  const staff = DataService.getCachedOrLive(CONFIG.SHEETS.STAFF);

  let selectedSheetRows = [];
  let headers = [];
  const selectedSheetName = payload.sheetName || (pmSheets.length > 0 ? pmSheets[0] : "");

  if (selectedSheetName) {
    const sheet = ss.getSheetByName(selectedSheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow >= 1 && lastCol >= 1) {
        const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        headers = rawValues[0].map(h => String(h).trim());

        const dispatchIdx = headers.indexOf("派工確認");
        const taskIdIdx = headers.indexOf("任務ID");
        let validations = [];
        if (dispatchIdx !== -1 && lastRow > 1) {
          try {
            validations = sheet.getRange(2, dispatchIdx + 1, lastRow - 1, 1).getDataValidations();
          } catch (e) {
            console.error("無法取得資料驗證", e);
          }
        }

        for (let i = 1; i < rawValues.length; i++) {
          const rowVals = rawValues[i];
          const hasVal = rowVals.some(v => v !== "" && v !== null);
          if (!hasVal) continue;

          const rowObj = {
            _rowIndex: i + 1
          };
          headers.forEach((header, colIdx) => {
            if (!header) return;
            const val = rowVals[colIdx];
            if (val instanceof Date) {
              rowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd");
            } else {
              rowObj[header] = val;
            }
          });
          selectedSheetRows.push(rowObj);
        }
      }
    }
  }

  let webAppUrl = "";
  try {
    webAppUrl = ScriptApp.getService().getUrl();
  } catch (e) {
    console.warn("無法取得 ScriptApp.getService().getUrl()", e);
  }

  let templates = [];
  let templatesFull = [];
  let checklists = [];
  let results = [];
  try {
    templatesFull = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES) || [];
    templates = templatesFull.map(t => String(t["樣板名稱"] || t[HEADER_MAP.TYPES.NAME] || "").trim()).filter(Boolean);
    checklists = DataService.getCachedOrLive(CONFIG.SHEETS.CHECKLIST) || [];
    results = DataService.getCachedOrLive(CONFIG.SHEETS.RESULTS) || [];
  } catch (err) {
    console.error("無法載入樣板 / 檢查項目 / 結果清單：", err);
  }

  return {
    pmSheets: pmSheets,
    staff: staff,
    templates: templates,
    templatesFull: templatesFull,
    checklists: checklists,
    results: results,
    selectedSheetName: selectedSheetName,
    headers: headers,
    rows: selectedSheetRows,
    webAppUrl: webAppUrl
  };
}

/**
 * 更新指定的 PM_ 欄位資訊並自動觸發 onEdit 連動
 */
function apiUpdatePmRow(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("系統忙碌中，無法取得鎖定，請稍後再試。");
  }

  try {
    const { sheetName, rowIndex, fields } = payload;
    if (!sheetName || !rowIndex) {
      throw new Error("錯誤：缺少工作表名稱或列編號");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("找不到工作表: " + sheetName);
    }

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const rIndex = parseInt(rowIndex, 10);

    const levelColIdx = headers.indexOf("階層") + 1;
    let levelVal = 1;
    if (levelColIdx > 0) {
      const rawLvl = sheet.getRange(rIndex, levelColIdx).getValue();
      levelVal = parseInt(rawLvl, 10) || 1;
    }

    for (const colName in fields) {
      if (colName === "階層" || colName === "WBS") {
        continue;
      }
      const colIdx = headers.indexOf(colName) + 1;
      if (colIdx > 0) {
        let val = fields[colName];

        if (colName === "任務名稱" && val !== undefined && val !== null) {
          val = String(val).trim();
          const numSpaces = Math.max(0, (levelVal - 1) * 2);
          const spaces = " ".repeat(numSpaces);
          val = spaces + val;
        }

        const cell = sheet.getRange(rIndex, colIdx);

        // Handle boolean for checkbox values like "派工確認"
        if (colName === "派工確認") {
          if (val === "true" || val === "TRUE" || val === true) {
            val = true;
          } else if (val === "false" || val === "FALSE" || val === false || val === "") {
            val = false;
          }
        }

        cell.setValue(val);
      }
    }

    // 一次性直接呼叫 PmDispatchService.handleEdit 以進行完美雙向同步，不受 onEdit Lock 機制干擾
    try {
      PmDispatchService.handleEdit(sheet, rIndex, rIndex, 1, lastCol, headers, null);
    } catch (err) {
      console.error("雙向同步功能(PmDispatchService)執行失敗:", err);
    }

    SpreadsheetApp.flush();

    // 獲取已更新的資料
    const updatedRowObj = { _rowIndex: rIndex };
    if (lastCol >= 1) {
      const rawRowVals = sheet.getRange(rIndex, 1, 1, lastCol).getValues()[0];
      headers.forEach((header, colIdx) => {
        if (!header) return;
        const val = rawRowVals[colIdx];
        if (val instanceof Date) {
          updatedRowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else {
          updatedRowObj[header] = val;
        }
      });
    }

    return {
      success: true,
      message: "資料已更新並觸發雙向同步",
      updatedRow: updatedRowObj
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 批次更新指定的 PM_ 欄位資訊並自動觸發 handleEdit 連動
 */
function apiBatchUpdatePmRows(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("系統忙碌中，無法取得鎖定，請稍後再試。");
  }

  try {
    const { sheetName, updates } = payload;
    if (!sheetName || !updates || !Array.isArray(updates)) {
      throw new Error("Missing sheetName or updates array");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol <= 0) {
      return { success: true, message: "No columns on sheet", updatedRows: [] };
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const levelColIdx = headers.indexOf("階層") + 1;
    const updatedRows = [];

    for (const update of updates) {
      const { rowIndex, fields } = update;
      if (!rowIndex) continue;

      const rIndex = parseInt(rowIndex, 10);
      let levelVal = 1;
      if (levelColIdx > 0) {
        const rawLvl = sheet.getRange(rIndex, levelColIdx).getValue();
        levelVal = parseInt(rawLvl, 10) || 1;
      }

      for (const colName in fields) {
        if (colName === "階層" || colName === "WBS") {
          continue;
        }
        const colIdx = headers.indexOf(colName) + 1;
        if (colIdx > 0) {
          let val = fields[colName];

          if (colName === "任務名稱" && val !== undefined && val !== null) {
            val = String(val).trim();
            const numSpaces = Math.max(0, (levelVal - 1) * 2);
            const spaces = " ".repeat(numSpaces);
            val = spaces + val;
          }

          // Handle boolean for checkbox values like "派工確認"
          if (colName === "派工確認") {
            if (val === "true" || val === "TRUE" || val === true) {
              val = true;
            } else if (val === "false" || val === "FALSE" || val === false || val === "") {
              val = false;
            }
          }

          sheet.getRange(rIndex, colIdx).setValue(val);
        }
      }

      // 呼叫 PmDispatchService 以自適應處理單筆更新與衍生指派
      try {
        PmDispatchService.handleEdit(sheet, rIndex, rIndex, 1, lastCol, headers, null);
      } catch (err) {
        console.error("Direct trigger of PmDispatchService.handleEdit failed inside batch updates for rowIndex " + rIndex + ":", err);
      }

      // 讀回該 rowIndex 真正最新的資料以利前端樂觀 UI 合併
      const updatedRowObj = { _rowIndex: rIndex };
      const rawRowVals = sheet.getRange(rIndex, 1, 1, lastCol).getValues()[0];
      headers.forEach((header, colIdx) => {
        if (!header) return;
        const val = rawRowVals[colIdx];
        if (val instanceof Date) {
          updatedRowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else {
          updatedRowObj[header] = val;
        }
      });
      updatedRows.push(updatedRowObj);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Batch rows updated successfully",
      updatedRows: updatedRows
    };
  } finally {
    lock.releaseLock();
  }
}



/**
 * 取得初始化所需的所有元數據 (優化版：全面快取 + Set 過濾)
 */
function apiGetInitData(params) {
  // 1. 取得所有原始數據 (全面使用 CacheService)
  const allChecklists = DataService.getCachedOrLive(CONFIG.SHEETS.CHECKLIST);
  const allStaff = DataService.getCachedOrLive(CONFIG.SHEETS.STAFF);
  const allTasks = DataService.getCachedOrLive(CONFIG.SHEETS.TASKS);
  const allResults = DataService.getCachedOrLive(CONFIG.SHEETS.RESULTS);
  const allTemplates = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);

  // 2. 實作派工篩選邏輯
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  const activeTasks = allTasks.filter((task) => {
    const assignDateStr = task[HEADER_MAP.TASKS.ASSIGN_DATE];
    let status = task[HEADER_MAP.TASKS.REPORT_STATUS] || "未開始";

    let isDateValid = false;
    if (assignDateStr !== "") {
      const assignDate = new Date(assignDateStr);
      isDateValid = assignDate <= today;
    }

    const activeStatuses = ["未開始", "執行中", "已提報結果", "退回修改", "已審查結果", "已生成報告"];
    const isStatusValid = activeStatuses.includes(status);
    return isDateValid && isStatusValid;
  });

  // 3. 連動過濾檢查結果 (優化：使用 Set 將 O(n^2) 降為 O(n))
  const activeTaskIdSet = new Set(activeTasks.map((t) => String(t[HEADER_MAP.TASKS.ID])));
  const activeResults = allResults.filter((r) =>
    activeTaskIdSet.has(String(r[HEADER_MAP.RESULTS.TASK_ID]))
  );

  let webAppUrl = "";
  try {
    webAppUrl = ScriptApp.getService().getUrl();
  } catch (e) {
    console.warn("無法取得 ScriptApp.getService().getUrl()", e);
  }

  return {
    checklists: allChecklists,
    staff: allStaff,
    tasks: activeTasks,
    results: activeResults,
    templates: allTemplates,
    projects: [],
    webAppUrl: webAppUrl
  };
}

/**
 * 提交單項檢查結果 (Idempotency 機制)
 * 預期 payload: { taskId, checklistId, value, result, date, base64Photo? }
 */
function apiSubmitResult(payload) {
  const {
    taskId, checklistId, value, result, date, base64Photo,
    taskName = "", templateName = "", checkItemTag = "", subItemName = "", taskCategory = "",
    componentId = ""
  } = payload;

  if (!taskId || !checklistId) {
    throw new Error("Missing taskId or checklistId");
  }

  // 1. 組合要寫入 RESULTS 的紀錄
  const reportDate = date ? date : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const uploadDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

  const compId = componentId || payload.COMPONENT_ID || "";
  const recordId = compId ? taskId + "_" + checklistId + "_" + compId : taskId + "_" + checklistId;

  // 取出原本紀錄
  const existingRecord = DataService.getRecordByKey(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, recordId) || {};
  let photoFileId = existingRecord[HEADER_MAP.RESULTS.PHOTO_ID] || "";

  // 2. 照片處理 (若有附圖)
  if (base64Photo) {
    // 前端已直接傳遞組合檔名所需的資訊，免去 2 次耗時的試算表查表！
    const existingIds = (photoFileId || "").split(",").filter(Boolean);
    const timestampSuffix = "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmmssSSS");
    const fileName = `${taskId}_${templateName}_${checkItemTag}_${taskName}_${subItemName}_${taskCategory}_${uploadDate}${timestampSuffix}.jpg`;

    // 尋找或建立「照片」資料夾
    const photoFolder = getOrCreateDriveFolder("照片", "CONFIG_PHOTO_FOLDER_ID");

    // 處理 base64 padding
    const base64Data = base64Photo.split(",")[1] || base64Photo;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", fileName);
    const newFile = photoFolder.createFile(blob);
    const newFileId = newFile.getId();
    photoFileId = photoFileId ? photoFileId + "," + newFileId : newFileId;
    console.log("Photo uploaded: " + newFileId);
  }

  // 對 value, result 做合併 (如果前端此次未傳遞，則保留現有資料)
  let finalValue = value !== undefined ? value : (existingRecord[HEADER_MAP.RESULTS.ACTUAL_VAL] || "");
  if (typeof finalValue === "string") {
    // 直接對所有字串做防呆，省去查 CHECKLIST 表邏輯型態的麻煩
    finalValue = sanitizeText(finalValue);
  }
  const finalResult = result !== undefined ? result : (existingRecord[HEADER_MAP.RESULTS.RESULT] || "");

  const recordObj = {
    [HEADER_MAP.RESULTS.ID]: recordId,
    [HEADER_MAP.RESULTS.TASK_ID]: taskId,
    [HEADER_MAP.RESULTS.CHECK_ID]: checklistId,
    [HEADER_MAP.RESULTS.COMPONENT_ID]: compId,
    [HEADER_MAP.RESULTS.ACTUAL_VAL]: finalValue,
    [HEADER_MAP.RESULTS.RESULT]: finalResult,
    [HEADER_MAP.RESULTS.REPORT_DATE]: reportDate,
    [HEADER_MAP.RESULTS.PHOTO_ID]: photoFileId,
  };

  // 3. 冪等覆寫 (以 記錄ID 為主鍵)
  DataService.upsertRecord(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, recordId, recordObj);

  // 4. 連動更新 TASKS 狀態為「執行中」
  const taskInfo = DataService.getRecordByKey(CONFIG.SHEETS.TASKS, HEADER_MAP.TASKS.ID, taskId);
  let newTaskStatus = taskInfo ? taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] : null;
  let newProcessStatus = taskInfo ? taskInfo[HEADER_MAP.TASKS.PROCESS_STATUS] : null;

  if (taskInfo && (newTaskStatus !== "執行中" || newTaskStatus === "未開始" || newTaskStatus === "退回修改" || !newTaskStatus)) {
    newTaskStatus = "執行中";
    newProcessStatus = "持續進行";
    DataService.updateRecordFields(CONFIG.SHEETS.TASKS, HEADER_MAP.TASKS.ID, taskId, {
      [HEADER_MAP.TASKS.REPORT_STATUS]: newTaskStatus,
      [HEADER_MAP.TASKS.PROCESS_STATUS]: newProcessStatus
    });
  }

  return {
    message: "Result saved successfully",
    recordId: recordId,
    photoId: photoFileId,
    newTaskStatus: newTaskStatus,
    newProcessStatus: newProcessStatus,
  };
}

/**
 * 重命名零件實例
 * 預期 payload: { taskId, checklistId, oldId, newId }
 */
function apiRenameComponent(payload) {
  const { taskId, checklistId, oldId, newId } = payload;

  if (!taskId || !checklistId || !newId) {
    throw new Error("Missing required parameters for renaming");
  }

  const oldIdVal = oldId || "";
  let oldRecordId = oldIdVal ? taskId + "_" + checklistId + "_" + oldIdVal : taskId + "_" + checklistId;
  const newRecordId = taskId + "_" + checklistId + "_" + newId;

  // 1. 防呆：檢查 newId 是否重複
  const duplicate = DataService.getRecordByKey(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, newRecordId);
  if (duplicate) {
    throw new Error("該零件名稱已存在，請使用其他名稱。");
  }

  // 2. 檢查舊紀錄是否存在，若不存在且 oldIdVal 不為空，嘗試尋找未設定零件實例的舊紀錄 (自癒機制)
  let recordExists = DataService.getRecordByKey(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, oldRecordId);
  if (!recordExists && oldIdVal) {
    const fallbackRecordId = taskId + "_" + checklistId;
    const fallbackRecord = DataService.getRecordByKey(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, fallbackRecordId);
    if (fallbackRecord && (!fallbackRecord[HEADER_MAP.RESULTS.COMPONENT_ID] || String(fallbackRecord[HEADER_MAP.RESULTS.COMPONENT_ID]).trim() === "")) {
      oldRecordId = fallbackRecordId;
    }
  }

  // 3. 局部更新
  const success = DataService.updateRecordFields(CONFIG.SHEETS.RESULTS, HEADER_MAP.RESULTS.ID, oldRecordId, {
    [HEADER_MAP.RESULTS.ID]: newRecordId,
    [HEADER_MAP.RESULTS.COMPONENT_ID]: newId
  });

  if (!success) {
    throw new Error("找不到原有的零件紀錄，編輯失敗。");
  }

  return {
    success: true,
    message: "Component instance renamed successfully",
    oldRecordId: oldRecordId,
    newRecordId: newRecordId,
  };
}

/**
 * 刪除結果紀錄或任務備註中的單張照片
 * 預期 payload: { taskId, checklistId?, photoIdToDelete }
 * 如果 checklistId 為 "REMARK"，則刪除任務備註中的照片
 */
function apiDeletePhoto(payload) {
  const { taskId, checklistId, photoIdToDelete } = payload;

  if (!taskId || !photoIdToDelete) {
    throw new Error("Missing taskId or photoIdToDelete");
  }

  let photoFileIds = "";

  if (checklistId === "REMARK") {
    // 刪除任務備註中的照片
    const taskInfo = DataService.getRecordByKey(
      CONFIG.SHEETS.TASKS,
      HEADER_MAP.TASKS.ID,
      taskId,
    );
    if (!taskInfo) throw new Error("Task not found");

    let currentPhotoIds = (taskInfo[HEADER_MAP.TASKS.REMARK_PHOTO_ID] || "")
      .split(",")
      .filter(Boolean);
    const updatedPhotoIds = currentPhotoIds.filter(
      (id) => id !== photoIdToDelete,
    );
    photoFileIds = updatedPhotoIds.join(",");

    TaskService.update(taskId, { [HEADER_MAP.TASKS.REMARK_PHOTO_ID]: photoFileIds });
  } else {
    // 刪除檢查項目中的照片
    if (!checklistId) throw new Error("Missing checklistId");

    const compId = payload.componentId || payload.COMPONENT_ID || "";
    const recordId = compId ? taskId + "_" + checklistId + "_" + compId : taskId + "_" + checklistId;
    const existingRecord = DataService.getRecordByKey(
      CONFIG.SHEETS.RESULTS,
      HEADER_MAP.RESULTS.ID,
      recordId,
    );
    if (!existingRecord) {
      throw new Error("Result record not found.");
    }

    let currentPhotoIds = (existingRecord[HEADER_MAP.RESULTS.PHOTO_ID] || "")
      .split(",")
      .filter(Boolean);
    const updatedPhotoIds = currentPhotoIds.filter(
      (id) => id !== photoIdToDelete,
    );
    photoFileIds = updatedPhotoIds.join(",");

    // 更新記錄
    existingRecord[HEADER_MAP.RESULTS.PHOTO_ID] = photoFileIds;

    const checklistInfo = DataService.getRecordByKey(
      CONFIG.SHEETS.CHECKLIST,
      HEADER_MAP.CHECKLIST.ID,
      checklistId,
    );
    if (
      checklistInfo &&
      checklistInfo[HEADER_MAP.CHECKLIST.LOGIC_TYPE] === "相片"
    ) {
      existingRecord[HEADER_MAP.RESULTS.ACTUAL_VAL] = photoFileIds;
      if (updatedPhotoIds.length === 0) {
        existingRecord[HEADER_MAP.RESULTS.RESULT] = "";
      }
    }

    DataService.upsertRecord(
      CONFIG.SHEETS.RESULTS,
      HEADER_MAP.RESULTS.ID,
      recordId,
      existingRecord,
    );
  }

  // 嘗試將該檔案移至垃圾桶
  try {
    const file = DriveApp.getFileById(photoIdToDelete);
    if (file) {
      file.setTrashed(true);
    }
  } catch (e) {
    console.error(
      "Failed to trash file (might already be deleted or permission denied):",
      e,
    );
  }

  return { message: "Photo deleted", photoIds: photoFileIds };
}

/**
 * 產出 Google Docs 報告 (Idempotency 機制)
 * 預期 payload: { taskId }
 */
function apiGenerateReport(payload) {
  const { taskId, bypassStatusCheck = false } = payload;
  if (!taskId) throw new Error("Missing taskId");

  let taskInfo = DataService.getRecordByKey(
    CONFIG.SHEETS.TASKS,
    HEADER_MAP.TASKS.ID,
    taskId,
  );
  if (!taskInfo) throw new Error("Task not found");

  if (
    !bypassStatusCheck &&
    taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] !== "已提交結果" &&
    taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] !== "已提報結果" &&
    taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] !== "已審查結果" &&
    taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] !== "已生成報告" &&
    taskInfo[HEADER_MAP.TASKS.REPORT_STATUS] !== "已結案"
  ) {
    throw new Error(
      "任務狀態必須為 '已提報結果' 或 '已審查結果' 才能產出報告。目前狀態: " +
      taskInfo[HEADER_MAP.TASKS.REPORT_STATUS],
    );
  }

  const templateCode = taskInfo[HEADER_MAP.TASKS.TEMPLATE_CODE];
  const codes = String(templateCode || "").split(",").map(s => s.trim()).filter(Boolean);
  const primaryCode = codes[0] || "";
  let templateInfo = null;
  if (primaryCode) {
    templateInfo = DataService.getRecordByKey(
      CONFIG.SHEETS.TYPES,
      HEADER_MAP.TYPES.CODE,
      primaryCode,
    );
  }
  if (!templateInfo)
    throw new Error("Template not found for code: " + primaryCode);

  const templateDocId = templateInfo[HEADER_MAP.TYPES.REPORT_ID];
  if (!templateDocId || templateDocId === "DOC_ID_PLACEHOLDER")
    throw new Error("Invalid or missing Template Doc ID");

  const uploadDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd",
  );
  const templateName = taskInfo[HEADER_MAP.TASKS.TEMPLATE_NAME] || "";
  const taskName = taskInfo[HEADER_MAP.TASKS.NAME] || "";
  const subItemName = taskInfo[HEADER_MAP.TASKS.SUB_ITEM] || "";
  const taskCategory = taskInfo[HEADER_MAP.TASKS.CATEGORY] || "";

  // 報告檔名規範: [任務ID]_[樣板名稱]_[任務名稱]_[任務分項]_[任務分類]_[YYYYMMDD]
  const reportFileName = `${taskId}_${templateName}_${taskName}_${subItemName}_${taskCategory}_${uploadDate}`;

  // 尋找或建立「報告」資料夾
  const reportFolder = getOrCreateDriveFolder("報告", "CONFIG_REPORT_FOLDER_ID");

  // Idempotent: 刪除既有相同名稱的報告
  const existingFiles = reportFolder.getFilesByName(reportFileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // 複製範本
  const templateFile = DriveApp.getFileById(templateDocId);
  const newReportFile = templateFile.makeCopy(reportFileName, reportFolder);
  const newReportDoc = DocumentApp.openById(newReportFile.getId());
  const body = newReportDoc.getBody();

  // 標籤替換 - 任務基本資訊
  for (const key in taskInfo) {
    body.replaceText(`{{任務_${key}}}`, taskInfo[key] || "");
  }

  // 向後兼容與原有的替換
  body.replaceText("{{任務ID}}", taskId || "");
  body.replaceText("{{任務分類}}", taskCategory || "");
  body.replaceText("{{任務分項}}", subItemName || "");
  body.replaceText("{{任務名稱}}", taskName || "");
  body.replaceText("{{負責人}}", taskInfo[HEADER_MAP.TASKS.ASSIGNEE] || "");
  body.replaceText("{{提報日期}}", taskInfo[HEADER_MAP.TASKS.SUBMIT_DATE] || "");

  // --- 報表內容處理階段 ---
  const allTasks = DataService.getCachedOrLive(CONFIG.SHEETS.TASKS);
  const allResults = DataService.getCachedOrLive(CONFIG.SHEETS.RESULTS);
  const taskResults = allResults.filter(
    (r) => r[HEADER_MAP.RESULTS.TASK_ID] === taskId,
  );
  const allChecklists = DataService.getCachedOrLive(CONFIG.SHEETS.CHECKLIST);
  const taskChecklists = allChecklists.filter(
    (c) => codes.includes(String(c[HEADER_MAP.CHECKLIST.TEMPLATE_CODE])),
  );

  let photosToAppend = [];
  const processedCheckIds = new Set();

  // 1. 處理個別標籤模式 (Mode A)
  taskChecklists.forEach((checkItem, index) => {
    let tagKey = checkItem[HEADER_MAP.CHECKLIST.REPORT_TAG];
    if (!tagKey) return;

    tagKey = tagKey.replace(/^{{/, "").replace(/}}$/, "");

    // 定義此項目的各種標籤格式
    const placeholders = {
      item: `{{項目_${tagKey}}}`,
      criteria: `{{標準_${tagKey}}}`,
      value: `{{數值_${tagKey}}}`,
      result: `{{結果_${tagKey}}}`,
      photo: `{{照片_${tagKey}}}`,
      legacy: `{{${tagKey}}}`,
    };

    // 檢查樣板中是否存在上述任一標籤，以啟動 Mode A
    const hasTagInTemplate = Object.values(placeholders).some((p) =>
      body.findText(p),
    );

    if (hasTagInTemplate) {
      processedCheckIds.add(checkItem[HEADER_MAP.CHECKLIST.ID]);
      const result = taskResults.find(
        (r) =>
          r[HEADER_MAP.RESULTS.CHECK_ID] === checkItem[HEADER_MAP.CHECKLIST.ID],
      );

      // 文字與數值處理
      let actualValue = result
        ? result[HEADER_MAP.RESULTS.ACTUAL_VAL] || ""
        : "";

      // 邏輯型態為空時，使用 單位或選項 作為預設實際值
      if (actualValue === "" && !checkItem[HEADER_MAP.CHECKLIST.LOGIC_TYPE]) {
        actualValue = checkItem[HEADER_MAP.CHECKLIST.UNIT_OPTS] || "";
      }
      const actualResult = result
        ? result[HEADER_MAP.RESULTS.RESULT] || ""
        : "";
      if (
        checkItem[HEADER_MAP.CHECKLIST.LOGIC_TYPE] === "數值" &&
        actualValue !== "" &&
        checkItem[HEADER_MAP.CHECKLIST.UNIT_OPTS]
      ) {
        actualValue += " " + checkItem[HEADER_MAP.CHECKLIST.UNIT_OPTS];
      }

      // 批量執行取代 (改為一對一單次取代，以防多個同標籤項目時相互覆蓋)
      replaceFirstTagOnly(body, placeholders.item, checkItem[HEADER_MAP.CHECKLIST.ITEM_NAME] || "");
      replaceFirstTagOnly(body, placeholders.criteria, checkItem[HEADER_MAP.CHECKLIST.CRITERIA] || "");
      replaceFirstIndividualTagWithFormat(
        body,
        placeholders.value,
        actualValue,
        "#0000FF",
      );
      replaceFirstIndividualTagWithFormat(
        body,
        placeholders.result,
        actualResult,
        actualResult === "不合格" ? "#FF0000" : null,
      );
      replaceFirstTagOnly(body, placeholders.legacy, actualValue || "無紀錄");

      // 照片處理
      const individualPhotoRange = body.findText(placeholders.photo);
      if (individualPhotoRange) {
        const photoIdsStr = result ? result[HEADER_MAP.RESULTS.PHOTO_ID] : null;
        const para = individualPhotoRange
          .getElement()
          .getParent()
          .asParagraph();

        // 精確刪除照片標籤，避免全域取代影響其他同名標籤
        const element = individualPhotoRange.getElement().asText();
        element.deleteText(individualPhotoRange.getStartOffset(), individualPhotoRange.getEndOffsetInclusive());

        if (photoIdsStr) {
          const photoIds = photoIdsStr.split(",").filter(Boolean);
          insertPhotosToPosition(para, photoIds);
        } else {
          para.appendText("－");
        }
      }
    } else {
      // 若無標籤，則保留 ID 以供 Mode B (動態表格) 處理，並準備照片緩存
      const result = taskResults.find(
        (r) =>
          r[HEADER_MAP.RESULTS.CHECK_ID] === checkItem[HEADER_MAP.CHECKLIST.ID],
      );
      const photoIdsStr = result ? result[HEADER_MAP.RESULTS.PHOTO_ID] : null;
      if (photoIdsStr) {
        const photoIds = photoIdsStr.split(",").filter(Boolean);
        photosToAppend.push({
          checkItem: checkItem,
          photoIds: photoIds,
          index: index + 1,
          photoReference: photoIds
            .map((_, pIdx) => `#${index + 1}-${pIdx + 1}`)
            .join(", "),
        });
      }
    }
  });

  // 2. 處理定錨表格擴充 (Mode B) - 排除 Mode A 已處理之項目
  const remainingChecklists = taskChecklists.filter(
    (c) => !processedCheckIds.has(c[HEADER_MAP.CHECKLIST.ID]),
  );
  expandAnchorTable(body, remainingChecklists, taskResults, photosToAppend);

  // 3. 處理底部照片集追加 (Mode C)
  appendPhotosToAggregateSection(body, photosToAppend);

  const reportUrl = newReportFile.getUrl();

  // 更新 TASKS 的 報告狀態與報告連結 (局部更新，避免覆蓋公式)
  DataService.updateRecordFields(CONFIG.SHEETS.TASKS, HEADER_MAP.TASKS.ID, taskId, {
    //[HEADER_MAP.TASKS.REPORT_STATUS]: "已生成報告",
    [HEADER_MAP.TASKS.PROCESS_STATUS]: "",
    //[HEADER_MAP.TASKS.REPORT_LINK]: `=HYPERLINK("${reportUrl}","報告")`,
    [HEADER_MAP.TASKS.REPORT_LINK]: reportUrl,
  });

  // --- 施工日誌整合：追蹤事項嵌套 ---
  if (subItemName === "施工日誌") {
    const projectName = taskInfo[HEADER_MAP.TASKS.PROJECT_NAME];
    // 篩選邏輯：同專案、分項為追蹤事項、執行進度不為空
    const relatedTrackingTasks = allTasks.filter(t =>
      t[HEADER_MAP.TASKS.PROJECT_NAME] === projectName &&
      (t[HEADER_MAP.TASKS.SUB_ITEM] === "追蹤事項") &&
      (t[HEADER_MAP.TASKS.PROCESS_STATUS] || "").trim() !== ""
    );

    if (relatedTrackingTasks.length > 0) {
      helperInsertTrackingSection(body, relatedTrackingTasks, allChecklists, allResults);
    } else {
      // 若無資料且有標籤，則優雅清除整個卡片區域（包含中間的範本表格）
      body.replaceText("\\{\\{追蹤事項總數\\}\\}", "0");
      const cardStartSearch = body.findText("\\{\\{卡片開始_追蹤\\}\\}");
      const cardEndSearch = body.findText("\\{\\{卡片結束_追蹤\\}\\}");
      if (cardStartSearch && cardEndSearch) {
        const cardStartPara = cardStartSearch.getElement().getParent().asParagraph();
        const cardEndPara = cardEndSearch.getElement().getParent().asParagraph();
        const cardStartIdx = body.getChildIndex(cardStartPara);
        const cardEndIdx = body.getChildIndex(cardEndPara);

        // 倒序刪除所有介於這兩個標記之間的元素（含標記段落與表格）
        for (let i = cardEndIdx; i >= cardStartIdx; i--) {
          body.removeChild(body.getChild(i));
        }
      } else {
        // 備用機制：若找不到成對標記，則純文字替換為空
        body.replaceText("\\{\\{卡片開始_追蹤\\}\\}", "");
        body.replaceText("\\{\\{卡片結束_追蹤\\}\\}", "");
        body.replaceText("\\{\\{執行進度_追蹤\\}\\}", "");
        body.replaceText("\\{\\{任務ID_追蹤\\}\\}", "");
        body.replaceText("\\{\\{任務名稱_追蹤\\}\\}", "");
        body.replaceText("\\{\\{議題描述_追蹤\\}\\}", "");
        body.replaceText("\\{\\{處理過程_追蹤\\}\\}", "");
        body.replaceText("\\{\\{處理結果_追蹤\\}\\}", "");
        body.replaceText("\\{\\{照片_追蹤\\}\\}", "");
        body.replaceText("\\{\\{派工日期_追蹤\\}\\}", "");
        body.replaceText("\\{\\{任務分類_追蹤\\}\\}", "");
        body.replaceText("\\{\\{負責人_追蹤\\}\\}", "");
      }
    }
  }

  newReportDoc.saveAndClose();

  return {
    success: true,
    message: "Report generated",
    reportUrl: reportUrl,
  };
}

/**
 * 提報任務完成 (不產報告，僅更新狀態)
 */
function apiFinishTask(payload) {
  const { taskId } = payload;
  if (!taskId) throw new Error("Missing taskId");

  let taskInfo = DataService.getRecordByKey(
    CONFIG.SHEETS.TASKS,
    HEADER_MAP.TASKS.ID,
    taskId,
  );
  if (!taskInfo) throw new Error("Task not found");

  // 1. 提交當前任務（局部更新狀態與日期）
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  DataService.updateRecordFields(CONFIG.SHEETS.TASKS, HEADER_MAP.TASKS.ID, taskId, {
    [HEADER_MAP.TASKS.REPORT_STATUS]: "已提報結果",
    [HEADER_MAP.TASKS.PROCESS_STATUS]: "處理完成",
    [HEADER_MAP.TASKS.SUBMIT_DATE]: todayStr,
  });

  // 前端已確保關聯工單（子任務）提交後，父工單才可提交，後端無須再連動

  return {
    success: true,
    message: "Task marked as Submitted",
  };
}

/**
 * 審查通過任務 (通過者就產出報告，並更新狀態)
 */
function apiApproveTask(payload) {
  const { taskId } = payload;
  if (!taskId) throw new Error("Missing taskId");
  // 先生成報告 (跳過審查階段的狀態檢查，允許任何狀態被審查通過並生成報告)
  const result = apiGenerateReport({ taskId: taskId, bypassStatusCheck: true });
  if (!result.success) throw new Error(result.message || "Failed to generate report");

  let taskInfo = DataService.getRecordByKey(
    CONFIG.SHEETS.TASKS,
    HEADER_MAP.TASKS.ID,
    taskId,
  );
  if (!taskInfo) throw new Error("Task not found");

  // 局部更新狀態
  TaskService.update(taskId, {
    [HEADER_MAP.TASKS.REPORT_STATUS]: "已審查結果",
    [HEADER_MAP.TASKS.PROCESS_STATUS]: "",
    [HEADER_MAP.TASKS.REVIEW_DATE]: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
  });

  return {
    success: true,
    message: "Task approved",
  };
}

/**
 * 審查退回任務
 */
function apiRejectTask(payload) {
  const { taskId, remark } = payload;
  if (!taskId) throw new Error("Missing taskId");

  let taskInfo = DataService.getRecordByKey(
    CONFIG.SHEETS.TASKS,
    HEADER_MAP.TASKS.ID,
    taskId,
  );
  if (!taskInfo) throw new Error("Task not found");

  const updates = {
    [HEADER_MAP.TASKS.REPORT_STATUS]: "退回修改",
    [HEADER_MAP.TASKS.PROCESS_STATUS]: "等待回覆",
  };

  if (remark) {
    updates[HEADER_MAP.TASKS.REMARK] = remark;
  }

  TaskService.update(taskId, updates);

  return {
    success: true,
    message: "Task rejected",
  };
}

/**
 * 內部輔助：簡化任務更新
 */
const TaskService = {
  update: (taskId, fields) => {
    DataService.updateRecordFields(CONFIG.SHEETS.TASKS, HEADER_MAP.TASKS.ID, taskId, fields);
  }
};

/**
 * 更新任務特定欄位
 */
function apiUpdateTaskFields(payload) {
  const { taskId, fields } = payload;
  if (!taskId || !fields) throw new Error("Missing parameters");

  TaskService.update(taskId, fields);
  return { success: true };
}

/**
 * 結案任務
 */
function apiCloseTask(payload) {
  const { taskId } = payload;
  if (!taskId) throw new Error("Missing taskId");

  let pdfUrl = "";
  try {
    const taskInfo = DataService.getRecordByKey(
      CONFIG.SHEETS.TASKS,
      HEADER_MAP.TASKS.ID,
      taskId
    );

    if (taskInfo) {
      const reportUrl = taskInfo[HEADER_MAP.TASKS.REPORT_LINK];
      if (reportUrl && String(reportUrl).trim() !== "") {
        const getFileIdFromUrl = function (url) {
          if (!url) return null;
          const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
          return match ? match[1] : null;
        };

        const docId = getFileIdFromUrl(reportUrl);
        if (docId) {
          const docFile = DriveApp.getFileById(docId);
          const pdfBlob = docFile.getAs("application/pdf");
          const reportFolder = getOrCreateDriveFolder("報告", "CONFIG_REPORT_FOLDER_ID");

          let pdfName = docFile.getName();
          if (!pdfName.toLowerCase().endsWith(".pdf")) {
            pdfName = pdfName + ".pdf";
          }

          // Idempotent: delete existing PDF files with the same name
          const existingPdfFiles = reportFolder.getFilesByName(pdfName);
          while (existingPdfFiles.hasNext()) {
            existingPdfFiles.next().setTrashed(true);
          }

          const pdfFile = reportFolder.createFile(pdfBlob);
          pdfFile.setName(pdfName);
          try {
            pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
          } catch (sharingErr) {
            console.error("Failed to set PDF sharing: " + sharingErr.toString());
          }
          pdfUrl = pdfFile.getUrl();
        }
      }
    }
  } catch (e) {
    console.error("Failed to convert Doc to PDF on close task: " + e.toString());
  }

  const updates = {
    [HEADER_MAP.TASKS.REPORT_STATUS]: "已結案",
    [HEADER_MAP.TASKS.PROCESS_STATUS]: "",
    [HEADER_MAP.TASKS.CLOSE_DATE]: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
  };

  if (pdfUrl) {
    updates[HEADER_MAP.TASKS.CLOSE_LINK] = pdfUrl;
  }

  TaskService.update(taskId, updates);

  return { success: true, message: "Task closed", closeLink: pdfUrl };
}

/**
 * 更新執行進度
 */
function apiUpdateTaskProcessStatus(payload) {
  const { taskId, status } = payload;
  if (!taskId) throw new Error("Missing taskId");

  DataService.updateRecordFields(
    CONFIG.SHEETS.TASKS,
    HEADER_MAP.TASKS.ID,
    taskId,
    { [HEADER_MAP.TASKS.PROCESS_STATUS]: status }
  );

  return { success: true, message: "Process status updated" };
}

function apiCreateTask(payload) {
  const { taskData, parentTaskId } = payload;
  if (!taskData) throw new Error("Missing taskData");

  // 補強優化：如果傳入有樣板代碼，但部分欄位空缺，自動從樣板資料庫中補齊 (包含樣板名稱、任務分類、任務分項、負責人)
  const tplCode = taskData[HEADER_MAP.TASKS.TEMPLATE_CODE];
  if (tplCode) {
    const templates = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
    const codes = String(tplCode).split(',').map(s => s.trim()).filter(Boolean);
    const matchedTpls = templates.filter(t => codes.includes(String(t[HEADER_MAP.TYPES.CODE])));
    if (matchedTpls.length > 0) {
      if (!taskData[HEADER_MAP.TASKS.TEMPLATE_NAME]) {
        taskData[HEADER_MAP.TASKS.TEMPLATE_NAME] = matchedTpls.map(t => t[HEADER_MAP.TYPES.NAME]).filter(Boolean).join(',');
      }
      const firstMatched = matchedTpls[0];
      const mappings = [
        { taskKey: HEADER_MAP.TASKS.PROJECT_NAME, typeKey: HEADER_MAP.TYPES.PROJECT_NAME },
        { taskKey: HEADER_MAP.TASKS.CATEGORY, typeKey: HEADER_MAP.TYPES.CATEGORY },
        { taskKey: HEADER_MAP.TASKS.SUB_ITEM, typeKey: HEADER_MAP.TYPES.SUB_ITEM },
        { taskKey: HEADER_MAP.TASKS.ASSIGNEE, typeKey: HEADER_MAP.TYPES.ASSIGNEE }
      ];
      mappings.forEach(m => {
        if (!taskData[m.taskKey] && firstMatched[m.typeKey]) {
          taskData[m.taskKey] = firstMatched[m.typeKey];
        }
      });
    }
  }

  const newId = IdService.nextId(CONFIG.SHEETS.TASKS, "T");
  taskData[HEADER_MAP.TASKS.ID] = newId;

  DataService.appendRecord(CONFIG.SHEETS.TASKS, taskData);

  // 優化：合併處理追蹤事項的父任務更新 (關聯工單回寫)，省去前端第二次 GAS 呼叫
  if (parentTaskId) {
    TaskService.update(parentTaskId, { "關聯工單": newId });
  }

  return { success: true, message: "Task created", taskId: newId };
}

/**
 * 批次建立任務
 */
function apiBatchCreateTasks(payload) {
  const { tasksArray } = payload;
  if (!tasksArray || !Array.isArray(tasksArray)) throw new Error("Invalid tasks array");

  const sheet = DataService.getSheet(CONFIG.SHEETS.TASKS);
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0];

  const { YYMM, maxNum } = IdService.getLatestSequence(CONFIG.SHEETS.TASKS, "T");
  let currentMax = maxNum;

  const rowsToAppend = [];
  const createdIds = [];

  for (let i = 0; i < tasksArray.length; i++) {
    currentMax++;
    const newId = "T" + YYMM + ("000" + currentMax).slice(-3);
    const taskData = tasksArray[i];
    taskData[HEADER_MAP.TASKS.ID] = newId;
    createdIds.push(newId);

    const newRow = headers.map(header => {
      let val = taskData[header] ?? "";
      if (typeof val === "string" && val.includes("T") && !isNaN(Date.parse(val))) {
        val = new Date(val);
      }
      return val;
    });
    rowsToAppend.push(newRow);
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    DataService.clearCache(CONFIG.SHEETS.TASKS); // 清除快取，確保資料最新
  }

  return { success: true, message: `Batch created ${rowsToAppend.length} tasks`, createdIds };
}

/**
 * [輔助函式] 渲染追蹤事項清單至指定 Doc Body
 * 搜尋 {{卡片開始_追蹤}} 與 {{卡片結束_追蹤}} 之間的表格作為範本進行循環插入
 */
function helperInsertTrackingSection(body, trackingTasks, allChecklists, allResults) {
  const cardStartSearch = body.findText("\\{\\{卡片開始_追蹤\\}\\}");
  const cardEndSearch = body.findText("\\{\\{卡片結束_追蹤\\}\\}");

  if (!cardStartSearch || !cardEndSearch) return false;

  const cardStartPara = cardStartSearch.getElement().getParent().asParagraph();
  const cardEndPara = cardEndSearch.getElement().getParent().asParagraph();
  const cardStartIdx = body.getChildIndex(cardStartPara);
  const cardEndIdx = body.getChildIndex(cardEndPara);

  // 1. 擷取範本表格
  let templateTable = null;
  for (let i = cardStartIdx + 1; i < cardEndIdx; i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.TABLE) {
      templateTable = el;
      break;
    }
  }
  if (!templateTable) return false;

  // 2. 移除原始標記與範本
  for (let i = cardEndIdx; i >= cardStartIdx; i--) {
    body.removeChild(body.getChild(i));
  }

  // 3. 準備分組與排序
  const STATUS_ORDER = ["等待回覆", "暫緩處理", "持續進行", "處理完成"];
  const groups = {};
  STATUS_ORDER.forEach(s => { groups[s] = []; });

  trackingTasks.forEach(task => {
    const s = task[HEADER_MAP.TASKS.PROCESS_STATUS] || "持續進行";
    if (groups[s]) {
      groups[s].push(task);
    } else {
      // 若有預期外的狀態，歸類到第一項或建立新組
      if (!groups["其他"]) groups["其他"] = [];
      groups["其他"].push(task);
    }
  });

  let insertIdx = cardStartIdx;
  const TRACKING_TAGS = ["議題描述", "處理過程", "處理結果", "佐證照片"];

  // 4. 依照順序逐組插入
  const finalOrder = [...STATUS_ORDER, "其他"];
  const heading0 = body.insertParagraph(insertIdx++, `追蹤事項`);
  heading0.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  finalOrder.forEach(statusKey => {
    const groupTasks = groups[statusKey];
    if (!groupTasks || groupTasks.length === 0) return;

    // 插入分組標題
    const heading = body.insertParagraph(insertIdx++, `執行進度：${statusKey}`);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    groupTasks.forEach(task => {
      const taskId = task[HEADER_MAP.TASKS.ID];
      const templateCode = task[HEADER_MAP.TASKS.TEMPLATE_CODE];

      // 建立該任務的結果對應表
      const checkItems = {};
      allChecklists
        .filter(c => c[HEADER_MAP.CHECKLIST.TEMPLATE_CODE] === templateCode &&
          TRACKING_TAGS.includes(c[HEADER_MAP.CHECKLIST.REPORT_TAG]))
        .forEach(c => { checkItems[c[HEADER_MAP.CHECKLIST.REPORT_TAG]] = c; });

      const resultMap = {};
      TRACKING_TAGS.forEach(tag => {
        const ci = checkItems[tag];
        if (!ci) return;
        resultMap[tag] = allResults.find(r =>
          r[HEADER_MAP.RESULTS.TASK_ID] === taskId &&
          r[HEADER_MAP.RESULTS.CHECK_ID] === ci[HEADER_MAP.CHECKLIST.ID]
        ) || null;
      });

      // 複製表格並替換標籤 (排除照片標籤，避免被文字取代)
      const newTable = body.insertTable(insertIdx++, templateTable.copy());
      newTable.replaceText("\\{\\{任務ID_追蹤\\}\\}", taskId || "");
      const taskName = task[HEADER_MAP.TASKS.NAME] || "";
      const nameColor = (statusKey === "等待回覆" || statusKey === "暫緩處理") ? "#FF0000" : null;
      replaceIndividualTagWithFormat(newTable, "\\{\\{任務名稱_追蹤\\}\\}", taskName, nameColor);
      newTable.replaceText("\\{\\{任務分類_追蹤\\}\\}", task[HEADER_MAP.TASKS.CATEGORY] || "");
      newTable.replaceText("\\{\\{負責人_追蹤\\}\\}", task[HEADER_MAP.TASKS.ASSIGNEE] || "");
      newTable.replaceText("\\{\\{派工日期_追蹤\\}\\}", task[HEADER_MAP.TASKS.ASSIGN_DATE] || "");
      newTable.replaceText("\\{\\{截止日期_追蹤\\}\\}", task[HEADER_MAP.TASKS.DUE_DATE] || "");
      newTable.replaceText("\\{\\{執行進度_追蹤\\}\\}", statusKey);

      TRACKING_TAGS.forEach(tag => {
        if (tag === "佐證照片") return; // 照片由下方專門邏輯處理
        const r = resultMap[tag];
        const val = r ? String(r[HEADER_MAP.RESULTS.ACTUAL_VAL] || "").trim() : "";

        const tagPattern = `\\{\\{${tag}_追蹤\\}\\}`;
        const found = newTable.findText(tagPattern);

        if (found) {
          if (val === "" || val === "－") {
            // 若無資料，刪除該標籤所在的整列
            const row = found.getElement().getParent().getParent().getParent();
            if (row.getType() === DocumentApp.ElementType.TABLE_ROW) {
              row.removeFromParent();
            }
          } else {
            newTable.replaceText(tagPattern, val);
          }
        }
      });

      // 處理照片插入 (彙整所有追蹤欄位：議題描述、處理過程、處理結果、佐證照片的照片)
      const allPhotoIds = [];
      TRACKING_TAGS.forEach(tag => {
        const r = resultMap[tag];
        if (r && r[HEADER_MAP.RESULTS.PHOTO_ID]) {
          const ids = r[HEADER_MAP.RESULTS.PHOTO_ID].split(",").filter(Boolean);
          allPhotoIds.push(...ids);
        }
      });

      if (allPhotoIds.length > 0) {
        // 搜尋表格內的照片標籤 (支援 {{照片_追蹤}} 或 {{佐證照片_追蹤}})
        let photoRange = newTable.findText("\\{\\{照片_追蹤\\}\\}") || newTable.findText("\\{\\{佐證照片_追蹤\\}\\}");

        if (photoRange) {
          const para = photoRange.getElement().getParent().asParagraph();
          const tagText = photoRange.getElement().asText().getText();
          para.replaceText(tagText, "");
          // 單元格內照片使用較小比例 (寬度 180px)
          insertPhotosToPosition(para, allPhotoIds, 180);
        } else {
          // 若表格內無標籤，則在表格下方插入
          const photoPara = body.insertParagraph(insertIdx++, "");
          insertPhotosToPosition(photoPara, allPhotoIds);
        }
      } else {
        // 若無照片，刪除照片標籤所在的整列
        let photoRange = newTable.findText("\\{\\{照片_追蹤\\}\\}") || newTable.findText("\\{\\{佐證照片_追蹤\\}\\}");
        if (photoRange) {
          const row = photoRange.getElement().getParent().getParent().getParent();
          if (row.getType() === DocumentApp.ElementType.TABLE_ROW) {
            row.removeFromParent();
          }
        }
      }

      body.insertParagraph(insertIdx++, ""); // 間隔
    });
  });

  return true;
}

/**
 * 生成追蹤事項彙整報告 (輔助函式簡化版)
 */
function apiGenerateTrackingReport(payload) {
  const { projectName } = payload || {};
  const allTasks = DataService.getSheetData(CONFIG.SHEETS.TASKS);
  const allChecklists = DataService.getSheetData(CONFIG.SHEETS.CHECKLIST);
  const allResults = DataService.getSheetData(CONFIG.SHEETS.RESULTS);

  let trackingTasks = allTasks.filter(t =>
    (t[HEADER_MAP.TASKS.SUB_ITEM] === "追蹤事項" || t[HEADER_MAP.TASKS.SUB_ITEM] === "追蹤案件") &&
    (t[HEADER_MAP.TASKS.PROCESS_STATUS] || "").trim() !== ""
  );
  if (projectName) {
    trackingTasks = trackingTasks.filter(t => t[HEADER_MAP.TASKS.PROJECT_NAME] === projectName);
  }
  if (trackingTasks.length === 0) throw new Error("沒有符合條件的追蹤事項");

  const templateDocId = PropertiesService.getScriptProperties().getProperty("TRACKING_REPORT_TEMPLATE_ID");
  if (!templateDocId) throw new Error("尚未設定追蹤報告範本 ID。請至試算表選單中進行設定。");

  let templateFile;
  try {
    templateFile = DriveApp.getFileById(templateDocId);
  } catch (e) {
    throw new Error(`找不到指定的追蹤報告範本（ID: ${templateDocId}）。請確認該 Google Doc 存在，且您擁有編輯或讀取權限。`);
  }

  const tz = Session.getScriptTimeZone();
  const todayDisplay = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd");
  const reportFileName = `追蹤事項彙整_${projectName || "全部專案"}_${Utilities.formatDate(new Date(), tz, "yyyyMMdd")}`;

  // 取得儲存資料夾
  const reportFolder = getOrCreateDriveFolder("報告", "CONFIG_REPORT_FOLDER_ID");

  let newFile;
  try {
    newFile = templateFile.makeCopy(reportFileName, reportFolder);
  } catch (e) {
    throw new Error(`建立報告檔案複本失敗：${e.message}`);
  }

  const newDoc = DocumentApp.openById(newFile.getId());
  const body = newDoc.getBody();

  body.replaceText("\\{\\{生成日期\\}\\}", todayDisplay);
  body.replaceText("\\{\\{追蹤事項總數\\}\\}", String(trackingTasks.length));
  body.replaceText("\\{\\{專案\\}\\}", projectName || "全部專案");

  // 使用封裝好的渲染函式
  helperInsertTrackingSection(body, trackingTasks, allChecklists, allResults);

  newDoc.saveAndClose();
  return { success: true, reportUrl: newFile.getUrl() };
}


/**
 * 自動產生週期性任務
 * 可透過 Time-driven trigger 呼叫，例如每日凌晨
 * -----------------------------------------
 * 如何設定自動觸發器？
 * 進入您的 Google Apps Script 專案網頁。
 * 點擊左側時鐘圖示的 「觸發器 (Triggers)」。
 * 點擊右下角的 「+ 新增觸發器」。
 * 選擇函數：generatePeriodicTasks。
 * 選取活動來源：時間驅動。
 * 選取時間型觸發器類型：日計時器。
 * 選取時段：凌晨 0 時至 1 時。
 * 按「儲存」即可。
 */
function generatePeriodicTasks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const tasks = DataService.getSheetData(CONFIG.SHEETS.TYPES);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 找出所有設定為週期性任務的範本
    const templates = tasks.filter(t => {
      const pType = t[HEADER_MAP.TYPES.PERIODIC_TYPE];
      if (!pType) return false;

      const startStr = t[HEADER_MAP.TYPES.PERIODIC_START_DATE];
      const endStr = t[HEADER_MAP.TYPES.PERIODIC_END_DATE];
      let isValidDate = true;
      if (startStr) {
        isValidDate = isValidDate && (new Date(startStr) <= today);
      }
      if (endStr) {
        isValidDate = isValidDate && (new Date(endStr) >= today);
      }
      return isValidDate;
    });

    if (templates.length === 0) return;

    const tasksToCreate = [];
    templates.forEach(tpl => {
      const freq = tpl[HEADER_MAP.TYPES.PERIODIC_TYPE];
      let shouldCreate = false;
      if (freq === "每日") shouldCreate = true;
      else if (freq === "每週" && today.getDay() === 1) shouldCreate = true; // 每週一
      else if (freq === "每月" && today.getDate() === 1) shouldCreate = true; // 每月一日

      if (shouldCreate) {
        const newTask = { ...tpl };
        newTask[HEADER_MAP.TASKS.PROJECT_NAME] = tpl[HEADER_MAP.TYPES.PROJECT_NAME] || "";
        newTask[HEADER_MAP.TASKS.NAME] = todayStr;
        newTask[HEADER_MAP.TASKS.TEMPLATE_NAME] = tpl[HEADER_MAP.TYPES.NAME];
        newTask[HEADER_MAP.TASKS.PERIODIC_TYPE] = "";
        newTask[HEADER_MAP.TASKS.PERIODIC_START_DATE] = "";
        newTask[HEADER_MAP.TASKS.PERIODIC_END_DATE] = "";
        newTask[HEADER_MAP.TASKS.REPORT_STATUS] = "未開始";
        newTask[HEADER_MAP.TASKS.ASSIGN_DATE] = todayStr;
        newTask[HEADER_MAP.TASKS.DUE_DATE] = todayStr;
        tasksToCreate.push(newTask);
      }
    });

    if (tasksToCreate.length > 0) {
      apiBatchCreateTasks({ tasksArray: tasksToCreate });
    }
  } catch (e) {
    console.error("generatePeriodicTasks Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 從 Sheet 選單觸發產出報告
 */
function menuGenerateReportFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEETS.TASKS,
  );
  const range = sheet.getActiveRange();
  const row = range.getRow();

  if (row <= 1) {
    SpreadsheetApp.getUi().alert("請選擇數據列（非標題列）");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const taskIdIndex = headers.indexOf(HEADER_MAP.TASKS.ID);

  if (taskIdIndex === -1) {
    SpreadsheetApp.getUi().alert("找不到 " + HEADER_MAP.TASKS.ID + " 欄位");
    return;
  }

  const taskId = sheet.getRange(row, taskIdIndex + 1).getValue();
  if (!taskId) {
    SpreadsheetApp.getUi().alert("選中行無任務ID");
    return;
  }

  try {
    const result = apiGenerateReport({ taskId: taskId });
    if (result.success) {
      // 報告生成成功
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("產出失敗：" + e.message);
  }
}


/**
 * Config.gs
 * 系統全域設定與資料庫結構定義
 */

const CONFIG = {
  APP_TITLE: "鍋爐安裝工程",
  APP_SUBTITLE: "自主檢查表",
  CACHE_KEY: "bisms_appdata_v1",
  // 資料表名稱管理
  SHEETS: {
    TASKS: "數據_任務清單",
    RESULTS: "數據_檢驗結果",
    STAFF: "設定_人員名單",
    CHECKLIST: "設定_檢查項目",
    TYPES: "設定_樣板定義",
  },
  // 照片相關設定
  UPLOAD_IMAGE_MAX_SIZE: 1200,
  UPLOAD_IMAGE_QUALITY: 0.8,
  REPORT_IMAGE_MAX_WIDTH: 220,
  REPORT_IMAGE_MAX_HEIGHT: 300,
};

// 用於程式內部邏輯對應試算表中文標題
const HEADER_MAP = {
  TYPES: {
    CODE: "樣板代碼",
    NAME: "樣板名稱",
    REPORT_ID: "樣板報告ID",
    DEFAULT_TASK: "預設任務",
    PROJECT_NAME: "專案名稱",
    CATEGORY: "任務分類",
    SUB_ITEM: "任務分項",
    ASSIGNEE: "負責人",
    DURATION: "工期",
    PERIODIC_TYPE: "週期類型",
    PERIODIC_START_DATE: "週期起始日",
    PERIODIC_END_DATE: "週期結束日",
    MULTI_MODE: "多零件模式",
    MATERIAL_CODE: "料號",
  },
  CHECKLIST: {
    ID: "檢查ID",
    TEMPLATE_CODE: "樣板代碼",
    REPORT_TAG: "報告標籤",
    ITEM_NAME: "檢查項目",
    CRITERIA: "判定標準",
    LOGIC_TYPE: "邏輯型態",
    MIN_VAL: "最小值",
    MAX_VAL: "最大值",
    UNIT_OPTS: "單位或選項",
    REQUIRE_PHOTO: "需附照片",
  },
  STAFF: {
    NAME: "人員姓名",
    EMP_NO: "員工編號",
    ROLE: "權限角色",
    STATUS: "啟動狀態",
  },
  TASKS: {
    ID: "任務ID",
    PROJECT_NAME: "專案名稱",
    CATEGORY: "任務分類",
    SUB_ITEM: "任務分項",
    NAME: "任務名稱",
    TEMPLATE_CODE: "樣板代碼",
    TEMPLATE_NAME: "樣板名稱",
    ASSIGNEE: "負責人",
    ASSIGN_DATE: "派工日期",
    DUE_DATE: "截止日期",
    RELATED_TASK: "關聯工單",
    SUBMIT_DATE: "提報日期",
    REVIEW_DATE: "審查日期",
    CLOSE_DATE: "結案日期",
    PROCESS_STATUS: "執行進度",
    REPORT_STATUS: "報告狀態",
    REPORT_LINK: "報告連結",
    CLOSE_LINK: "結案連結",
    REMARK: "補充說明",
  },
  RESULTS: {
    ID: "記錄ID",
    TASK_ID: "任務ID",
    CHECK_ID: "檢查ID",
    COMPONENT_ID: "零件實例",
    ACTUAL_VAL: "實際數值",
    RESULT: "判定結果",
    REPORT_DATE: "回報日期",
    PHOTO_ID: "相片ID",
  },
};

const DB_SCHEMA = {};
DB_SCHEMA[CONFIG.SHEETS.TYPES] = [
  HEADER_MAP.TYPES.CODE,
  HEADER_MAP.TYPES.NAME,
  HEADER_MAP.TYPES.REPORT_ID,
  HEADER_MAP.TYPES.DEFAULT_TASK,
  HEADER_MAP.TYPES.PROJECT_NAME,
  HEADER_MAP.TYPES.CATEGORY,
  HEADER_MAP.TYPES.SUB_ITEM,
  HEADER_MAP.TYPES.ASSIGNEE,
  HEADER_MAP.TYPES.DURATION,
  HEADER_MAP.TYPES.PERIODIC_TYPE,
  HEADER_MAP.TYPES.PERIODIC_START_DATE,
  HEADER_MAP.TYPES.PERIODIC_END_DATE,
  HEADER_MAP.TYPES.MULTI_MODE,
  HEADER_MAP.TYPES.MATERIAL_CODE,
];
DB_SCHEMA[CONFIG.SHEETS.CHECKLIST] = [
  HEADER_MAP.CHECKLIST.ID,
  HEADER_MAP.CHECKLIST.TEMPLATE_CODE,
  HEADER_MAP.CHECKLIST.REPORT_TAG,
  HEADER_MAP.CHECKLIST.ITEM_NAME,
  HEADER_MAP.CHECKLIST.CRITERIA,
  HEADER_MAP.CHECKLIST.LOGIC_TYPE,
  HEADER_MAP.CHECKLIST.MIN_VAL,
  HEADER_MAP.CHECKLIST.MAX_VAL,
  HEADER_MAP.CHECKLIST.UNIT_OPTS,
  HEADER_MAP.CHECKLIST.REQUIRE_PHOTO,
];
DB_SCHEMA[CONFIG.SHEETS.STAFF] = [
  HEADER_MAP.STAFF.NAME,
  HEADER_MAP.STAFF.EMP_NO,
  HEADER_MAP.STAFF.ROLE,
  HEADER_MAP.STAFF.STATUS,
];
DB_SCHEMA[CONFIG.SHEETS.TASKS] = [
  HEADER_MAP.TASKS.ID,
  HEADER_MAP.TASKS.PROJECT_NAME,
  HEADER_MAP.TASKS.CATEGORY,
  HEADER_MAP.TASKS.SUB_ITEM,
  HEADER_MAP.TASKS.NAME,
  HEADER_MAP.TASKS.TEMPLATE_CODE,
  HEADER_MAP.TASKS.TEMPLATE_NAME,
  HEADER_MAP.TASKS.ASSIGNEE,
  HEADER_MAP.TASKS.ASSIGN_DATE,
  HEADER_MAP.TASKS.DUE_DATE,
  HEADER_MAP.TASKS.RELATED_TASK,
  HEADER_MAP.TASKS.SUBMIT_DATE,
  HEADER_MAP.TASKS.PROCESS_STATUS,
  HEADER_MAP.TASKS.REVIEW_DATE,
  HEADER_MAP.TASKS.CLOSE_DATE,
  HEADER_MAP.TASKS.REPORT_STATUS,
  HEADER_MAP.TASKS.REPORT_LINK,
  HEADER_MAP.TASKS.CLOSE_LINK,
  HEADER_MAP.TASKS.REMARK,
];
DB_SCHEMA[CONFIG.SHEETS.RESULTS] = [
  HEADER_MAP.RESULTS.ID,
  HEADER_MAP.RESULTS.TASK_ID,
  HEADER_MAP.RESULTS.CHECK_ID,
  HEADER_MAP.RESULTS.COMPONENT_ID,
  HEADER_MAP.RESULTS.ACTUAL_VAL,
  HEADER_MAP.RESULTS.RESULT,
  HEADER_MAP.RESULTS.REPORT_DATE,
  HEADER_MAP.RESULTS.PHOTO_ID,
];
/**
 * AdminSetup.gs
 * Google Sheets 自訂選單與管理工具
 * 提供初始化資料表等系統維護功能。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🏭 系統管理工具")
    .addItem("生成指定列的報告", "menuGenerateReportFromSheet")
    .addItem("批次生成報告 (未產出者)", "batchGeneratePendingReports")
    .addSeparator()
    .addItem("生成追蹤事項彙整報告", "menuGenerateTrackingReport")
    .addItem("設定追蹤報告範本 ID", "menuSetTrackingTemplateId")
    .addItem("設定報告儲存資料夾 ID", "menuSetReportFolderId")
    .addSeparator()
    .addItem("初始化資料庫 (建立缺少的表單)", "initDatabase")
    .addItem("生成測試用範例資料 (會寫入各表單底部)", "generateMockData")
    .addSeparator()
    .addItem("清除系統快取 (強制全體使用者更新)", "clearSystemCache")
    .addToUi();
}

/**
 * 選單：設定追蹤報告範本 ID
 */
function menuSetTrackingTemplateId() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty("TRACKING_REPORT_TEMPLATE_ID") || "（尚未設定）";
  const response = ui.prompt(
    "設定追蹤報告範本 ID",
    `請輸入 Google Docs 追蹤報告範本的 Doc ID。\n目前值：${current}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const newId = response.getResponseText().trim();
    if (newId) {
      PropertiesService.getScriptProperties().setProperty("TRACKING_REPORT_TEMPLATE_ID", newId);
      ui.alert("✅ 設定成功", `追蹤報告範本 ID 已更新為：\n${newId}`, ui.ButtonSet.OK);
    } else {
      ui.alert("⚠️ 未輸入任何值，設定未變更。");
    }
  }
}

/**
 * 選單：設定報告儲存資料夾 ID
 */
function menuSetReportFolderId() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty("CONFIG_REPORT_FOLDER_ID") || "（尚未設定，預設為本試算表相同資料夾）";
  const response = ui.prompt(
    "設定報告儲存資料夾 ID",
    `請輸入 Google Drive 報告儲存資料夾的 ID。\n目前值：${current}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const newId = response.getResponseText().trim();
    if (newId) {
      PropertiesService.getScriptProperties().setProperty("CONFIG_REPORT_FOLDER_ID", newId);
      ui.alert("✅ 設定成功", `報告儲存資料夾 ID 已更新為：\n${newId}`, ui.ButtonSet.OK);
    } else {
      ui.alert("⚠️ 未輸入任何值，設定未變更。");
    }
  }
}

/**
 * 選單：生成追蹤事項彙整報告
 */
function menuGenerateTrackingReport() {
  const ui = SpreadsheetApp.getUi();
  const filterResp = ui.prompt(
    "生成追蹤事項彙整報告",
    "請輸入要篩選的「專案名稱」（留空 = 全部專案）：",
    ui.ButtonSet.OK_CANCEL
  );
  if (filterResp.getSelectedButton() !== ui.Button.OK) return;

  const projectName = filterResp.getResponseText().trim() || null;
  try {
    const result = apiGenerateTrackingReport({ projectName });
    ui.alert(
      "✅ 報告生成成功",
      `追蹤事項彙整報告已生成！\n\n報告連結：\n${result.reportUrl}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("❌ 生成失敗", e.message, ui.ButtonSet.OK);
  }
}

/**
 * 根據 DB_SCHEMA 初始化建立試算表、防呆驗證機制與表頭設計。
 */
function initDatabase() {
  const ui = SpreadsheetApp.getUi();

  // --- [權限觸發門] ---
  // 透過呼叫各服務的簡單唯讀函數，強制 GAS 彈出授權視窗要求 Drive, Docs, Sheets 權限
  try {
    DriveApp.getStorageUsed(); // 觸發 Google Drive 權限
    DocumentApp.getFontFamilies(); // 觸發 Google Docs 權限
  } catch (e) {
    console.log("授權初始化提示: " + e.toString());
  }
  // -------------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let createdCount = 0;
  let updatedCount = 0;

  for (const sheetName in DB_SCHEMA) {
    let sheet = ss.getSheetByName(sheetName);
    const headers = DB_SCHEMA[sheetName];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      createdCount++;
    } else {
      updatedCount++;
    }

    // 寫入/更新表頭 (固定將第一列設為資料表的屬性標題)
    const currentCols = sheet.getLastColumn() || 1;
    const targetCols = Math.max(headers.length, currentCols);

    // 清理舊的殘餘表頭格式 (若現有欄數超過 schema 規劃)
    if (currentCols > headers.length) {
      sheet
        .getRange(1, headers.length + 1, 1, currentCols - headers.length)
        .clear();
    }

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f3f3");
    headerRange.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "black",
      SpreadsheetApp.BorderStyle.SOLID,
    );

    // 固定第一列為凍結標題列
    sheet.setFrozenRows(1);

  }

  var msg =
    "✅ 系統初始化與表單驗證更新完成！\n\n" +
    "• 新增表單：" +
    createdCount +
    " 個\n" +
    "• 檢查與更新：" +
    updatedCount +
    " 個\n\n" +
    "※ 系統已自動於「啟動狀態」、「需附照片」欄位插入核取方塊，並對「狀態」、「結果」欄位套用清單驗證。";
  ui.alert("資料庫建置報告", msg, ui.ButtonSet.OK);
}

/**
 * 實作 generateMockData 函數
 * 將 Mock.gs 中的範例資料批次寫入對應的資料表
 */
function generateMockData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("生成測試資料", "是否要在各表單末尾加入測試範例資料？", ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // 對應 Mock.gs 中的 key 與 CONFIG.SHEETS 中的名稱
  const mapping = {
    tasks: CONFIG.SHEETS.TASKS,
    results: CONFIG.SHEETS.RESULTS,
    staff: CONFIG.SHEETS.STAFF,
    checklists: CONFIG.SHEETS.CHECKLIST,
    templates: CONFIG.SHEETS.TYPES
  };

  let totalCount = 0;
  try {
    for (const key in mapping) {
      const sheetName = mapping[key];
      const dataArray = res[key]; // res 來自 Mock.gs

      if (dataArray && dataArray.length > 0) {
        dataArray.forEach(record => {
          DataService.appendRecord(sheetName, record);
          totalCount++;
        });
      }
    }
    ui.alert("成功", `測試資料已寫入完成，共計 ${totalCount} 筆。`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("錯誤", "寫入測試資料時發生失敗：" + e.message, ui.ButtonSet.OK);
  }
}


/**
 * 清除全域快取 (強制重新抓取資料表最新資料)
 */
function clearSystemCache() {
  const sheetsToClear = Object.values(CONFIG.SHEETS);

  // 1. 清除 Script Cache 中的各表資料
  sheetsToClear.forEach(sheetName => {
    DataService.clearCache(sheetName);
  });

  // 2. 清除舊有的全域快取 Key (如有殘留)
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
}

/**
 * IdService.gs
 * 統一 ID 生成服務，支援單項與批次 ID 序號產生
 */
const IdService = {
  /**
   * 取得指定工作表與 prefix 開頭的最新 ID 序號資訊
   * 回傳 { YYMM, maxNum }
   */
  getLatestSequence: function (sheetName, prefix) {
    const YYMM = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMM");
    const fullPrefix = prefix + YYMM;
    const allRecords = DataService.getSheetData(sheetName);

    // 尋找主鍵名稱
    let pkKey = "";
    for (const [sName, map] of Object.entries(CONFIG.SHEETS)) {
      if (map === sheetName) {
        pkKey = HEADER_MAP[sName].ID;
        break;
      }
    }
    if (!pkKey) throw new Error("找不到工作表對應的主鍵定義：" + sheetName);

    let maxNum = 0;
    for (let i = 0; i < allRecords.length; i++) {
      const currentId = allRecords[i][pkKey];
      if (currentId && typeof currentId === "string" && currentId.startsWith(fullPrefix)) {
        const numPart = currentId.substring(fullPrefix.length);
        const seq = parseInt(numPart, 10);
        if (!isNaN(seq) && seq > maxNum) {
          maxNum = seq;
        }
      }
    }
    return { YYMM, maxNum };
  },

  /**
   * 取得指定工作表與 prefix 開頭的下一個可用 ID (例如 T202605001)
   */
  nextId: function (sheetName, prefix) {
    const { YYMM, maxNum } = this.getLatestSequence(sheetName, prefix);
    const nextSeq = maxNum + 1;
    return prefix + YYMM + ("000" + nextSeq).slice(-3);
  }
};

/**
 * DataService.gs
 * 核心資料庫服務
 * 透過 Google Sheets 第一列表頭作為 MetaData，進行資料的序列化與反序列化
 * 將工作表資料並轉為物件陣列
 */
const DataService = {
  _sheetCache: {},
  _rawCache: {},

  /**
   * 取得指定名稱的工作表 (底層統一調用)
   */
  getSheet: function (sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`工作表 [${sheetName}] 不存在`);
    return sheet;
  },

  /**
   * 取得快取或即時資料 (Script Cache 層)
   */
  getCachedOrLive: function (sheetName) {
    const cache = CacheService.getScriptCache();
    const CACHE_PREFIX = CONFIG.CACHE_KEY + "_" + sheetName;
    const chunkCountStr = cache.get(CACHE_PREFIX + "_chunks");
    let cachedData = null;

    if (chunkCountStr) {
      const chunkCount = parseInt(chunkCountStr, 10);
      if (chunkCount > 0) {
        const keys = [];
        for (let i = 0; i < chunkCount; i++) {
          keys.push(CACHE_PREFIX + "_chunk_" + i);
        }
        try {
          const chunkMap = cache.getAll(keys);
          let completeString = "";
          let success = true;
          for (let i = 0; i < chunkCount; i++) {
            const chunk = chunkMap[CACHE_PREFIX + "_chunk_" + i];
            if (chunk === undefined || chunk === null) {
              success = false;
              break;
            }
            completeString += chunk;
          }
          if (success) {
            cachedData = completeString;
          }
        } catch (e) {
          console.error("Failed to read chunked cache for " + sheetName, e);
        }
      }
    } else {
      cachedData = cache.get(CACHE_PREFIX);
    }

    if (cachedData) {
      try {
        return JSON.parse(cachedData);
      } catch (e) {
        console.error("Cache parse error", e);
      }
    }

    const liveData = this.getSheetData(sheetName);
    try {
      const jsonStr = JSON.stringify(liveData);
      const CHUNK_SIZE = 85000; // 85KB, safely under the 100KB (102,400 bytes) limit
      if (jsonStr.length > CHUNK_SIZE) {
        const chunks = {};
        let index = 0;
        for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
          chunks[CACHE_PREFIX + "_chunk_" + index] = jsonStr.substring(i, i + CHUNK_SIZE);
          index++;
        }
        chunks[CACHE_PREFIX + "_chunks"] = String(index);
        cache.putAll(chunks, 600); // 10 minutes (600 seconds)
        cache.remove(CACHE_PREFIX);
      } else {
        cache.put(CACHE_PREFIX, jsonStr, 600);
        cache.remove(CACHE_PREFIX + "_chunks");
      }
    } catch (e) {
      console.warn("無法寫入 Cache: " + sheetName, e);
    }
    return liveData;
  },

  /**
   * 清除特定工作表的快取 (包含記憶體與 Script Cache)
   */
  clearCache: function (sheetName) {
    this._sheetCache[sheetName] = null;
    this._rawCache[sheetName] = null;
    const cache = CacheService.getScriptCache();
    const CACHE_PREFIX = CONFIG.CACHE_KEY + "_" + sheetName;

    const chunkCountStr = cache.get(CACHE_PREFIX + "_chunks");
    if (chunkCountStr) {
      const chunkCount = parseInt(chunkCountStr, 10);
      if (chunkCount > 0) {
        const keysToRemove = [CACHE_PREFIX + "_chunks"];
        for (let i = 0; i < chunkCount; i++) {
          keysToRemove.push(CACHE_PREFIX + "_chunk_" + i);
        }
        cache.removeAll(keysToRemove);
      }
    }
    cache.remove(CACHE_PREFIX);
  },

  /**
   * 取得資料：排除 A 欄為空的行，並處理日期格式 (包含 Request-level Cache)
   */
  getSheetData: function (sheetName) {
    if (this._sheetCache[sheetName]) {
      return this._sheetCache[sheetName];
    }

    const sheet = this.getSheet(sheetName);

    let rawData = this._rawCache[sheetName];
    if (!rawData) {
      rawData = sheet.getDataRange().getValues();
      this._rawCache[sheetName] = rawData;
    }

    if (rawData.length <= 1) {
      this._sheetCache[sheetName] = [];
      return [];
    }

    const headers = rawData[0];
    const tz = Session.getScriptTimeZone();

    let formulas = null;
    // 優化：只有 TASKS 表單需要讀取公式，且僅讀取「報告連結」那一欄
    if (sheetName === CONFIG.SHEETS.TASKS) {
      const linkColIdx = headers.indexOf(HEADER_MAP.TASKS.REPORT_LINK) + 1;
      if (linkColIdx > 0) {
        formulas = sheet.getRange(1, linkColIdx, sheet.getLastRow(), 1).getFormulas();
      }
    }

    const result = [];
    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (row[0] === "" || row[0] === null || String(row[0]).trim() === "") {
        continue;
      }

      const obj = {};
      headers.forEach((header, colIdx) => {
        let value = row[colIdx];
        // 優化：從單欄公式矩陣中取得公式
        const formula = (formulas && header === HEADER_MAP.TASKS.REPORT_LINK) ? formulas[i][0] : null;

        // 處理 HYPERLINK 公式，萃取出原始 URL 供前端使用
        if (formula && formula.toUpperCase().indexOf("=HYPERLINK") !== -1) {
          const match = formula.match(/=HYPERLINK\("([^"]+)"/i);
          if (match) {
            value = match[1];
          }
        }

        // 基本且優雅的日期處理
        if (value instanceof Date && !isNaN(value.getTime())) {
          value = Utilities.formatDate(value, tz, "yyyy/MM/dd");
        }

        obj[header] = value;
      });
      result.push(obj);
    }

    this._sheetCache[sheetName] = result;
    return result;
  },

  /**
   * 根據指定 Key 與 Value 取得一筆紀錄
   */
  getRecordByKey: function (sheetName, key, value) {
    return this.getSheetData(sheetName).find((r) => r[key] == value) || null;
  },

  /**
   * 新增單筆資料
   * 處理可能的 ISO 日期字串，確保寫入時是「日期格式」而非「純文字」
   */
  appendRecord: function (sheetName, recordObj) {
    const sheet = this.getSheet(sheetName);

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const newRow = headers.map((header) => {
      let val = recordObj[header] ?? "";

      // 簡單檢查：如果是前端傳來的 ISO 字串，轉回 Date 物件
      if (
        typeof val === "string" &&
        val.includes("T") &&
        !isNaN(Date.parse(val))
      ) {
        val = new Date(val);
      }
      return val;
    });

    sheet.appendRow(newRow);
    this.clearCache(sheetName); // 觸發全域快取更新
    return true;
  },

  /**
   * 更新或新增 (Upsert)
   */
  upsertRecord: function (sheetName, primaryKey, pkValue, recordObj) {
    const sheet = this.getSheet(sheetName);

    let data = this._rawCache[sheetName];
    if (!data) {
      data = sheet.getDataRange().getValues();
      this._rawCache[sheetName] = data;
    }
    const headers = data[0];
    const pkIndex = headers.indexOf(primaryKey);
    if (pkIndex === -1) return false;

    // 準備新資料列 (包含基本的日期字串轉物件)
    const newRow = headers.map((header) => {
      let val = recordObj[header] ?? "";
      if (
        typeof val === "string" &&
        val.includes("T") &&
        !isNaN(Date.parse(val))
      ) {
        val = new Date(val);
      }
      return val;
    });

    const rowIndex = data.findIndex(
      (row, idx) => idx > 0 && row[pkIndex] == pkValue,
    );

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    this.clearCache(sheetName); // 觸發全域快取更新
    return true;
  },

  /**
   * 局部更新欄位 (避免覆寫未更動的欄位或公式)
   */
  updateRecordFields: function (sheetName, primaryKey, pkValue, updatesObj) {
    const sheet = this.getSheet(sheetName);

    let data = this._rawCache[sheetName];
    if (!data) {
      data = sheet.getDataRange().getValues();
      this._rawCache[sheetName] = data;
    }
    const headers = data[0];
    const pkIndex = headers.indexOf(primaryKey);
    if (pkIndex === -1) return false;

    const rowIndex = data.findIndex(
      (row, idx) => idx > 0 && row[pkIndex] == pkValue,
    );
    if (rowIndex === -1) return false;

    const rowNum = rowIndex + 1;
    let hasChange = false;

    // 逐一優化更新特定欄位，避免讀取整列 setValues 造成 ARRAYFORMULA 被硬編碼(常數化)覆蓋
    for (const [key, value] of Object.entries(updatesObj)) {
      if (sheetName.indexOf("PM_") === 0 && (key === "階層" || key === "WBS")) {
        continue;
      }
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        let newVal = value;
        // 處理日期格式
        if (
          typeof newVal === "string" &&
          newVal.includes("T") &&
          !isNaN(Date.parse(newVal))
        ) {
          newVal = new Date(newVal);
        }

        const cell = sheet.getRange(rowNum, colIndex + 1);
        if (String(cell.getValue()) !== String(newVal)) {
          cell.setValue(newVal);
          hasChange = true;
        }
      }
    }

    if (hasChange) {
      this.clearCache(sheetName); // 觸發全域快取更新
    }
    return true;
  },
};

/**
 * [系統工具] 批次處理：優化版本，加入時間守衛避免 GAS 執行逾時
 */
function batchGeneratePendingReports() {
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 280000; // 4分40秒 (安全範圍)

  const ui = SpreadsheetApp.getUi();

  const tasks = DataService.getSheetData(CONFIG.SHEETS.TASKS);
  let processCount = 0;
  let successCount = 0;
  let errorCount = 0;
  let isTimeOut = false;

  for (let i = 0; i < tasks.length; i++) {
    // 檢查剩餘時間
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      isTimeOut = true;
      break;
    }

    const task = tasks[i];
    const taskId = task[HEADER_MAP.TASKS.ID];
    const reportStatus = task[HEADER_MAP.TASKS.REPORT_STATUS];
    const templateCode = task[HEADER_MAP.TASKS.TEMPLATE_CODE];

    // 篩選條件：有任務編號、有樣板編號
    // 且狀態符合 (已提報 或 已審查 或 已生成報告但無連結)
    const hasLink = task[HEADER_MAP.TASKS.REPORT_LINK] && String(task[HEADER_MAP.TASKS.REPORT_LINK]).trim() !== "";
    const isValidStatus =
      reportStatus === "已提報結果" || reportStatus === "已審查結果" || (reportStatus === "已生成報告" && !hasLink) || (reportStatus === "已結案" && !hasLink);

    if (taskId && templateCode && isValidStatus) {
      processCount++;
      try {
        const result = apiGenerateReport({ taskId: taskId });
        if (result.success) {
          successCount++;
        } else {
          errorCount++;
        }

        // 週期性寫入，確保進度存檔
        if (successCount % 3 === 0) {
          SpreadsheetApp.flush();
        }
      } catch (e) {
        console.error(`Batch process error for Task ${taskId}: ${e.message}`);
        errorCount++;
      }
    }
  }

  const resultMsg = isTimeOut
    ? `⚠️ 因執行時間限制，系統已自動暫停並結算進度。\n\n`
    : `🏁 全部處理完畢！\n\n`;

  ui.alert(
    isTimeOut ? "⏰ 執行逾時自動結算" : "🏁 批次執行成功",
    `${resultMsg}本輪檢查任務：${tasks.length}\n本輪處理：${processCount}\n成功產出：${successCount}\n處理失敗：${errorCount}${isTimeOut ? "\n\n提示：若還有剩餘項目，請再次執行。" : ""}`,
    ui.ButtonSet.OK,
  );
}

/**
 * 輔助函式：僅取代第一個匹配的標籤且不套用顏色 (一對一精密代入)
 */
function replaceFirstTagOnly(body, tag, value) {
  let range = body.findText(tag);
  if (range) {
    let element = range.getElement().asText();
    let start = range.getStartOffset();
    element.deleteText(start, range.getEndOffsetInclusive());
    element.insertText(start, value);
    return true;
  }
  return false;
}

/**
 * 輔助函式：僅取代第一個匹配的標籤並設定顏色格式 (一對一精密代入)
 */
function replaceFirstIndividualTagWithFormat(body, tag, value, color) {
  let range = body.findText(tag);
  if (range) {
    let element = range.getElement().asText();
    let start = range.getStartOffset();
    element.deleteText(start, range.getEndOffsetInclusive());
    element.insertText(start, value);
    if (color) {
      element.setForegroundColor(start, start + value.length - 1, color);
    }
    return true;
  }
  return false;
}

/**
 * 輔助函式：取代單一標籤並設定格式 (顏色)
 */
function replaceIndividualTagWithFormat(body, tag, value, color) {
  let range = body.findText(tag);
  while (range) {
    let element = range.getElement().asText();
    let start = range.getStartOffset();
    element.deleteText(start, range.getEndOffsetInclusive());
    element.insertText(start, value);
    if (color) {
      element.setForegroundColor(start, start + value.length - 1, color);
    }
    range = body.findText(tag, range);
  }
}

/**
 * 輔助函式：在指定段落位置插入照片 (模式 A)
 */
/**
 * 輔助函式：追加照片集到指定段落 (支援自定義寬度)
 */
function insertPhotosToPosition(para, photoIds, customWidth) {
  const targetW = customWidth || 220; // 預設寬度

  if (photoIds.length === 1) {
    try {
      const blob = DriveApp.getFileById(photoIds[0].trim()).getBlob();
      const img = para.appendInlineImage(blob);
      scaleImage(img, targetW, targetW * 1.5);
    } catch (e) {
      para.appendText(" [圖片載入失敗]");
    }
  } else {
    const container = para.getParent();
    const index = container.getChildIndex(para);
    const table = (container.getType() === DocumentApp.ElementType.TABLE_CELL)
      ? container.appendTable().setBorderWidth(0)
      : container.insertTable(index + 1).setBorderWidth(0);

    const numRows = Math.ceil(photoIds.length / 2);
    for (let r = 0; r < numRows; r++) {
      const row = table.appendTableRow();
      for (let c = 0; c < 2; c++) {
        const cell = row.appendTableCell();
        const pIdx = r * 2 + c;
        if (pIdx < photoIds.length) {
          try {
            const blob = DriveApp.getFileById(photoIds[pIdx].trim()).getBlob();
            const cellPara = cell
              .getChild(0)
              .asParagraph()
              .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            const img = cellPara.appendInlineImage(blob);
            scaleImage(img, targetW, targetW * 1.5);
          } catch (e) {
            cell.setText("圖片插入失敗");
          }
        }
      }
    }
  }
}

/**
 * 輔助函式：擴充定錨表格 (模式 B)
 */
function expandAnchorTable(body, checklists, results, photosBuffered) {
  let table = null;
  let templateRow = null;
  let templateRowIndex = -1;

  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const tbl = tables[t];
    for (let r = 0; r < tbl.getNumRows(); r++) {
      const row = tbl.getRow(r);
      if (row.getText().includes("{{序號_檢查清單}}")) {
        table = tbl;
        templateRow = row;
        templateRowIndex = r;
        break;
      }
    }
    if (table) break;
  }

  if (table && templateRow) {
    checklists.forEach((item, index) => {
      const result = results.find(
        (r) => r[HEADER_MAP.RESULTS.CHECK_ID] === item[HEADER_MAP.CHECKLIST.ID],
      );
      let actualResult = result
        ? result[HEADER_MAP.RESULTS.RESULT] || ""
        : "－";
      let actualValue = result
        ? result[HEADER_MAP.RESULTS.ACTUAL_VAL] || ""
        : "";

      // 邏輯型態為空時，使用 單位或選項 作為預設實際值
      if (actualValue === "" && !item[HEADER_MAP.CHECKLIST.LOGIC_TYPE]) {
        actualValue = item[HEADER_MAP.CHECKLIST.UNIT_OPTS] || "";
      }

      if (
        item[HEADER_MAP.CHECKLIST.LOGIC_TYPE] === "數值" &&
        actualValue !== "" &&
        item[HEADER_MAP.CHECKLIST.UNIT_OPTS]
      ) {
        actualValue =
          actualValue + " " + (item[HEADER_MAP.CHECKLIST.UNIT_OPTS] || "");
      }

      // 當實際數值為空時，填入 "－" 
      if (!actualValue) actualValue = "未填寫";

      const pBuffer = photosBuffered.find(
        (pb) =>
          pb.checkItem[HEADER_MAP.CHECKLIST.ID] ===
          item[HEADER_MAP.CHECKLIST.ID],
      );
      const photoRef = pBuffer ? pBuffer.photoReference : "－";

      const newRow = table.insertTableRow(templateRowIndex + 1 + index);
      for (let c = 0; c < templateRow.getNumCells(); c++) {
        newRow.appendTableCell(templateRow.getCell(c).copy());
      }

      newRow.replaceText("{{序號_檢查清單}}", (index + 1).toString());
      newRow.replaceText(
        "{{項目_檢查清單}}",
        item[HEADER_MAP.CHECKLIST.ITEM_NAME] || "",
      );
      newRow.replaceText(
        "{{標準_檢查清單}}",
        item[HEADER_MAP.CHECKLIST.CRITERIA] || "",
      );
      newRow.replaceText("{{數值_檢查清單}}", actualValue);
      newRow.replaceText("{{結果_檢查清單}}", actualResult);
      newRow.replaceText("{{照片_檢查清單}}", photoRef);

      if (actualResult === "不合格" || actualValue === "未填寫") {
        for (let c = 0; c < newRow.getNumCells(); c++) {
          const cell = newRow.getCell(c);
          const targetText = actualValue === "未填寫" ? "未填寫" : actualResult;
          const range = cell.findText(targetText);
          if (range) {
            range
              .getElement()
              .asText()
              .setForegroundColor(
                range.getStartOffset(),
                range.getEndOffsetInclusive(),
                "#FF0000",
              );
          }
        }
      }
    });
    table.removeRow(templateRowIndex);
  }
}

/**
 * 輔助函式：追加照片集到指定定位點 (模式 B)
 */
function appendPhotosToAggregateSection(body, photosToAppend) {
  const tag = "{{照片集_檢查清單}}";
  const range = body.findText(tag);
  if (range && photosToAppend.length > 0) {
    const paragraph = range.getElement().getParent().asParagraph();
    const container = paragraph.getParent();
    let currentIdx = container.getChildIndex(paragraph);
    container.removeChild(paragraph); //移除原始的{{照片集_檢查清單}}
    container.insertPageBreak(currentIdx++);//插入分頁符號
    const title = container.insertParagraph(currentIdx++, "佐證照片");//插入標題
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);//設定標題格式

    photosToAppend.forEach((p) => {
      const heading = container.insertParagraph(
        currentIdx++,
        `檢查項目 #${p.index}: ${p.checkItem[HEADER_MAP.CHECKLIST.ITEM_NAME]}`,
      );//插入小標題
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);//設定小標題格式

      const photoTbl = container.insertTable(currentIdx++);//插入表格
      photoTbl.setBorderWidth(0);//設定表格邊框

      const nRows = Math.ceil(p.photoIds.length / 2);//計算表格行數
      for (let r = 0; r < nRows; r++) {
        const row = photoTbl.appendTableRow();//插入表格列
        for (let c = 0; c < 2; c++) {
          const cell = row.appendTableCell();
          const pIdx = r * 2 + c;
          if (pIdx < p.photoIds.length) {
            try {
              const blob = DriveApp.getFileById(
                p.photoIds[pIdx].trim(),
              ).getBlob();
              const cellPara = cell
                .getChild(0)
                .asParagraph()
                .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              const img = cellPara.appendInlineImage(blob);
              scaleImage(img, 220, 300);
              cell
                .appendParagraph(`#${p.index}-${pIdx + 1}`)
                .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
                .setItalic(true)
                .setFontSize(8);
            } catch (e) {
              cell.setText("圖片插入失敗");
            }
          }
        }
      }
      //container.insertParagraph(currentIdx++, ""); //插入空白行
    });
  } else if (range) {
    body.replaceText(tag, "");//本次檢查無照片存檔。
  }
}

/**
 * 輔助函式：縮放圖片
 */
function scaleImage(img, maxW, maxH) {
  const mW = maxW || Number(CONFIG.REPORT_IMAGE_MAX_WIDTH) || 400;
  const mH = maxH || Number(CONFIG.REPORT_IMAGE_MAX_HEIGHT) || 600;
  const w = img.getWidth();
  const h = img.getHeight();
  const ratio = Math.min(mW / w, mH / h, 1.0);
  img.setWidth(w * ratio).setHeight(h * ratio);
}
