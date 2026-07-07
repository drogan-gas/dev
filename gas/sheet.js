/**
 * sheet.gs
 * 🏭 鍋爐設備安裝工程自動化管理系統 - google sheet 自動化工具
 * Google Apps Script V8 Engine
 */

/**
 * 規格化 WBS 碼，處理 Google Sheets 浮點數、空格，確保精準關聯
 */
function normalizeWbs(wbs) {
  if (wbs === null || wbs === undefined) return "";
  return String(wbs).split(".")
    .map(p => p.trim())
    .filter(p => p !== "")
    .join(".");
}

/**
 * onEdit 統一網關與批次範圍解析器
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const startRow = e.range.getRow();
  const endRow = e.range.getLastRow();
  const col = e.range.getColumn();
  const endCol = e.range.getLastColumn();

  // 忽略標題列與空白工作表編輯
  if (startRow <= 1) return;

  // 使用 CacheService 防止同一範圍的遞迴或併發觸發
  const cache = CacheService.getScriptCache();
  const lockKey = "onedit_active_" + sheetName + "_" + startRow + "_" + endRow;
  if (cache.get(lockKey)) {
    return;
  }
  cache.put(lockKey, "1", 10);

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    // 1. 設定_檢查項目連動
    if (sheetName === CONFIG.SHEETS.CHECKLIST) {
      ChecklistService.handleEdit(sheet, startRow, endRow, col, endCol, headers, e);
    }

    // 2. 數據_任務清單連動
    if (sheetName === CONFIG.SHEETS.TASKS) {
      TaskSheetService.handleEdit(sheet, startRow, endRow, col, endCol, headers, e);
    }

    // 3. PM_ 專案管理多維度派工與渲染連動
    if (sheetName.indexOf("PM_") === 0) {
      PmDispatchService.handleEdit(sheet, startRow, endRow, col, endCol, headers, e);
    }
  } finally {
    cache.remove(lockKey);
  }
}

/**
 * ========================================================
 * 1. Checklist 檢查樣板自動化服務 (ChecklistService)
 * ========================================================
 */
const ChecklistService = {
  handleEdit: function (sheet, startRow, endRow, col, endCol, headers, e) {
    const reportTagCol = headers.indexOf(HEADER_MAP.CHECKLIST.REPORT_TAG) + 1;
    const idCol = headers.indexOf(HEADER_MAP.CHECKLIST.ID) + 1;
    if (reportTagCol <= 0 || idCol <= 0) return;

    // 確保編輯區與報告標籤欄重疊
    if (col <= reportTagCol && endCol >= reportTagCol) {
      let isChanged = false;
      const rangeValues = sheet.getRange(startRow, reportTagCol, (endRow - startRow + 1), 1).getValues();

      for (let r = startRow; r <= endRow; r++) {
        const val = rangeValues[r - startRow][0];
        if (val && String(val).trim() !== "") {
          const idCell = sheet.getRange(r, idCol);
          if (!idCell.getValue()) {
            idCell.setValue(IdService.nextId(CONFIG.SHEETS.CHECKLIST, "C"));
            isChanged = true;
          }
        }
      }
      if (isChanged) {
        DataService.clearCache(CONFIG.SHEETS.CHECKLIST);
      }
    }
  }
};

/**
 * ========================================================
 * 2. Task 數據任務清單自動化服務 (TaskService)
 * ========================================================
 */
const TaskSheetService = {
  handleEdit: function (sheet, startRow, endRow, col, endCol, headers, e) {
    const tplCodeCol = headers.indexOf(HEADER_MAP.TASKS.TEMPLATE_CODE) + 1;
    if (tplCodeCol <= 0) return;

    if (col <= tplCodeCol && endCol >= tplCodeCol) {
      const templates = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
      let isChanged = false;

      const editVals = sheet.getRange(startRow, tplCodeCol, (endRow - startRow + 1), 1).getValues();

      for (let r = startRow; r <= endRow; r++) {
        const val = editVals[r - startRow][0];
        const matchedTpl = templates.find(t => t[HEADER_MAP.TYPES.CODE] === val);
        if (matchedTpl) {
          const mappings = [
            { taskKey: HEADER_MAP.TASKS.TEMPLATE_NAME, typeKey: HEADER_MAP.TYPES.NAME },
            { taskKey: HEADER_MAP.TASKS.CATEGORY, typeKey: HEADER_MAP.TYPES.CATEGORY },
            { taskKey: HEADER_MAP.TASKS.SUB_ITEM, typeKey: HEADER_MAP.TYPES.SUB_ITEM },
            { taskKey: HEADER_MAP.TASKS.ASSIGNEE, typeKey: HEADER_MAP.TYPES.ASSIGNEE }
          ];

          mappings.forEach(m => {
            const colIndex = headers.indexOf(m.taskKey) + 1;
            if (colIndex > 0) {
              const cell = sheet.getRange(r, colIndex);
              if (!cell.getValue() && matchedTpl[m.typeKey]) {
                cell.setValue(matchedTpl[m.typeKey]);
                isChanged = true;
              }
            }
          });
        }
      }
      if (isChanged) {
        DataService.clearCache(CONFIG.SHEETS.TASKS);
      }
    }
  }
};

/**
 * ========================================================
 * 3. PM 工作表派工核心邏輯服務 (PmDispatchService)
 * ========================================================
 */
const PmDispatchService = {
  handleEdit: function (sheet, startRow, endRow, col, endCol, headers, e) {
    const levelColIdx = headers.indexOf("階層") + 1;
    const taskNameColIdx = headers.indexOf("任務名稱") + 1;
    const dueDateColIdx = headers.indexOf("截止日期") + 1;
    const taskIdColIdx = headers.indexOf("任務ID") + 1;
    const dispatchConfirmColIdx = headers.indexOf("派工確認") + 1;
    const assigneeColIdx = headers.indexOf("負責人") + 1;

    // A. 處理 L2/L3 底層無子項目自動展開料號 L4 與 自檢表單欄位自動對應
    const checkFormColIdx = headers.indexOf("自檢表單") + 1;
    const isEditTarget = (col <= taskNameColIdx && endCol >= taskNameColIdx) ||
      (checkFormColIdx > 0 && col <= checkFormColIdx && endCol >= checkFormColIdx);
    if (isEditTarget && levelColIdx > 0) {
      const types = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
      const wbsColIdx = headers.indexOf("WBS") + 1;
      let allWbsList = [];
      if (wbsColIdx > 0) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          allWbsList = sheet.getRange(1, wbsColIdx, lastRow, 1).getValues().map(row => normalizeWbs(row[0]));
        }
      }

      // 批次讀取編輯範圍內所有單元格
      const totalEditRows = endRow - startRow + 1;
      const editBlockRange = sheet.getRange(startRow, 1, totalEditRows, headers.length);
      const editBlockValues = editBlockRange.getValues();

      // 批次讀取自檢表單
      let checkFormUpdatedValues = null;
      if (checkFormColIdx > 0) {
        checkFormUpdatedValues = sheet.getRange(startRow, checkFormColIdx, totalEditRows, 1).getValues();
      }

      let isCheckFormChanged = false;

      for (let r = endRow; r >= startRow; r--) {
        const localIdx = r - startRow;
        const levelVal = parseInt(editBlockValues[localIdx][levelColIdx - 1], 10);
        const wbsVal = wbsColIdx > 0 ? String(editBlockValues[localIdx][wbsColIdx - 1]).trim() : "";
        const normWbs = normalizeWbs(wbsVal);
        let hasChild = false;
        if (normWbs && allWbsList.length > 0) {
          const childPrefix = normWbs + ".";
          hasChild = allWbsList.some((otherWbs, idx) => idx > 0 && otherWbs.indexOf(childPrefix) === 0);
        }

        if (levelVal >= 2 && !hasChild) {
          const taskNameVal = String(editBlockValues[localIdx][taskNameColIdx - 1]).trim();
          let matchedType = null;
          if (taskNameVal) {
            matchedType = types.find(t => String(t[HEADER_MAP.TYPES.NAME]).trim() === taskNameVal);
          }

          if (matchedType) {
            const templateName = String(matchedType[HEADER_MAP.TYPES.NAME]).trim();
            if (checkFormColIdx > 0) {
              checkFormUpdatedValues[localIdx][0] = templateName;
              isCheckFormChanged = true;
            }
            const materialCodeVal = matchedType[HEADER_MAP.TYPES.MATERIAL_CODE];
            if (materialCodeVal && String(materialCodeVal).trim() !== "") {
              this.expandLevel4Materials(sheet, r, templateName, headers);
            }
          } else {
            let checkFormVal = "";
            if (checkFormColIdx > 0) {
              checkFormVal = String(checkFormUpdatedValues[localIdx][0]).trim();
            } else {
              checkFormVal = taskNameVal;
            }
            const templateName = checkFormVal;
            if (templateName) {
              const matchedTypeByForm = types.find(t => String(t[HEADER_MAP.TYPES.NAME]).trim() === templateName);
              if (matchedTypeByForm) {
                const materialCodeVal = matchedTypeByForm[HEADER_MAP.TYPES.MATERIAL_CODE];
                if (materialCodeVal && String(materialCodeVal).trim() !== "") {
                  this.expandLevel4Materials(sheet, r, templateName, headers);
                }
              }
            }
          }
        }
      }

      // 整批回寫自檢表單
      if (isCheckFormChanged && checkFormColIdx > 0 && checkFormUpdatedValues) {
        sheet.getRange(startRow, checkFormColIdx, totalEditRows, 1).setValues(checkFormUpdatedValues);
      }
    }

    // B. 處理截止日期雙向同步回 數據_任務清單
    if (dueDateColIdx > 0 && col <= dueDateColIdx && endCol >= dueDateColIdx && taskIdColIdx > 0) {
      for (let r = startRow; r <= endRow; r++) {
        const taskIdVal = String(sheet.getRange(r, taskIdColIdx).getValue()).trim();
        if (taskIdVal && taskIdVal !== "" && taskIdVal !== "null" && taskIdVal !== "undefined" && taskIdVal !== "等待派工") {
          const rawDueDate = sheet.getRange(r, dueDateColIdx).getValue();
          let formattedDueDate = "";
          if (rawDueDate instanceof Date) {
            formattedDueDate = Utilities.formatDate(rawDueDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
          } else if (rawDueDate && String(rawDueDate).trim() !== "") {
            formattedDueDate = String(rawDueDate).trim();
          }

          try {
            DataService.updateRecordFields(
              CONFIG.SHEETS.TASKS,
              HEADER_MAP.TASKS.ID,
              taskIdVal,
              { [HEADER_MAP.TASKS.DUE_DATE]: formattedDueDate }
            );
            DataService.clearCache(CONFIG.SHEETS.TASKS);
            showToastMessage(`已同步更新任務 ${taskIdVal} 的截止日期為 [${formattedDueDate}]。`);
          } catch (err) {
            console.error("同步截止日期失敗:", err);
          }
        }
      }
    }

    // B.2 處理負責人雙向同步回 數據_任務清單
    if (assigneeColIdx > 0 && col <= assigneeColIdx && endCol >= assigneeColIdx && taskIdColIdx > 0) {
      for (let r = startRow; r <= endRow; r++) {
        const taskIdVal = String(sheet.getRange(r, taskIdColIdx).getValue()).trim();
        if (taskIdVal && taskIdVal !== "" && taskIdVal !== "null" && taskIdVal !== "undefined" && taskIdVal !== "等待派工") {
          const assigneeVal = String(sheet.getRange(r, assigneeColIdx).getValue()).trim();
          try {
            DataService.updateRecordFields(
              CONFIG.SHEETS.TASKS,
              HEADER_MAP.TASKS.ID,
              taskIdVal,
              { [HEADER_MAP.TASKS.ASSIGNEE]: assigneeVal }
            );
            DataService.clearCache(CONFIG.SHEETS.TASKS);
            showToastMessage(`已同步更新任務 ${taskIdVal} 的負責人為 [${assigneeVal}]。`);
          } catch (err) {
            console.error("同步負責人失敗:", err);
          }
        }
      }
    }

    // B.3 處理任務名稱雙向同步回 數據_任務清單
    if (taskNameColIdx > 0 && col <= taskNameColIdx && endCol >= taskNameColIdx && taskIdColIdx > 0) {
      for (let r = startRow; r <= endRow; r++) {
        const taskIdVal = String(sheet.getRange(r, taskIdColIdx).getValue()).trim();
        if (taskIdVal && taskIdVal !== "" && taskIdVal !== "null" && taskIdVal !== "undefined" && taskIdVal !== "等待派工") {
          const taskNameVal = String(sheet.getRange(r, taskNameColIdx).getValue()).trim();
          try {
            DataService.updateRecordFields(
              CONFIG.SHEETS.TASKS,
              HEADER_MAP.TASKS.ID,
              taskIdVal,
              { [HEADER_MAP.TASKS.NAME]: taskNameVal }
            );
            DataService.clearCache(CONFIG.SHEETS.TASKS);
            showToastMessage(`已同步更新任務 ${taskIdVal} 的任務名稱為 [${taskNameVal}]。`);
          } catch (err) {
            console.error("同步任務名稱失敗:", err);
          }
        }
      }
    }

    // B.4 處理自檢表單雙向同步回 數據_任務清單
    if (checkFormColIdx > 0 && col <= checkFormColIdx && endCol >= checkFormColIdx && taskIdColIdx > 0) {
      for (let r = startRow; r <= endRow; r++) {
        const taskIdVal = String(sheet.getRange(r, taskIdColIdx).getValue()).trim();
        if (taskIdVal && taskIdVal !== "" && taskIdVal !== "null" && taskIdVal !== "undefined" && taskIdVal !== "等待派工") {
          const checkFormVal = String(sheet.getRange(r, checkFormColIdx).getValue()).trim();
          let templateCode = "";
          try {
            const types = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
            const tNames = checkFormVal.split(',').map(n => n.trim()).filter(Boolean);
            const tCodes = [];
            tNames.forEach(tName => {
              const matchedType = types.find(t => String(t[HEADER_MAP.TYPES.NAME]).trim() === tName);
              if (matchedType) {
                tCodes.push(matchedType[HEADER_MAP.TYPES.CODE] || "");
              }
            });
            templateCode = tCodes.join(',');
          } catch (err) {
            console.error("同步解析樣板代碼失敗:", err);
          }
          try {
            DataService.updateRecordFields(
              CONFIG.SHEETS.TASKS,
              HEADER_MAP.TASKS.ID,
              taskIdVal,
              {
                [HEADER_MAP.TASKS.TEMPLATE_NAME]: checkFormVal,
                [HEADER_MAP.TASKS.TEMPLATE_CODE]: templateCode
              }
            );
            DataService.clearCache(CONFIG.SHEETS.TASKS);
            showToastMessage(`已同步更新任務 ${taskIdVal} 的自檢表單（樣板）為 [${checkFormVal}]。`);
          } catch (err) {
            console.error("同步自檢表單失敗:", err);
          }
        }
      }
    }

    // C. 處理派工確認 checkbox 點擊
    let isDispatchActive = false;
    if (dispatchConfirmColIdx > 0 && col <= dispatchConfirmColIdx && endCol >= dispatchConfirmColIdx) {
      for (let r = startRow; r <= endRow; r++) {
        const val = sheet.getRange(r, dispatchConfirmColIdx).getValue();
        if (val === true || val === "TRUE") {
          isDispatchActive = true;
          this.dispatchSingleTask(sheet, r, headers);
        }
      }
    }

    // D. 批量渲染顏色樣式 (若非勾選派工時段，則做防護與即時著色更新)
    if (!isDispatchActive) {
      PmStyleRenderer.updateStyles(sheet, startRow, endRow);
    }
  },

  /**
   * 料號下級自動展開 (Level 3 or Level 4)
   */
  expandLevel4Materials: function (sheet, row, templateName, headers) {
    try {
      const types = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
      const matchedType = types.find(t => String(t[HEADER_MAP.TYPES.NAME]).trim() === templateName);
      if (!matchedType) {
        const dispatchConfirmColIdx = headers.indexOf("派工確認") + 1;
        if (dispatchConfirmColIdx > 0) {
          sheet.getRange(row, dispatchConfirmColIdx).setBackground("#FCE8E6");
        }
        showAlertMessage(`找不到對應的樣板名稱 [${templateName}]，請先在「設定_樣板定義」中新增對應的 自主檢查表。`);
        return;
      }

      const materialCodeVal = matchedType[HEADER_MAP.TYPES.MATERIAL_CODE];
      if (!materialCodeVal || String(materialCodeVal).trim() === "") return;

      const materialCodes = String(materialCodeVal).split(",").map(s => s.trim()).filter(s => s !== "");
      if (materialCodes.length === 0) return;

      const levelColIdx = headers.indexOf("階層") + 1;
      const wbsColIdx = headers.indexOf("WBS") + 1;
      const taskNameColIdx = headers.indexOf("任務名稱") + 1;
      const checkFormColIdx = headers.indexOf("自檢表單") + 1;

      let parentLevel = 3;
      if (levelColIdx > 0) {
        const pLevelVal = parseInt(sheet.getRange(row, levelColIdx).getValue(), 10);
        if (!isNaN(pLevelVal)) {
          parentLevel = pLevelVal;
        }
      }

      let parentWbs = "";
      if (wbsColIdx > 0) {
        parentWbs = String(sheet.getRange(row, wbsColIdx).getValue()).trim();
      }

      const childLevel = parentLevel + 1;
      const indentSpaces = " ".repeat(parentLevel * 2);

      const levelNameMap = { 2: "二", 3: "三", 4: "四", 5: "五" };
      const childLevelName = levelNameMap[childLevel] || String(childLevel);

      let proceed = true;
      try {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          "❓ 自動展開確認",
          `選取的樣板 [${templateName}] 含有 ${materialCodes.length} 筆料號（${materialCodes.join(", ")}），是否要依序在下方自動新增第${childLevelName}階層任務？`,
          ui.ButtonSet.YES_NO
        );
        proceed = (response === ui.Button.YES);
      } catch (e) {
        proceed = true; // 無 UI 環境預設允許
      }

      if (proceed) {
        if (taskNameColIdx > 0) {
          const lock = LockService.getScriptLock();
          if (lock.tryLock(10000)) {
            try {
              sheet.insertRowsAfter(row, materialCodes.length);
              for (let i = 0; i < materialCodes.length; i++) {
                const insertedRow = row + 1 + i;
                sheet.getRange(insertedRow, taskNameColIdx).setValue(indentSpaces + materialCodes[i]);
                if (checkFormColIdx > 0) {
                  sheet.getRange(insertedRow, checkFormColIdx).setValue(templateName);
                }
              }
              // 新增後重新全距樣式更新
              PmStyleRenderer.updateStyles(sheet, row, row + materialCodes.length);
              showToastMessage(`已自動依 [${templateName}] 的子料號展開建立 ${materialCodes.length} 項${childLevelName}階工程。`);
            } finally {
              lock.releaseLock();
            }
          }
        }
      }
    } catch (err) {
      console.error("層級料號自動展開錯誤:", err);
    }
  },

  /**
   * 單一底層工作任務實體指派派工 (原子事務防鎖死)
   */
  dispatchSingleTask: function (sheet, row, headers) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(15000)) {
      try {
        const levelColIdx = headers.indexOf("階層") + 1;
        const wbsColIdx = headers.indexOf("WBS") + 1;
        const taskNameColIdx = headers.indexOf("任務名稱") + 1;
        const taskIdColIdx = headers.indexOf("任務ID") + 1;
        const assigneeColIdx = headers.indexOf("負責人") + 1;
        const dispatchConfirmColIdx = headers.indexOf("派工確認") + 1;
        const dueDateColIdx = headers.indexOf("截止日期") + 1;

        if (levelColIdx === 0 || wbsColIdx === 0 || taskNameColIdx === 0 || taskIdColIdx === 0 || assigneeColIdx === 0 || dispatchConfirmColIdx === 0 || dueDateColIdx === 0) {
          console.error("PM_ 表頭缺少關鍵欄位！");
          return;
        }

        const rowRange = sheet.getRange(row, 1, 1, headers.length);
        const rowVals = rowRange.getValues()[0];
        const currentTaskId = String(rowVals[taskIdColIdx - 1]).trim();
        const levelVal = parseInt(rowVals[levelColIdx - 1], 10);
        const wbsVal = String(rowVals[wbsColIdx - 1]).trim();
        let assigneeVal = String(rowVals[assigneeColIdx - 1]).trim();

        // 已有真實 Task ID 則略過
        if (currentTaskId && currentTaskId !== "" && currentTaskId !== "null" && currentTaskId !== "undefined" && currentTaskId !== "等待派工") {
          return;
        }

        // 階層防禦 (支援 L2 或 L3 底層無子項目之直接衍生派工)
        if (isNaN(levelVal) || levelVal < 2) {
          sheet.getRange(row, dispatchConfirmColIdx).setValue('');
          showToastMessage("只有階層大於等於 2 的任務才可以進行派工！");
          return;
        }

        const normWbs = normalizeWbs(wbsVal);
        const parts = normWbs.split(".");
        if (parts.length < 2) {
          sheet.getRange(row, dispatchConfirmColIdx).setValue('');
          showToastMessage("WBS 編碼無效或層級不足！");
          return;
        }

        const lastRow = sheet.getLastRow();
        const allWbsRange = sheet.getRange(1, wbsColIdx, lastRow, 1).getValues();
        const allWbsList = allWbsRange.map(r => normalizeWbs(r[0]));

        // 偵測是否有下層子節點
        const childPrefix = normWbs + ".";
        const hasChild = allWbsList.some((otherWbs, idx) => idx > 0 && otherWbs.indexOf(childPrefix) === 0);
        if (hasChild) {
          sheet.getRange(row, dispatchConfirmColIdx).setValue('');
          showToastMessage("此項目含有下層細分任務，請直接對其底下子任務進行派工！");
          return;
        }

        // 負責人自動承襲自上級節點
        if (!assigneeVal || assigneeVal === "undefined" || assigneeVal === "null" || assigneeVal === "") {
          for (let len = parts.length - 1; len >= 1; len--) {
            const parentWbs = parts.slice(0, len).join(".");
            const foundIdx = allWbsList.indexOf(normalizeWbs(parentWbs));
            if (foundIdx !== -1) {
              const allAssigneesRange = sheet.getRange(1, assigneeColIdx, lastRow, 1).getValues();
              const pAssignee = String(allAssigneesRange[foundIdx][0]).trim();
              if (pAssignee && pAssignee !== "undefined" && pAssignee !== "null" && pAssignee !== "") {
                sheet.getRange(row, assigneeColIdx).setValue(pAssignee);
                assigneeVal = pAssignee;
                break;
              }
            }
          }
        }

        if (!assigneeVal || assigneeVal === "undefined" || assigneeVal === "null" || assigneeVal === "") {
          sheet.getRange(row, dispatchConfirmColIdx).setValue(false);
          showAlertMessage("負責人未填，無法指派派工任務！");
          return;
        }

        // 日期轉化
        let dueDateVal = "";
        const rawDueDate = rowVals[dueDateColIdx - 1];
        if (rawDueDate instanceof Date) {
          dueDateVal = Utilities.formatDate(rawDueDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else if (rawDueDate && String(rawDueDate).trim() !== "") {
          dueDateVal = String(rawDueDate).trim();
        }

        // 解析各階名稱項目以形成工單任務
        const allTaskNamesRange = sheet.getRange(1, taskNameColIdx, lastRow, 1).getValues();
        const getTaskNameByWbs = (targetWbs) => {
          const foundIdx = allWbsList.indexOf(targetWbs);
          if (foundIdx !== -1) {
            return String(allTaskNamesRange[foundIdx][0]).trim();
          }
          return "";
        };

        const checkFormColIdx = headers.indexOf("自檢表單") + 1;

        // 1. 任務分類：對應其所屬的 L1 任務名稱
        const categoryName = getTaskNameByWbs(parts[0]);

        // 2. 任務分項：對應其所屬的 L2 任務名稱，如果自己本身就是 L2 則 為空，若本身L3以上，則 結合 L2~L3 作為任務分項
        let subItemName = "";
        if (levelVal === 2) {
          subItemName = "";
        } else if (levelVal >= 3) {
          const l2Name = getTaskNameByWbs(parts[0] + "." + parts[1]);
          const l3Name = getTaskNameByWbs(parts[0] + "." + parts[1] + "." + parts[2]);
          if (l2Name && l3Name) {
            subItemName = (l2Name === l3Name) ? l2Name : `${l2Name} - ${l3Name}`;
          } else {
            subItemName = l2Name || l3Name || "";
          }
        }

        // 3. 任務名稱對應其對應其所屬的任務名稱
        let newTaskName = String(rowVals[taskNameColIdx - 1]).trim();

        // 4. 樣板名稱：會以 父項目 與 自己 的自檢表單 作為 新增任務的 樣板名稱，採用","分隔
        const formsSet = new Set();

        let selfForm = "";
        if (checkFormColIdx > 0) {
          selfForm = String(rowVals[checkFormColIdx - 1]).trim();
        }

        if (parts.length >= 2) {
          for (let len = parts.length - 1; len >= 1; len--) {
            const parentWbs = parts.slice(0, len).join(".");
            const foundIdx = allWbsList.indexOf(normalizeWbs(parentWbs));
            if (foundIdx !== -1 && checkFormColIdx > 0) {
              const pCheckForm = String(sheet.getRange(foundIdx + 1, checkFormColIdx).getValue()).trim();
              if (pCheckForm && pCheckForm !== "undefined" && pCheckForm !== "null" && pCheckForm !== "" && pCheckForm !== "缺自檢表單") {
                pCheckForm.split(',').map(s => s.trim()).filter(Boolean).forEach(f => formsSet.add(f));
              }
            }
          }
        }

        if (selfForm && selfForm !== "undefined" && selfForm !== "null" && selfForm !== "" && selfForm !== "缺自檢表單") {
          selfForm.split(',').map(s => s.trim()).filter(Boolean).forEach(f => formsSet.add(f));
        }

        let finalForms = Array.from(formsSet);
        if (finalForms.length === 0) {
          finalForms = ["安裝進度"];
        }
        let templateName = finalForms.join(',');

        if (!templateName) {
          sheet.getRange(row, dispatchConfirmColIdx).setValue(false);
          sheet.getRange(row, dispatchConfirmColIdx).setBackground("#FCE8E6");
          showAlertMessage(`找不到對應的樣板項目 [WBS: ${wbsVal}]`);
          return;
        }

        let templateCode = "";
        try {
          const types = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
          const tNames = templateName.split(',').map(n => n.trim()).filter(Boolean);
          const tCodes = [];
          tNames.forEach(tName => {
            const matchedType = types.find(t => String(t[HEADER_MAP.TYPES.NAME]).trim() === tName);
            if (matchedType) {
              tCodes.push(matchedType[HEADER_MAP.TYPES.CODE] || "");
            }
          });
          templateCode = tCodes.join(',');
        } catch (err) {
          console.error("解析樣板代碼失敗:", err);
        }

        if (!templateCode) {
          sheet.getRange(row, dispatchConfirmColIdx).setValue(false);
          sheet.getRange(row, dispatchConfirmColIdx).setBackground("#FCE8E6");
          showAlertMessage(`「設定_樣板定義」中查無配對到名稱為 [${templateName}] 的樣板對應代碼。`);
          return;
        }

        const sheetName = sheet.getName();
        const projectName = sheetName.replace("PM_", "");
        const newTaskId = IdService.nextId(CONFIG.SHEETS.TASKS, "T");
        const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

        const newTaskObj = {
          [HEADER_MAP.TASKS.ID]: newTaskId,
          [HEADER_MAP.TASKS.PROJECT_NAME]: projectName,
          [HEADER_MAP.TASKS.CATEGORY]: categoryName,
          [HEADER_MAP.TASKS.SUB_ITEM]: subItemName,
          [HEADER_MAP.TASKS.NAME]: newTaskName,
          [HEADER_MAP.TASKS.TEMPLATE_CODE]: templateCode,
          [HEADER_MAP.TASKS.TEMPLATE_NAME]: templateName,
          [HEADER_MAP.TASKS.ASSIGNEE]: assigneeVal,
          [HEADER_MAP.TASKS.ASSIGN_DATE]: todayStr,
          [HEADER_MAP.TASKS.DUE_DATE]: dueDateVal,
          [HEADER_MAP.TASKS.PROCESS_STATUS]: "",
          [HEADER_MAP.TASKS.REPORT_STATUS]: "未開始"
        };

        // 新增工單至工單資料表
        DataService.appendRecord(CONFIG.SHEETS.TASKS, newTaskObj);

        // 回填 ID 於 PM 主表
        sheet.getRange(row, taskIdColIdx).setValue(newTaskId);
        if (checkFormColIdx > 0) {
          sheet.getRange(row, checkFormColIdx).setValue(templateName);
        }

        // 即時美化
        PmStyleRenderer.updateStyles(sheet, row, row);
        DataService.clearCache(CONFIG.SHEETS.TASKS);
      } finally {
        lock.releaseLock();
      }
    }
  }
};

/**
 * ========================================================
 * 4. PM 工作表高性能渲染引擎 (PmStyleRenderer)
 * ========================================================
 */
const PmStyleRenderer = {
  /**
   * 批次更新特定行範圍或整網頁 PM 表格色彩與 Cbx
   * 徹底避免單個 range get/set 的 Rtt 浪費 (O(N) 內存計算)
   */
  updateStyles: function (sheet, startRow, endRow) {
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;

      const renderStart = 2;
      const renderEnd = lastRow;

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const levelColIdx = headers.indexOf("階層");
      const wbsColIdx = headers.indexOf("WBS");
      const taskIdColIdx = headers.indexOf("任務ID");
      const assigneeColIdx = headers.indexOf("負責人");
      const dispatchConfirmColIdx = headers.indexOf("派工確認");
      const taskNameColIdx = headers.indexOf("任務名稱");
      const checkFormColIdx = headers.indexOf("自檢表單");

      if (levelColIdx === -1 || wbsColIdx === -1 || taskIdColIdx === -1 || assigneeColIdx === -1 || dispatchConfirmColIdx === -1 || taskNameColIdx === -1) {
        return;
      }

      // 一次性抓取整張表格 WBS 與核心內容作為關聯判定快取
      const allRange = sheet.getRange(1, 1, lastRow, headers.length);
      const allVals = allRange.getValues();
      const allWbsList = allVals.map(r => normalizeWbs(r[wbsColIdx]));

      const types = DataService.getCachedOrLive(CONFIG.SHEETS.TYPES);
      const validTemplateNamesSet = new Set(types.map(t => String(t[HEADER_MAP.TYPES.NAME]).trim()));

      const updateRowsCount = renderEnd - renderStart + 1;
      const targetRange = sheet.getRange(renderStart, 1, updateRowsCount, headers.length);
      const bgColors = targetRange.getBackgrounds();

      // 量化 Cbx validations
      const validationRange = sheet.getRange(renderStart, dispatchConfirmColIdx + 1, updateRowsCount, 1);
      const validations = validationRange.getDataValidations();
      const cellValues = validationRange.getValues();

      // 批次讀取自檢表單，防範行內單元格單獨讀取與寫入
      let checkFormRange = null;
      let checkFormValues = null;
      if (checkFormColIdx !== -1) {
        checkFormRange = sheet.getRange(renderStart, checkFormColIdx + 1, updateRowsCount, 1);
        checkFormValues = checkFormRange.getValues();
      }

      const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

      for (let r = renderStart; r <= renderEnd; r++) {
        const arrIdx = r - renderStart;
        const rowVals = allVals[r - 1];

        const levelVal = parseInt(rowVals[levelColIdx], 10);
        const wbsVal = String(rowVals[wbsColIdx]).trim();
        const normWbs = normalizeWbs(wbsVal);
        const currentTaskId = String(rowVals[taskIdColIdx]).trim();
        const assigneeVal = String(rowVals[assigneeColIdx]).trim();

        // 偵測樣板名稱 (子項目若空則向上繼承母項目)
        let templateName = "";
        if (checkFormColIdx !== -1) {
          templateName = String(rowVals[checkFormColIdx]).trim();
        }
        if (!templateName || templateName === "undefined" || templateName === "null") {
          templateName = "";
        }

        if (!templateName && normWbs) {
          const parts = normWbs.split(".");
          for (let len = parts.length - 1; len >= 1; len--) {
            const parentWbs = parts.slice(0, len).join(".");
            const foundIdx = allWbsList.indexOf(normalizeWbs(parentWbs));
            if (foundIdx !== -1 && checkFormColIdx !== -1) {
              const pVal = String(allVals[foundIdx][checkFormColIdx]).trim();
              if (pVal && pVal !== "undefined" && pVal !== "null" && pVal !== "") {
                templateName = pVal;
                break;
              }
            }
          }
        }

        // 重設底色
        bgColors[arrIdx][assigneeColIdx] = null;
        bgColors[arrIdx][dispatchConfirmColIdx] = null;
        bgColors[arrIdx][taskIdColIdx] = null;
        if (checkFormColIdx !== -1) {
          bgColors[arrIdx][checkFormColIdx] = null;
        }

        // 子代偵測
        let hasChild = false;
        if (normWbs) {
          const childPrefix = normWbs + ".";
          hasChild = allWbsList.some((otherWbs, idx) => idx > 0 && otherWbs.indexOf(childPrefix) === 0);
        }

        const isEligible = !isNaN(levelVal) && levelVal >= 2 && normWbs !== "" && !hasChild;

        if (!isEligible) {
          // 非底層工單：移除 Cbx 狀態
          validations[arrIdx][0] = null;
          cellValues[arrIdx][0] = "";
          if ((levelVal === 2 || levelVal === 3) && (!assigneeVal || assigneeVal === "")) {
            bgColors[arrIdx][assigneeColIdx] = "#FCE8E6"; // 紅色警示 (負責人缺漏)
          }
        } else {
          // 待派工：檢索繼承負責人
          let effectiveAssignee = assigneeVal;
          if (!effectiveAssignee || effectiveAssignee === "undefined" || effectiveAssignee === "null") {
            effectiveAssignee = "";
          }
          if (!effectiveAssignee && normWbs) {
            const parts = normWbs.split(".");
            for (let len = parts.length - 1; len >= 1; len--) {
              const parentWbs = parts.slice(0, len).join(".");
              const foundIdx = allWbsList.indexOf(normalizeWbs(parentWbs));
              if (foundIdx !== -1) {
                const pAssign = String(allVals[foundIdx][assigneeColIdx]).trim();
                if (pAssign && pAssign !== "null" && pAssign !== "undefined" && pAssign !== "") {
                  effectiveAssignee = pAssign;
                  break;
                }
              }
            }
          }

          const templates = templateName ? templateName.split(",").map(t => t.trim()).filter(Boolean) : [];
          const hasTemplateError = templates.length === 0 || templates.some(t => !validTemplateNamesSet.has(t));
          const hasAssigneeError = !effectiveAssignee || effectiveAssignee === "";

          if (hasTemplateError) {
            // ⚠️ 找不到對應自主檢查表樣板別或完全空白
            validations[arrIdx][0] = null;
            if (!templateName) {
              cellValues[arrIdx][0] = "缺自檢表單";
            } else {
              cellValues[arrIdx][0] = "無此樣板";
            }
            bgColors[arrIdx][dispatchConfirmColIdx] = "#FCE8E6"; // 🔴 紅色
            if (checkFormColIdx !== -1) {
              bgColors[arrIdx][checkFormColIdx] = "#FCE8E6"; // 🔴 標紅
            }
          } else if (hasAssigneeError) {
            // ⚠️ 缺負責人
            validations[arrIdx][0] = null;
            cellValues[arrIdx][0] = "缺負責人";
            bgColors[arrIdx][dispatchConfirmColIdx] = "#FEF3C7"; // 🟡 黃色
            bgColors[arrIdx][assigneeColIdx] = "#FCE8E6";       // 🔴 負責人欄位標紅
          } else {
            // 正常：符合底層且資料就緒，應為 checkbox
            validations[arrIdx][0] = checkboxRule;

            // 已發配工單 (已有 ID)
            if (currentTaskId && currentTaskId !== "" && currentTaskId !== "null" && currentTaskId !== "undefined" && currentTaskId !== "等待派工") {
              bgColors[arrIdx][dispatchConfirmColIdx] = "#D4EDDA"; // 🟢 綠色派工
              bgColors[arrIdx][taskIdColIdx] = "#D4EDDA";
              cellValues[arrIdx][0] = true;
            } else {
              // 待派工就緒
              bgColors[arrIdx][dispatchConfirmColIdx] = "#E0F2FE"; // 🔵 可派工就緒
              if (cellValues[arrIdx][0] !== false && cellValues[arrIdx][0] !== true) {
                cellValues[arrIdx][0] = false;
              }
            }
          }
        }
      }

      // 整批回寫
      validationRange.setValues(cellValues);
      validationRange.setDataValidations(validations);
      if (checkFormColIdx !== -1 && checkFormRange && checkFormValues) {
        checkFormRange.setValues(checkFormValues);
      }
      targetRange.setBackgrounds(bgColors);
    } catch (err) {
      console.error("高性能渲染出錯:", err);
    }
  }
};

/**
 * ========================================================
 * 5. 全域通知與互動彈窗通用套件
 * ========================================================
 */
function showToastMessage(msg) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "系統提示", 5);
  } catch (e) {
    console.log(msg);
  }
}

function showAlertMessage(msg) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("⚠️ 系統提示", msg, ui.ButtonSet.OK);
  } catch (e) {
    showToastMessage(msg);
  }
}
