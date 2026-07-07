/**
 * Code.js - 收發文與進度追蹤管理系統後端入口點 (Google Apps Script - dispatch2)
 * 📂 位於 /dispatch2/Code.js
 * 
 * 採用標準 Finite State Machine (FSM) 架構與表頭防禦映射
 * 由「流程設定」工作表驅動所有狀態 (State) 與動作 (Action) 轉換
 */

const SCRIPT_VERSION = "v3.0.0-FSM";
const SHEET_MASTER = "收發文主表";
const SHEET_TRACKING = "進度追蹤表";
const SHEET_CONFIG = "系統設定";
const SHEET_FLOW_CONFIG = "流程設定";

// ==================== [表頭防禦映射 (HEADER MAPS)] ====================
// 將 Google Sheets 上對使用者極友善的「自訂說明標題」映射為程式碼內乾淨的「英文 camelCase 鍵」
const MAP_MASTER = {
    "docId": "公文案號",
    "date": "收發日期",
    "type": "收發分類",
    "vendor": "往來單位/對象",
    "tag": "分類標籤",
    "subject": "主旨說明",
    "assignee": "負責人員",
    "order": "累計處理次數",
    "lastAction": "最後執行動作",
    "status": "目前流轉狀態",
    "fileName": "最新附件名稱",
    "fileId": "最新附件雲端檔案ID",
    "signatureFileId": "電子簽章圖片ID",
    "relatedDocId": "關聯對應公文案號",
    "dueDate": "截止日期"
};

const MAP_TRACKING = {
    "docId": "公文案號",
    "order": "次序",
    "date": "異動日期",
    "action": "執行動作",
    "description": "處理說明與備註",
    "dueDate": "本次截止日期",
    "status": "異動後狀態",
    "fileName": "本次附件名稱",
    "fileId": "本次附件雲端檔案ID",
    "signatureFileId": "本次電子簽章圖片ID"
};

const MAP_CONFIG = {
    "vendor": "往來對象",
    "tag": "分類標籤",
    "prefix": "字號前綴",
    "assignee": "負責人員",
    "templateId": "預設範本文件ID"
};

const MAP_FLOW_CONFIG = {
    "type": "流程分類",
    "order": "顯示順序",
    "status": "目前狀態",
    "action": "觸發動作",
    "nextStatus": "後續狀態",
    "defaultDesc": "觸發動作的處理說明",
    "reqFile": "附件 (是/否)",
    "reqDueDate": "期限 (是/否/天數)",
    "reqSignature": "簽章 (是/否)",
    "phase": "流程階段"
};

// 計算真正的表頭陣列以利寫入表格
const MASTER_HEADERS = Object.keys(MAP_MASTER).map(function (k) { return MAP_MASTER[k]; });
const TRACKING_HEADERS = Object.keys(MAP_TRACKING).map(function (k) { return MAP_TRACKING[k]; });
const CONFIG_HEADERS = Object.keys(MAP_CONFIG).map(function (k) { return MAP_CONFIG[k]; });
const FLOW_CONFIG_HEADERS = Object.keys(MAP_FLOW_CONFIG).map(function (k) { return MAP_FLOW_CONFIG[k]; });

/**
 * 網頁進入點，渲染 Index.html
 */
function doGet(e) {
    try {
        DriveApp.getRootFolder();
    } catch (err) {
        console.warn("DriveApp initialization check: " + err);
    }

    const title = "📬 收發文管理系統 (FSM)";
    return HtmlService.createHtmlOutputFromFile("Index")
        .setTitle(title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no");
}

/**
 * HTTP POST 請求入口 (接收二進位流上傳)
 */
function doPost(e) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000); // 15 秒互斥鎖防併發

        let data = {};
        if (e.postData && e.postData.contents) {
            try {
                data = JSON.parse(e.postData.contents);
            } catch (ex) {
                // Ignore parsing errors
            }
        }

        const action = data.action || e.parameter.action;
        if (action === "uploadAttachment") {
            const docId = data.docId || e.parameter.docId;
            const fileName = data.fileName || e.parameter.fileName;
            const base64Data = data.base64Data;
            const oldFileId = data.oldFileId || e.parameter.oldFileId;

            if (!docId || !fileName || !base64Data) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false,
                    message: "參數不足"
                })).setMimeType(ContentService.MimeType.JSON);
            }

            const base64Clean = base64Data.split(",")[1] || base64Data;
            const decodedBytes = Utilities.base64Decode(base64Clean);
            const blob = Utilities.newBlob(decodedBytes, "application/pdf", fileName);
            const folder = getOrCreateAttachmentFolder();
            const file = folder.createFile(blob);

            try {
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (shareErr) {
                console.warn("無法將上傳的附件設定為公開共用:", shareErr.toString());
            }

            updateDocumentFile(docId, file.getId(), file.getName());

            if (oldFileId && oldFileId !== "uploading_temp_id") {
                try {
                    DriveApp.getFileById(oldFileId).setTrashed(true);
                } catch (trashErr) {
                    console.warn("無法將舊檔案移至垃圾桶 (File ID: " + oldFileId + "): " + trashErr.toString());
                }
            }

            SpreadsheetApp.flush();

            return ContentService.createTextOutput(JSON.stringify({
                success: true,
                fileId: file.getId(),
                fileName: file.getName(),
                message: "附件上傳成功"
            })).setMimeType(ContentService.MimeType.JSON);
        }

        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            message: "不支援的 POST Action: " + action
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        console.error("doPost 執行失敗:", err);
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            message: err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    } finally {
        lock.releaseLock();
    }
}

/**
 * API 統一入口網關 (透過 LockService 確保資料庫讀寫一致性)
 */
function api(action, payload) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000); // 獲取排隊鎖，最長等待 15 秒

        switch (action) {
            case "getInitialData":
                return getInitialData();
            case "createDocument":
                return createDocument(payload);
            case "addTrackingRecord":
                return addTrackingRecord(payload);
            case "initializeSheets":
                return initializeSheets();
            case "getNextDocId":
                return getNextDocId(payload);
            case "saveSystemConfig":
                return saveSystemConfig(payload);
            case "saveFlowConfig":
                return saveFlowConfig(payload);
            case "uploadAttachment":
                return uploadAttachment(payload);
            case "uploadSignature":
                return uploadSignature(payload);
            case "getWebAppUrl":
                return getWebAppUrl();
            case "copyTemplateFile":
                return copyTemplateFile(payload);
            case "deleteTemplateFile":
                return deleteTemplateFile(payload);
            default:
                throw new Error("未支援的 API 動作: " + action);
        }
    } catch (err) {
        console.error("API 執行失敗 (" + action + "):", err);
        return {
            success: false,
            message: err.toString()
        };
    } finally {
        lock.releaseLock(); // 100% 確保解除死鎖
    }
}

// ==================== [試算表與關聯防禦映射的核心讀寫函式] ====================

/**
 * 將指定試算表依照防禦表頭映射 (Map) 解析為英文屬性物件陣列
 */
function readSheetToEnglishObjects(sheet, map) {
    if (sheet.getLastRow() < 2) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // 建立 [真實中文字頭 -> 英文對應 key] 的快取對應
    const headerToKey = {};
    for (const key in map) {
        headerToKey[map[key]] = key;
    }

    return values.map(function (row, rowIndex) {
        const obj = { _rowIndex: rowIndex + 2 };
        headers.forEach(function (header, colIndex) {
            const trimmedHeader = String(header).trim();
            const key = headerToKey[trimmedHeader];
            if (key) {
                let val = row[colIndex];
                if (val instanceof Date) {
                    obj[key] = formatDate(val);
                } else {
                    obj[key] = val;
                }
            }
        });
        return obj;
    });
}

/**
 * 依照防禦表頭映射，將物件屬性值寫回特定行
 */
function writeEnglishObjectToRow(sheet, map, rowNum, obj) {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    headers.forEach(function (header, colIndex) {
        const trimmedHeader = String(header).trim();
        for (const key in map) {
            if (map[key] === trimmedHeader) {
                if (obj[key] !== undefined) {
                    sheet.getRange(rowNum, colIndex + 1).setValue(obj[key]);
                }
                break;
            }
        }
    });
}

/**
 * 依照防禦表頭映射，新增一行物件資料到試算表
 */
function appendEnglishObject(sheet, map, obj) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const rowValues = [];

    headers.forEach(function (header) {
        const trimmedHeader = String(header).trim();
        let val = "";
        for (const key in map) {
            if (map[key] === trimmedHeader) {
                val = obj[key] !== undefined ? obj[key] : "";
                break;
            }
        }
        rowValues.push(val);
    });

    sheet.appendRow(rowValues);
    return sheet.getLastRow();
}

/**
 * 確保系統設定工作表存在並寫入初始資料
 */
function ensureConfigSheet(ss) {
    let configSheet = ss.getSheetByName(SHEET_CONFIG);
    if (!configSheet) {
        configSheet = ss.insertSheet(SHEET_CONFIG);
        configSheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS])
            .setFontWeight("bold")
            .setBackground("#f1f5f9");

        const initialConfigs = [
            { vendor: "中鼎工程", tag: "一般公文", prefix: "D", assignee: "工地經理", templateId: "1lidKUODsPWE5PKJbzNmk8Rr4tNp3-l_DTtqBDEg-x3U" },
            { vendor: "勞動局", tag: "安全檢查", prefix: "S", assignee: "公安組長", templateId: "" },
            { vendor: "台電", tag: "採購招標", prefix: "P", assignee: "專案經理", templateId: "" },
            { vendor: "", tag: "備忘錄", prefix: "M", assignee: "專案經理", templateId: "" }
        ];

        initialConfigs.forEach(function (c) {
            appendEnglishObject(configSheet, MAP_CONFIG, c);
        });
        SpreadsheetApp.flush();
    }
    return configSheet;
}

/**
 * 初始化工作表與標準 FSM 流程設定
 */
function initializeSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. 初始化收發文主表
    let masterSheet = ss.getSheetByName(SHEET_MASTER);
    if (!masterSheet) {
        masterSheet = ss.insertSheet(SHEET_MASTER);
    }
    masterSheet.clear();
    masterSheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS])
        .setFontWeight("bold")
        .setBackground("#e2e8f0");

    // 2. 初始化進度追蹤表
    let trackingSheet = ss.getSheetByName(SHEET_TRACKING);
    if (!trackingSheet) {
        trackingSheet = ss.insertSheet(SHEET_TRACKING);
    }
    trackingSheet.clear();
    trackingSheet.getRange(1, 1, 1, TRACKING_HEADERS.length).setValues([TRACKING_HEADERS])
        .setFontWeight("bold")
        .setBackground("#e2e8f0");

    // 3. 初始化系統常規設定
    let configSheet = ss.getSheetByName(SHEET_CONFIG);
    if (configSheet) {
        ss.deleteSheet(configSheet);
    }
    ensureConfigSheet(ss);

    // 4. 初始化標準 FSM 流程設定
    let flowConfigSheet = ss.getSheetByName(SHEET_FLOW_CONFIG);
    if (flowConfigSheet) {
        ss.deleteSheet(flowConfigSheet);
    }
    flowConfigSheet = ss.insertSheet(SHEET_FLOW_CONFIG);
    flowConfigSheet.getRange(1, 1, 1, FLOW_CONFIG_HEADERS.length).setValues([FLOW_CONFIG_HEADERS])
        .setFontWeight("bold")
        .setBackground("#e2e8f0");

    // 標準 FSM 關卡流轉設定
    const flowRows = [
        // 收文 FSM 流轉 / nextStatus 的最後一步一定要含有 "結案"字眼，才會自動閉環
        { type: "收文", order: 1, phase: "1. 收件與登記", status: "待掃描", action: "傳閱來文", nextStatus: "傳閱中", defaultDesc: "文件掃描完成後上傳，開始傳閱確認處理方式", reqFile: "是", reqDueDate: "3", reqSignature: "否" },
        { type: "收文", order: 2, phase: "2. 傳閱與研判", status: "傳閱中", action: "存查結案", nextStatus: "已存查結案", defaultDesc: "無需回覆，直接存查歸檔", reqFile: "否", reqDueDate: "否", reqSignature: "否" },
        { type: "收文", order: 2, phase: "2. 傳閱與研判", status: "傳閱中", action: "啟動回文", nextStatus: "先存查結案", defaultDesc: "需撰寫回文，來文先存查歸檔", reqFile: "否", reqDueDate: "否", reqSignature: "否" },

        // 發文 FSM 流轉
        { type: "發文", order: 1, phase: "1. 草稿起草", status: "起草中", action: "送交審核", nextStatus: "審核中", defaultDesc: "起草完成後上傳，提送主管審查", reqFile: "是", reqDueDate: "否", reqSignature: "否" },
        { type: "發文", order: 2, phase: "2. 主管審核", status: "審核中", action: "核准送印", nextStatus: "待用印", defaultDesc: "審核通過，送交文管用印", reqFile: "否", reqDueDate: "否", reqSignature: "否" },
        { type: "發文", order: 2, phase: "2. 主管審核", status: "審核中", action: "退回重寫", nextStatus: "起草中", defaultDesc: "審核不通過，退回修改", reqFile: "否", reqDueDate: "否", reqSignature: "否" },
        { type: "發文", order: 3, phase: "3. 用印送件", status: "待用印", action: "用印送件", nextStatus: "已送件", defaultDesc: "蓋印完成經掃描後上傳，紙本送交對方", reqFile: "是", reqDueDate: "3", reqSignature: "否" },
        { type: "發文", order: 4, phase: "4. 外部簽收", status: "已送件", action: "簽收後結案", nextStatus: "已簽收結案", defaultDesc: "對方代表簽收完成，本案正式結案", reqFile: "否", reqDueDate: "否", reqSignature: "是" },
        { type: "發文", order: 4, phase: "4. 外部簽收", status: "已送件", action: "簽收後等回文", nextStatus: "待對方回文", defaultDesc: "對方代表簽收完成，本案需等對方回文", reqFile: "否", reqDueDate: "7", reqSignature: "是" },
        { type: "發文", order: 5, phase: "5. 回文歸檔", status: "待對方回文", action: "回文歸檔", nextStatus: "已歸檔結案", defaultDesc: "收到回文掃描完成後上傳，本案正式結案", reqFile: "是", reqDueDate: "否", reqSignature: "否" }
    ];

    flowRows.forEach(function (row) {
        appendEnglishObject(flowConfigSheet, MAP_FLOW_CONFIG, row);
    });

    SpreadsheetApp.flush();

    return {
        success: true,
        message: "所有 BISMS 工作表已完全重新初始化為全新 FSM 與防禦表頭架構。"
    };
}

/**
 * 取得完整初始資料 (統一為英文屬性傳遞至前端)
 */
function getInitialData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let masterSheet = ss.getSheetByName(SHEET_MASTER);
    let trackingSheet = ss.getSheetByName(SHEET_TRACKING);
    let flowConfigSheet = ss.getSheetByName(SHEET_FLOW_CONFIG);

    if (!masterSheet || !trackingSheet || !flowConfigSheet) {
        initializeSheets();
        masterSheet = ss.getSheetByName(SHEET_MASTER);
        trackingSheet = ss.getSheetByName(SHEET_TRACKING);
        flowConfigSheet = ss.getSheetByName(SHEET_FLOW_CONFIG);
    }

    const configSheet = ensureConfigSheet(ss);

    const masterData = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
    const trackingData = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
    const configData = readSheetToEnglishObjects(configSheet, MAP_CONFIG);
    const flowConfigData = readSheetToEnglishObjects(flowConfigSheet, MAP_FLOW_CONFIG);

    return {
        success: true,
        version: SCRIPT_VERSION,
        masters: masterData,
        trackings: trackingData,
        configs: configData,
        flowConfigs: flowConfigData,
        webAppUrl: getWebAppUrl()
    };
}

/**
 * 新建公文收發記錄
 */
function createDocument(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);

    if (!masterSheet || !trackingSheet) {
        throw new Error("工作表未正確初始化");
    }

    const docId = (payload && payload.docId) ? String(payload.docId).trim() : "";
    if (!docId) {
        throw new Error("文件編號不能為空");
    }

    const existingDocs = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
    const exists = existingDocs.some(function (d) { return String(d.docId) === String(docId); });
    if (exists) {
        throw new Error("文件編號 " + docId + " 已存在");
    }

    // 1. 處理附件上傳
    let fileId = "";
    let fileName = "";
    if (payload.fileBase64) {
        if (payload.fileBase64.indexOf("EXISTING_FILE_ID:") === 0) {
            fileId = payload.fileBase64.substring("EXISTING_FILE_ID:".length);
            fileName = payload.fileName || "範本附件";
        } else if (payload.fileBase64.indexOf("COPIED_TEMPLATE_ID:") === 0) {
            const templateId = payload.fileBase64.substring("COPIED_TEMPLATE_ID:".length);
            try {
                const folder = getOrCreateAttachmentFolder();
                const templateFile = DriveApp.getFileById(templateId);
                const destFileName = docId + "-" + (payload.category || "無分類") + "-1-" + (payload.initAction || "登記文件") + "_" + templateFile.getName();
                const file = templateFile.makeCopy(destFileName, folder);
                try {
                    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
                } catch (e) {
                    console.warn("無法設定共用權限:", e);
                }
                fileId = file.getId();
                fileName = file.getName();
            } catch (err) {
                console.error("複製範本建立附件失敗: " + err);
            }
        } else {
            let cleanBase64 = payload.fileBase64;
            if (cleanBase64.indexOf(",") > -1) {
                cleanBase64 = cleanBase64.split(",")[1];
            }
            const cat = payload.category || "無分類";
            const act = payload.initAction || "登記文件";
            let originalName = payload.fileName || "";
            let ext = ".pdf";
            let mimeType = "application/pdf";
            if (originalName) {
                let lastDotIdx = originalName.lastIndexOf(".");
                if (lastDotIdx > -1) {
                    ext = originalName.substring(lastDotIdx).toLowerCase();
                }
            }
            if (ext === ".docx") {
                mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            } else if (ext === ".doc") {
                mimeType = "application/msword";
            }
            const destFileName = docId + "-" + cat + "-1-" + act + ext;
            try {
                const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, destFileName);
                const folder = getOrCreateAttachmentFolder();
                const file = folder.createFile(blob);
                fileId = file.getId();
                fileName = file.getName();
            } catch (err) {
                console.error("建立附件失敗: " + err);
            }
        }
    }

    const initAction = payload.initAction || "登記文件";
    const initDesc = payload.initDesc || "完成文件登記";
    const dueDate = payload.dueDate || "";
    const initStatus = payload.initStatus || "待掃描";

    // 2. 寫入主表 (利用英文字段防禦映射)
    const masterObj = {
        docId: docId,
        date: payload.date || getTodayString(),
        type: payload.type || "收文",
        vendor: payload.vendor || "",
        tag: payload.category || "",
        subject: payload.subject || "",
        assignee: payload.assignee || "",
        order: 1,
        lastAction: initAction,
        status: initStatus,
        fileName: fileName,
        fileId: fileId,
        signatureFileId: payload.signatureFileId || "",
        relatedDocId: payload.relatedDocId || "",
        dueDate: dueDate
    };
    const mainRowIdx = appendEnglishObject(masterSheet, MAP_MASTER, masterObj);

    // 建立雙向關聯
    if (payload.relatedDocId) {
        try {
            const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
            for (let i = 0; i < masters.length; i++) {
                if (String(masters[i].docId).trim() === String(payload.relatedDocId).trim()) {
                    masters[i].relatedDocId = docId;
                    writeEnglishObjectToRow(masterSheet, MAP_MASTER, masters[i]._rowIndex, masters[i]);
                    break;
                }
            }
        } catch (linkErr) {
            console.warn("無法建立雙向關聯: " + linkErr.toString());
        }
    }

    // 3. 建立第一筆歷程 (利用英文字段防禦映射)
    const trackingObj = {
        docId: docId,
        order: 1,
        date: payload.date || getTodayString(),
        action: initAction,
        description: initDesc,
        dueDate: dueDate,
        status: initStatus,
        fileName: fileName,
        fileId: fileId,
        signatureFileId: payload.signatureFileId || ""
    };
    appendEnglishObject(trackingSheet, MAP_TRACKING, trackingObj);

    SpreadsheetApp.flush();

    return {
        success: true,
        message: "公文案號 " + docId + " 已成功建立。"
    };
}

/**
 * FSM 狀態移轉：新增進度追蹤紀錄並更新主表狀態
 */
function addTrackingRecord(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);
    const masterSheet = ss.getSheetByName(SHEET_MASTER);

    if (!trackingSheet || !masterSheet) {
        throw new Error("工作表未確實初始化");
    }

    const docId = (payload && payload.docId) ? String(payload.docId).trim() : "";
    if (!docId) {
        throw new Error("文件編號不能為空");
    }

    // 自動計算歷程處理次數
    const trackingData = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
    const matchedRecords = trackingData.filter(function (r) { return String(r.docId) === String(docId); });
    const nextOrder = matchedRecords.length + 1;

    // 1. 處理歷程附件上傳
    let fileId = "";
    let fileName = "";
    if (payload.fileBase64) {
        if (payload.fileBase64.indexOf("EXISTING_FILE_ID:") === 0) {
            fileId = payload.fileBase64.substring("EXISTING_FILE_ID:".length);
            fileName = payload.fileName || "歷程附件";
        } else if (payload.fileBase64.indexOf("COPIED_TEMPLATE_ID:") === 0) {
            const templateId = payload.fileBase64.substring("COPIED_TEMPLATE_ID:".length);
            try {
                const folder = getOrCreateAttachmentFolder();
                const templateFile = DriveApp.getFileById(templateId);
                let category = "無分類";
                const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
                const matchedM = masters.find(function (m) { return m.docId === docId; });
                if (matchedM) {
                    category = matchedM.tag || "無分類";
                }
                const act = payload.action || "推進狀態";
                const destFileName = docId + "-" + category + "-" + nextOrder + "-" + act + "_" + templateFile.getName();
                const file = templateFile.makeCopy(destFileName, folder);
                try {
                    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
                } catch (e) {
                    console.warn("無法設定共用權限:", e);
                }
                fileId = file.getId();
                fileName = file.getName();
            } catch (err) {
                console.error("複製範本建立歷程附件失敗: " + err);
            }
        } else {
            let cleanBase64 = payload.fileBase64;
            if (cleanBase64.indexOf(",") > -1) {
                cleanBase64 = cleanBase64.split(",")[1];
            }
            let category = "無分類";
            const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
            const matchedM = masters.find(function (m) { return m.docId === docId; });
            if (matchedM) {
                category = matchedM.tag || "無分類";
            }
            const act = payload.action || "狀態流轉";
            let originalName = payload.fileName || "";
            let ext = ".pdf";
            let mimeType = "application/pdf";
            if (originalName) {
                let lastDotIdx = originalName.lastIndexOf(".");
                if (lastDotIdx > -1) {
                    ext = originalName.substring(lastDotIdx).toLowerCase();
                }
            }
            if (ext === ".docx") {
                mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            } else if (ext === ".doc") {
                mimeType = "application/msword";
            }
            const destFileName = docId + "-" + category + "-" + nextOrder + "-" + act + ext;
            try {
                const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, destFileName);
                const folder = getOrCreateAttachmentFolder();
                const file = folder.createFile(blob);
                fileId = file.getId();
                fileName = file.getName();
            } catch (err) {
                console.error("建立歷程附件失敗: " + err);
            }
        }
    }

    // 2. 寫入歷程表 (防禦表頭映射)
    const trackingObj = {
        docId: docId,
        order: nextOrder,
        date: payload.date || getTodayString(),
        action: payload.action || "進行動作",
        description: payload.description || "",
        dueDate: payload.dueDate || "",
        status: payload.status || "流轉狀態",
        fileName: fileName,
        fileId: fileId,
        signatureFileId: payload.signatureFileId || ""
    };
    appendEnglishObject(trackingSheet, MAP_TRACKING, trackingObj);

    // 3. 更新主表對應屬性
    const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
    const matchedMaster = masters.find(function (m) { return m.docId === docId; });
    if (matchedMaster) {
        matchedMaster.order = nextOrder;
        matchedMaster.lastAction = payload.action || "";
        matchedMaster.status = payload.status || "";
        if (fileId) {
            matchedMaster.fileName = fileName;
            matchedMaster.fileId = fileId;
        }
        if (payload.signatureFileId) {
            matchedMaster.signatureFileId = payload.signatureFileId;
        }
        if (payload.dueDate) {
            matchedMaster.dueDate = payload.dueDate;
        }
        writeEnglishObjectToRow(masterSheet, MAP_MASTER, matchedMaster._rowIndex, matchedMaster);
    }

    let message = "公文 " + docId + " 狀態已成功流轉至 [" + (payload.status || "完成") + "]";

    // FSM 連動：若收文觸發了「啟動回文」，自動在背景建立對應的發文起草案並鏈結
    if (payload.action === "啟動回文" && matchedMaster) {
        const newDocId = docId + "_回";
        const newDocPayload = {
            docId: newDocId,
            date: getTodayString(),
            type: "發文",
            vendor: matchedMaster.vendor || "",
            tag: "回文",
            subject: "回覆：" + (matchedMaster.subject || ""),
            assignee: matchedMaster.assignee || "",
            order: 1,
            lastAction: "起草中",
            status: "起草中",
            fileName: "",
            fileId: "",
            signatureFileId: "",
            relatedDocId: docId,
            dueDate: ""
        };

        // 寫入發文主表
        appendEnglishObject(masterSheet, MAP_MASTER, newDocPayload);

        // 寫入發文初始歷程
        const subTrackObj = {
            docId: newDocId,
            order: 1,
            date: getTodayString(),
            action: "起草中",
            description: "針對收文案號 " + docId + " 自動建立回文起草",
            dueDate: "",
            status: "起草中",
            fileName: "",
            fileId: "",
            signatureFileId: ""
        };
        appendEnglishObject(trackingSheet, MAP_TRACKING, subTrackObj);

        // 將收文的關聯編號回寫成該發文
        matchedMaster.relatedDocId = newDocId;
        writeEnglishObjectToRow(masterSheet, MAP_MASTER, matchedMaster._rowIndex, matchedMaster);

        message += "，且系統已自動建立對應發文起草案 (案號：" + newDocId + ")";
    }

    SpreadsheetApp.flush();

    return {
        success: true,
        fileId: fileId,
        fileName: fileName,
        message: message
    };
}

/**
 * 格式化日期為 YYYY-MM-DD
 */
function formatDate(date) {
    const y = date.getFullYear();
    const m = ("0" + (date.getMonth() + 1)).slice(-2);
    const d = ("0" + date.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
}

/**
 * 取得今日 YYYY-MM-DD 格式
 */
function getTodayString() {
    const today = new Date();
    return formatDate(today);
}

/**
 * 依據民國年序列自動生成下一筆唯一的公文文號
 */
function getNextDocId(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = payload.date || getTodayString();
    const category = payload.category || "";

    const parts = dateStr.split("-");
    if (parts.length < 3) {
        throw new Error("日期格式錯誤，必須為 YYYY-MM-DD");
    }
    const rocYear = parseInt(parts[0], 10) - 1911;
    const mm = parts[1];
    const dd = parts[2];
    const datePart = String(rocYear) + mm + dd; // e.g. "1150704"

    let prefix = "";
    const configSheet = ss.getSheetByName(SHEET_CONFIG);
    if (configSheet) {
        const configs = readSheetToEnglishObjects(configSheet, MAP_CONFIG);
        const matched = configs.find(function (c) { return String(c.tag).trim() === String(category).trim(); });
        if (matched && matched.prefix) {
            prefix = String(matched.prefix).trim();
        }
    }

    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    let maxSeq = 0;
    if (masterSheet) {
        const existingDocs = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
        existingDocs.forEach(function (d) {
            const idStr = String(d.docId).trim();
            let rawId = idStr;
            if (prefix && idStr.startsWith(prefix)) {
                rawId = idStr.substring(prefix.length);
            }
            if (rawId.startsWith(datePart) && rawId.length === datePart.length + 2) {
                const seqStr = rawId.substring(datePart.length);
                const seq = parseInt(seqStr, 10);
                if (!isNaN(seq) && seq > maxSeq) {
                    maxSeq = seq;
                }
            }
        });
    }

    const nextSeq = maxSeq + 1;
    const seqStr = ("0" + nextSeq).slice(-2);
    const nextDocId = prefix + datePart + seqStr;

    return {
        success: true,
        docId: nextDocId
    };
}

/**
 * 儲存系統配置參數
 */
function saveSystemConfig(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = ss.getSheetByName(SHEET_CONFIG);
    if (!configSheet) {
        configSheet = ss.insertSheet(SHEET_CONFIG);
    }
    configSheet.clear();
    configSheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS])
        .setFontWeight("bold")
        .setBackground("#e2e8f0");

    const rows = Array.isArray(payload) ? payload : (payload && payload.rows ? payload.rows : []);
    if (rows.length > 0) {
        rows.forEach(function (r) {
            appendEnglishObject(configSheet, MAP_CONFIG, r);
        });
    }

    SpreadsheetApp.flush();
    return {
        success: true,
        message: "系統往來設定已成功同步寫回 Google Sheet。"
    };
}

/**
 * 儲存 FSM 流程配置 (含 範本ID 備用)
 */
function saveFlowConfig(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let flowSheet = ss.getSheetByName(SHEET_FLOW_CONFIG);
    if (!flowSheet) {
        flowSheet = ss.insertSheet(SHEET_FLOW_CONFIG);
    }
    flowSheet.clear();
    flowSheet.getRange(1, 1, 1, FLOW_CONFIG_HEADERS.length).setValues([FLOW_CONFIG_HEADERS])
        .setFontWeight("bold")
        .setBackground("#e2e8f0");

    const rows = Array.isArray(payload) ? payload : (payload && payload.rows ? payload.rows : []);
    if (rows.length > 0) {
        rows.forEach(function (r) {
            r.order = Number(r.order) || 1;
            appendEnglishObject(flowSheet, MAP_FLOW_CONFIG, r);
        });
    }

    SpreadsheetApp.flush();
    return {
        success: true,
        message: "流程配置已成功同步寫回 Google Sheet。"
    };
}

/**
 * 取得當前試算表父級資料夾
 */
function getSpreadsheetFolder() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const file = DriveApp.getFileById(ss.getId());
        const parents = file.getParents();
        if (parents.hasNext()) {
            return parents.next();
        }
    } catch (e) {
        console.error("取得試算表所在資料夾失敗:", e);
    }
    return DriveApp.getRootFolder();
}

/**
 * 建立或獲取「收發文附件」雲端硬碟共用資料夾
 */
function getOrCreateAttachmentFolder() {
    const parentFolder = getSpreadsheetFolder();
    const folderName = "收發文附件";
    const folders = parentFolder.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        folder = parentFolder.createFolder(folderName);
    }
    try {
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
        console.warn("無法設定收發文附件資料夾公開共用:", e.toString());
    }
    return folder;
}

/**
 * 同步更新主表與最新歷程之文件連結
 */
function updateDocumentFile(docId, fileId, fileName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    if (masterSheet) {
        const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
        const matchedMaster = masters.find(function (m) { return m.docId === docId; });
        if (matchedMaster) {
            matchedMaster.fileName = fileName;
            matchedMaster.fileId = fileId;
            writeEnglishObjectToRow(masterSheet, MAP_MASTER, matchedMaster._rowIndex, matchedMaster);
        }
    }

    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);
    if (trackingSheet) {
        const trackings = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
        const docTrackings = trackings.filter(function (t) { return t.docId === docId; });
        if (docTrackings.length > 0) {
            let maxTrack = docTrackings[0];
            docTrackings.forEach(function (t) {
                if (t.order > maxTrack.order) maxTrack = t;
            });
            maxTrack.fileName = fileName;
            maxTrack.fileId = fileId;
            writeEnglishObjectToRow(trackingSheet, MAP_TRACKING, maxTrack._rowIndex, maxTrack);
        }
    }
}

/**
 * 上傳數位簽章圖片
 */
function uploadSignature(payload) {
    const docId = payload.docId;
    const base64Data = payload.base64Data;
    if (!docId || !base64Data) {
        throw new Error("參數不足");
    }

    let cleanBase64 = base64Data;
    if (base64Data.indexOf(",") > -1) {
        cleanBase64 = base64Data.split(",")[1];
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);

    let category = "無分類";
    if (masterSheet) {
        const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
        const matchedM = masters.find(function (m) { return m.docId === docId; });
        if (matchedM) {
            category = matchedM.tag || "無分類";
        }
    }

    let nextOrder = 1;
    if (trackingSheet) {
        const trackingData = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
        const matchedRecords = trackingData.filter(function (r) { return String(r.docId) === String(docId); });
        nextOrder = matchedRecords.length + 1;
    }

    const act = payload.action || "狀態流轉";
    // 依據上傳附件規格: docId + "-" + category + "-" + nextOrder + "-" + act + ext
    const fileName = docId + "-" + category + "-" + nextOrder + "-" + act + "-簽收單.png";

    const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/png", fileName);
    const folder = getOrCreateAttachmentFolder();
    const file = folder.createFile(blob);

    try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
        console.warn("無法設定簽收單公開共用:", shareErr.toString());
    }

    updateSignatureFile(docId, file.getId());

    SpreadsheetApp.flush();

    return {
        success: true,
        fileId: file.getId(),
        message: "簽收單上傳成功"
    };
}

/**
 * 同步更新最新歷程與主表之簽章 ID
 */
function updateSignatureFile(docId, signatureFileId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    if (masterSheet) {
        const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
        const matchedMaster = masters.find(function (m) { return m.docId === docId; });
        if (matchedMaster) {
            matchedMaster.signatureFileId = signatureFileId;
            writeEnglishObjectToRow(masterSheet, MAP_MASTER, matchedMaster._rowIndex, matchedMaster);
        }
    }

    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);
    if (trackingSheet) {
        const trackings = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
        const docTrackings = trackings.filter(function (t) { return t.docId === docId; });
        if (docTrackings.length > 0) {
            let maxTrack = docTrackings[0];
            docTrackings.forEach(function (t) {
                if (t.order > maxTrack.order) maxTrack = t;
            });
            maxTrack.signatureFileId = signatureFileId;
            writeEnglishObjectToRow(trackingSheet, MAP_TRACKING, maxTrack._rowIndex, maxTrack);
        }
    }
}

/**
 * 上傳常規 PDF 文件附件
 */
function uploadAttachment(payload) {
    const docId = payload.docId;
    const originalName = payload.fileName;
    const base64Data = payload.base64Data;
    const oldFileId = payload.oldFileId;

    if (!docId || !originalName || !base64Data) {
        throw new Error("參數不足");
    }

    let cleanBase64 = base64Data;
    if (base64Data.indexOf(",") > -1) {
        cleanBase64 = base64Data.split(",")[1];
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_MASTER);
    const trackingSheet = ss.getSheetByName(SHEET_TRACKING);

    let category = "無分類";
    if (masterSheet) {
        const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
        const matchedM = masters.find(function (m) { return m.docId === docId; });
        if (matchedM) {
            category = matchedM.tag || "無分類";
        }
    }

    let nextOrder = 1;
    if (trackingSheet) {
        const trackingData = readSheetToEnglishObjects(trackingSheet, MAP_TRACKING);
        const matchedRecords = trackingData.filter(function (r) { return String(r.docId) === String(docId); });
        nextOrder = matchedRecords.length;
    }

    let ext = ".pdf";
    let mimeType = "application/pdf";
    if (originalName) {
        let lastDotIdx = originalName.lastIndexOf(".");
        if (lastDotIdx > -1) {
            ext = originalName.substring(lastDotIdx).toLowerCase();
        }
    }
    if (ext === ".docx") {
        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    } else if (ext === ".doc") {
        mimeType = "application/msword";
    }

    const destFileName = docId + "-" + category + "-" + nextOrder + "-重傳" + ext;

    const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, destFileName);
    const folder = getOrCreateAttachmentFolder();
    const file = folder.createFile(blob);

    updateDocumentFile(docId, file.getId(), file.getName());

    if (oldFileId && oldFileId !== "uploading_temp_id") {
        try {
            DriveApp.getFileById(oldFileId).setTrashed(true);
        } catch (trashErr) {
            console.warn("無法移除舊附件: " + oldFileId);
        }
    }

    SpreadsheetApp.flush();

    return {
        success: true,
        fileId: file.getId(),
        fileName: file.getName(),
        message: "附件更新成功"
    };
}

/**
 * 取得網頁應用程式 (Web App) 部署網址
 */
function getWebAppUrl() {
    try {
        return ScriptApp.getService().getUrl();
    } catch (e) {
        console.warn("無法取得 GAS 部署網址:", e);
        return "";
    }
}

/**
 * 複製並連結文件範本
 */
function copyTemplateFile(payload) {
    const docId = payload.docId;
    const templateId = payload.templateId;
    const fileName = payload.fileName || "公文附件_" + docId;

    if (!templateId) {
        throw new Error("請提供有效範本ID");
    }

    const folder = getOrCreateAttachmentFolder();
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(fileName, folder);

    try {
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (e) {
        console.warn("無法設定複製檔案權限:", e);
    }

    updateDocumentFile(docId, newFile.getId(), newFile.getName());

    return {
        success: true,
        fileId: newFile.getId(),
        fileName: newFile.getName(),
        fileUrl: newFile.getUrl(),
        message: "範本複製並連結成功"
    };
}

/**
 * 刪除並解除連結複製出的檔案
 */
function deleteTemplateFile(payload) {
    const fileId = payload.fileId;
    const docId = payload.docId;

    if (!fileId) {
        return { success: false, message: "無效的檔案ID" };
    }

    try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
    } catch (e) {
        console.warn("移至垃圾桶失敗:", e.toString());
    }

    if (docId) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const masterSheet = ss.getSheetByName(SHEET_MASTER);
        if (masterSheet) {
            const masters = readSheetToEnglishObjects(masterSheet, MAP_MASTER);
            const matchedMaster = masters.find(function (m) { return m.docId === docId; });
            if (matchedMaster && matchedMaster.fileId === fileId) {
                matchedMaster.fileName = "";
                matchedMaster.fileId = "";
                writeEnglishObjectToRow(masterSheet, MAP_MASTER, matchedMaster._rowIndex, matchedMaster);
            }
        }
    }

    return {
        success: true,
        message: "複製出的檔案已成功刪除並解約"
    };
}
