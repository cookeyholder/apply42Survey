// 安全性和效能常數
const MAX_SHEET_ROWS = 10000;
const REQUEST_TIMEOUT = 30000; // 30 秒
const MAX_PARAMETER_LENGTH = 500;
const ALLOWED_PARAMETERS = ['系統名稱', '系統關閉時間', '報名學校代碼'];

/**
 * @description 安全地取得工作表資料
 * @param {Sheet} sheet - 工作表物件
 * @param {Array} requiredHeaders - 必要的標頭欄位
 * @returns {{headers: Array, data: Array}|null} 工作表資料或 null
 */
function getSheetDataSafely(sheet, requiredHeaders = []) {
    try {
        if (!sheet) {
            Logger.log('工作表不存在');
            return null;
        }

        const numRows = sheet.getLastRow();
        const numCols = sheet.getLastColumn();

        // 檢查工作表大小
        if (numRows > MAX_SHEET_ROWS || numCols > 100) {
            Logger.log('工作表過大：%d 列 %d 欄', numRows, numCols);
            return null;
        }

        if (numRows === 0 || numCols === 0) {
            Logger.log('工作表為空');
            return { headers: [], data: [] };
        }

        // 取得標頭
        const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

        // 驗證必要標頭
        if (requiredHeaders.length > 0) {
            const missingHeaders = requiredHeaders.filter(
                (header) => !headers.includes(header)
            );
            if (missingHeaders.length > 0) {
                Logger.log('工作表缺少必要標頭：%s', missingHeaders.join(', '));
                return null;
            }
        }

        // 取得資料（如果有的話）
        let data = [];
        if (numRows > 1) {
            data = sheet.getRange(2, 1, numRows - 1, numCols).getValues();
        }

        Logger.log(
            '成功讀取工作表 %s：%d 列資料',
            sheet.getName(),
            data.length
        );
        return { headers, data };
    } catch (error) {
        Logger.log('讀取工作表時發生錯誤：%s', error.message);
        return null;
    }
}

/**
 * @description 驗證使用者電子郵件
 * @param {string} email - 電子郵件地址
 * @returns {boolean} 是否為有效的電子郵件
 */
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return (
        typeof email === 'string' &&
        email.length > 0 &&
        email.length <= 100 &&
        emailRegex.test(email)
    );
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
const configSheet = ss.getSheetByName('參數設定');
const examDataSheet = ss.getSheetByName('統測報名資料');
const choicesSheet = ss.getSheetByName('志願選項');
const studentChoiceSheet = ss.getSheetByName('考生志願列表');
const limitOfSchoolsSheet = ss.getSheetByName('可報名之系科組學程數');
const forImportSheet = ss.getSheetByName('匯入報名系統');
const mentorSheet = ss.getSheetByName('導師名單');
const limitsOfChoices = 6; // 最多可填的志願數量

/**
 * @description 建立自訂功能表「志願調查系統」
 */
function onOpen() {
    try {
        SpreadsheetApp.getUi()
            .createMenu('志願調查系統')
            .addItem('匯出報名用CSV', 'exportCsv')
            .addItem('清除快取', 'clearAllCache')
            .addToUi();
        Logger.log('功能表建立成功');
    } catch (error) {
        Logger.log('建立功能表時發生錯誤：%s', error.message);
    }
}

/**
 * @description 驗證請求參數的安全性
 * @param {Object} parameters - 請求參數
 * @returns {boolean} 參數是否安全
 */
function validateRequestParameters(parameters) {
    if (!parameters || typeof parameters !== 'object') {
        return false;
    }

    // 檢查參數數量
    if (Object.keys(parameters).length > 20) {
        Logger.log('請求參數過多');
        return false;
    }

    // 檢查每個參數
    for (const [key, value] of Object.entries(parameters)) {
        if (typeof key !== 'string' || key.length > 100) {
            Logger.log('無效的參數鍵：%s', key);
            return false;
        }

        if (Array.isArray(value)) {
            if (value.length > 10) {
                Logger.log('參數陣列過大：%s', key);
                return false;
            }
            for (const item of value) {
                if (
                    typeof item === 'string' &&
                    item.length > MAX_PARAMETER_LENGTH
                ) {
                    Logger.log('參數值過長：%s', key);
                    return false;
                }
            }
        } else if (
            typeof value === 'string' &&
            value.length > MAX_PARAMETER_LENGTH
        ) {
            Logger.log('參數值過長：%s', key);
            return false;
        }
    }

    return true;
}

/**
 * @description 處理 GET 請求，回傳表單頁面或錯誤訊息（安全版本）
 * @param {Object} e - 請求參數
 * @returns {HtmlOutput} HTML 輸出
 */
function doGet(e) {
    try {
        // 驗證請求參數
        if (!validateRequestParameters(e.parameters)) {
            Logger.log('請求參數驗證失敗');
            return HtmlService.createHtmlOutput(
                '<div style="padding: 20px; color: red;">無效的請求參數</div>'
            );
        }

        Logger.log('doGet 請求參數：%s', JSON.stringify(e.parameters));

        const user = getUserData();
        Logger.log('doGet 取得的使用者資料：%s', JSON.stringify(user));

        const parameters = getParameters();
        Logger.log('doGet 取得的系統設定資訊：%s', JSON.stringify(parameters));

        // 如果使用者未登入或登入的不在允許名單之中
        if (!user) {
            return HtmlService.createHtmlOutput(`
                <div style="padding: 20px; text-align: center; font-family: Arial, sans-serif;">
                    <h2 style="color: #d32f2f;">存取受限</h2>
                    <p>請先登入學校的信箱帳號，並使用 Chrome 瀏覽器。</p>
                    <p style="color: #666; font-size: 0.9em;">如有問題請聯絡系統管理員</p>
                </div>
            `);
        }

        // 驗證系統參數
        if (!parameters || !parameters['系統名稱']) {
            Logger.log('系統參數不完整');
            return HtmlService.createHtmlOutput(
                '<div style="padding: 20px; color: red;">系統設定錯誤，請聯絡管理員</div>'
            );
        }

        // 如果使用者資料中有「統一入學測驗報名序號」，則表示他是學生
        if (user['統一入學測驗報名序號']) {
            return renderStudentPage(user, parameters);
        }

        // 如果使用者資料中沒有「統一入學測驗報名序號」，則表示他是導師
        if (!user['統一入學測驗報名序號']) {
            return renderMentorPage(user, parameters);
        }
    } catch (err) {
        Logger.log('doGet 發生錯誤：%s\n%s', err.message, err.stack);

        // 轉義 HTML 特殊字符以防止 XSS
        const escapeHtml = (text) => {
            return text
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#x27;');
        };

        return HtmlService.createHtmlOutput(`
            <div style="padding: 20px; text-align: center; font-family: Arial, sans-serif;">
                <h2 style="color: #d32f2f;">系統錯誤</h2>
                <p>很抱歉，系統發生錯誤，請稍後再試。</p>
                <p style="color: #666; font-size: 0.9em;">錯誤時間：${new Date().toLocaleString(
                    'zh-TW'
                )}</p>
                <div style="margin-top: 20px; padding: 10px; background-color: #f5f5f5; border-left: 4px solid #d32f2f; text-align: left;">
                    <h3 style="color: #d32f2f; margin-top: 0;">錯誤詳情：</h3>
                    <p><strong>錯誤訊息：</strong> ${escapeHtml(
                        err.message || '未知錯誤'
                    )}</p>
                    <details style="margin-top: 10px;">
                        <summary style="cursor: pointer; color: #666;">顯示詳細堆疊資訊</summary>
                        <pre style="background-color: #fff; padding: 10px; border: 1px solid #ddd; margin-top: 10px; font-size: 12px; overflow: auto;">${escapeHtml(
                            err.stack || '無堆疊資訊'
                        )}</pre>
                    </details>
                </div>
            </div>
        `);
    }
}

/**
 * @description 渲染學生頁面
 * @param {Object} user - 使用者資料
 * @param {Object} parameters - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderStudentPage(user, parameters) {
    try {
        const template = HtmlService.createTemplateFromFile('index');
        template.loginEmail = Session.getActiveUser().getEmail();
        template.serviceUrl = getServiceUrl();
        template.user = user;
        template.parameters = parameters;
        template.notifications = getNotifications(parameters);
        template.limitOfSchools = getLimitOfSchools();

        const optionData = getOptionData(user);
        template.isJoined = optionData.isJoined;
        template.selectedChoices = optionData.selectedChoices;
        template.departmentOptions = optionData.departmentOptions;

        return setXFrameOptionsSafely(
            template.evaluate().setTitle('四技二專甄選入學志願調查系統')
        );
    } catch (error) {
        Logger.log('渲染學生頁面時發生錯誤：%s', error.message);
        throw error;
    }
}

/**
 * @description 渲染導師頁面
 * @param {Object} user - 使用者資料
 * @param {Object} parameters - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderMentorPage(user, parameters) {
    try {
        const studentData = getTraineesDepartmentChoices(user);
        const template = HtmlService.createTemplateFromFile('mentorView');

        template.loginEmail = Session.getActiveUser().getEmail();
        template.serviceUrl = getServiceUrl();
        template.user = user;
        template.parameters = parameters;
        template.headers = studentData.headers;
        template.data = studentData.data;

        return setXFrameOptionsSafely(
            template.evaluate().setTitle('導師查詢班級學生志願')
        );
    } catch (error) {
        Logger.log('渲染導師頁面時發生錯誤：%s', error.message);
        throw error;
    }
}

/**
 * @description 處理 POST 請求（安全版本）
 * @param {Object} e - 請求參數
 * @returns {HtmlOutput|ContentService.TextOutput} 回應內容
 */
function doPost(e) {
    try {
        // 驗證請求參數
        if (!validateRequestParameters(e.parameters)) {
            Logger.log('POST 請求參數驗證失敗');
            return ContentService.createTextOutput(
                '無效的請求參數'
            ).setMimeType(ContentService.MimeType.TEXT);
        }

        Logger.log('doPost 請求參數：%s', JSON.stringify(e.parameters));

        const user = getUserData();
        if (!user || !user['統一入學測驗報名序號']) {
            Logger.log('無效的使用者或非學生帳號嘗試提交');
            return ContentService.createTextOutput('存取被拒絕').setMimeType(
                ContentService.MimeType.TEXT
            );
        }

        const parameters = getParameters();
        if (!parameters || !parameters['系統關閉時間']) {
            Logger.log('系統參數不完整');
            return ContentService.createTextOutput('系統設定錯誤').setMimeType(
                ContentService.MimeType.TEXT
            );
        }

        // 檢查截止時間
        const endTime = new Date(parameters['系統關閉時間']);
        const now = new Date();
        const tolerance = 60000; // 1 分鐘容忍時間

        if (isNaN(endTime.getTime())) {
            Logger.log('系統關閉時間格式錯誤：%s', parameters['系統關閉時間']);
            return ContentService.createTextOutput(
                '系統時間設定錯誤'
            ).setMimeType(ContentService.MimeType.TEXT);
        }

        if (now - endTime > tolerance) {
            Logger.log(
                '提交時間已過截止時間，現在：%s，截止：%s',
                now,
                endTime
            );
            return ContentService.createTextOutput(
                '志願調查已結束'
            ).setMimeType(ContentService.MimeType.TEXT);
        }

        // 驗證和清理輸入資料
        const joinedParam = String(
            e.parameters.isJoinedInput?.[0] || '否'
        ).trim();
        const isJoined = joinedParam === '是';

        let departmentChoices = [];
        if (isJoined) {
            // 取得並驗證志願選擇
            for (let i = 1; i <= limitsOfChoices; i++) {
                const choice = String(
                    e.parameters[`departmentChoices_${i}`]?.[0] || ''
                ).trim();
                // 驗證志願格式（應為6位數字）
                if (choice && !/^\d{6}$/.test(choice)) {
                    Logger.log('無效的志願格式：%s', choice);
                    return ContentService.createTextOutput(
                        '無效的志願格式'
                    ).setMimeType(ContentService.MimeType.TEXT);
                }
                departmentChoices.push(choice);
            }

            // 排序志願（空值排到後面）
            departmentChoices.sort((a, b) => {
                if (a === '' && b === '') return 0;
                if (a === '') return 1;
                if (b === '') return -1;
                return Number(a) - Number(b);
            });
        }

        // 更新資料
        const userEmail = Session.getActiveUser().getEmail();
        const row = findValueRow(studentChoiceSheet, userEmail);

        if (!row || row === 0) {
            Logger.log('找不到使用者資料列：%s', userEmail);
            return ContentService.createTextOutput(
                '找不到使用者資料'
            ).setMimeType(ContentService.MimeType.TEXT);
        }

        // 準備更新的資料
        const updateData = isJoined
            ? [joinedParam, ...departmentChoices]
            : [joinedParam, '', '', '', '', '', ''];

        updateSpecificRow(row, updateData);
        Logger.log('成功更新使用者 %s 的志願資料', userEmail);

        // 渲染成功頁面
        return renderSuccessPage(user, parameters);
    } catch (err) {
        Logger.log('doPost 發生錯誤：%s\n%s', err.message, err.stack);
        return ContentService.createTextOutput(
            '系統錯誤，請稍後再試'
        ).setMimeType(ContentService.MimeType.TEXT);
    }
}

/**
 * @description 渲染成功頁面
 * @param {Object} user - 使用者資料
 * @param {Object} parameters - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderSuccessPage(user, parameters) {
    try {
        const template = HtmlService.createTemplateFromFile('success');
        template.loginEmail = Session.getActiveUser().getEmail();
        template.serviceUrl = getServiceUrl();
        template.user = user;
        template.parameters = parameters;
        template.notifications = getNotifications(parameters);
        template.limitOfSchools = getLimitOfSchools();

        const optionData = getOptionData(user);
        template.isJoined = optionData.isJoined;
        template.selectedChoices = optionData.selectedChoices;
        template.departmentOptions = optionData.departmentOptions;

        return setXFrameOptionsSafely(
            template.evaluate().setTitle('四技二專甄選入學志願調查系統')
        );
    } catch (error) {
        Logger.log('渲染成功頁面時發生錯誤：%s', error.message);
        throw error;
    }
}

/**
 * @description 清除所有快取（管理功能）
 */
function clearAllCache() {
    try {
        const cache = CacheService.getScriptCache();
        cache.removeAll(Object.values(CACHE_KEYS));
        Logger.log('已清除所有快取');

        const ui = SpreadsheetApp.getUi();
        ui.alert('快取已清除', '所有快取資料已成功清除。', ui.ButtonSet.OK);
    } catch (error) {
        Logger.log('清除快取時發生錯誤：%s', error.message);
        const ui = SpreadsheetApp.getUi();
        ui.alert(
            '錯誤',
            '清除快取時發生錯誤：' + error.message,
            ui.ButtonSet.OK
        );
    }
}

/**
 * @description 安全地設定 XFrame 選項
 * @param {HtmlOutput} htmlOutput - HTML 輸出物件
 * @returns {HtmlOutput} 設定完成的 HTML 輸出物件
 */
function setXFrameOptionsSafely(htmlOutput) {
    try {
        // 檢查 HtmlService.XFrameOptionsMode 是否存在
        if (
            HtmlService &&
            HtmlService.XFrameOptionsMode &&
            HtmlService.XFrameOptionsMode.SAMEORIGIN
        ) {
            return htmlOutput.setXFrameOptionsMode(
                HtmlService.XFrameOptionsMode.SAMEORIGIN
            );
        } else {
            Logger.log('XFrameOptionsMode.SAMEORIGIN 未定義，跳過設定');
            return htmlOutput;
        }
    } catch (error) {
        Logger.log('設定 XFrameOptionsMode 時發生錯誤：%s', error.message);
        return htmlOutput;
    }
}
