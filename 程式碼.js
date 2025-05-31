const ss = SpreadsheetApp.getActiveSpreadsheet();
const configSheet = ss.getSheetByName('參數設定');
const examDataSheet = ss.getSheetByName('統測報名資料');
const choicesSheet = ss.getSheetByName('志願選項');
const studentChoiceSheet = ss.getSheetByName('考生志願列表');
const limitOfSchoolsheet = ss.getSheetByName('可報名之系科組學程數');
const forImportSheet = ss.getSheetByName('匯入報名系統');
const limitsOfChoices = 6; // 最多可填的志願數量

/**
 * @description 建立自訂功能表「志願調查系統」
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('志願調查系統')
        .addItem('匯出報名用CSV', 'exportCsv')
        .addToUi();
}

// 快取相關常數
const CACHE_EXPIRATION = 21600; // 快取時效 6 小時
const CACHE_KEYS = {
    LIMIT_OF_SCHOOLS: 'limitOfSchools',
    DEPARTMENT_OPTIONS: 'departmentOptions',
    EXAM_DATA: 'examData',
    CHOICES_DATA: 'choicesData',
};

/**
 * @description 將大型資料分段儲存到快取
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 */
function setChunkedCacheData(key, data) {
    const cache = CacheService.getScriptCache();
    const str = JSON.stringify(data);

    // 將資料分段，每段約 90KB
    const chunkSize = 90000;
    const chunks = [];
    for (let i = 0; i < str.length; i += chunkSize) {
        chunks.push(str.slice(i, i + chunkSize));
    }

    // 儲存分段數量
    const cacheObj = {};
    cacheObj[`${key}_chunks`] = chunks.length;

    // 儲存每個分段
    chunks.forEach((chunk, i) => {
        cacheObj[`${key}_${i}`] = chunk;
    });

    cache.putAll(cacheObj, CACHE_EXPIRATION);
}

/**
 * @description 從快取中讀取分段資料並組合
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getChunkedCacheData(key) {
    const cache = CacheService.getScriptCache();
    const numChunks = Number(cache.get(`${key}_chunks`));

    if (!numChunks) {
        return null;
    }

    // 讀取所有分段
    const keys = Array.from({ length: numChunks }, (_, i) => `${key}_${i}`);
    const chunks = cache.getAll(keys);

    if (!chunks || Object.keys(chunks).length === 0) {
        return null;
    }

    // 組合所有分段
    const jsonStr = Array.from(
        { length: numChunks },
        (_, i) => chunks[`${key}_${i}`]
    ).join('');

    try {
        return JSON.parse(jsonStr);
    } catch (e) {
        Logger.log('快取資料解析錯誤：%s', e.message);
        return null;
    }
}

/**
 * @description 設定快取資料（自動判斷是否需要分段）
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 */
function setCacheData(key, data) {
    const str = JSON.stringify(data);
    if (str.length > 90000) {
        setChunkedCacheData(key, data);
    } else {
        const cache = CacheService.getScriptCache();
        cache.put(key, str, CACHE_EXPIRATION);
    }
}

/**
 * @description 取得快取資料（自動判斷是否為分段資料）
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getCacheData(key) {
    const cache = CacheService.getScriptCache();
    const chunksCount = cache.get(`${key}_chunks`);

    if (chunksCount) {
        return getChunkedCacheData(key);
    }

    const data = cache.get(key);
    return data ? JSON.parse(data) : null;
}

/**
 * @description 處理 GET 請求，回傳表單頁面或錯誤訊息
 * @param {Object} e - 請求參數
 * @returns {HtmlOutput} HTML 輸出
 */
function doGet(e) {
    try {
        Logger.log('doGet 請求參數：%s', JSON.stringify(e.parameters));
        const user = getUserData();
        Logger.log('doGet 取得的使用者資料：%s', JSON.stringify(user));
        const parameters = getParameters();
        Logger.log('doGet 取得的系統設定資訊：%s', JSON.stringify(parameters));
        if (!user) {
            return HtmlService.createHtmlOutput(
                '請先登入學校的信箱帳號，並使用 Chrome 瀏覽器。'
            );
        } else {
            const template = HtmlService.createTemplateFromFile('index');
            template.loginEmail = Session.getActiveUser().getEmail();
            template.serviceUrl = getServiceUrl();
            template.user = user;
            template.parameters = parameters;
            template.notifications = getNotifications(parameters);
            template.limitOfSchools = getLimitOfSchools();
            template.isJoined = getOptionData(user).isJoined;
            template.selectedChoices = getOptionData(user).selectedChoices;
            template.departmentOptions = getOptionData(user).departmentOptions;
            return template.evaluate().setTitle('四技二專甄選入學志願調查系統');
        }
    } catch (err) {
        Logger.log('doGet 發生錯誤：%s\n%s', err.message, err.stack);
        return HtmlService.createHtmlOutput(
            '系統錯誤，請稍後再試。<br><pre>' + err.message + '</pre>'
        );
    }
}

/**
 * @description 取得執行個體 URL
 * @returns {string} serviceUrl
 */
function getServiceUrl() {
    return ScriptApp.getService().getUrl();
}

/**
 * @description 從參數設定工作表取得所有參數
 * @returns {Object<string, any>} 參數鍵值對
 */
function getParameters() {
    const data = configSheet.getDataRange().getValues();
    return data.reduce((acc, row) => {
        const [key, value] = row;
        acc[key] = value;
        return acc;
    }, {});
}

function getNotifications(parameters) {
    const notifications = [];
    const descriptionKeys = Object.keys(parameters).filter((key) =>
        key.startsWith('說明')
    );
    descriptionKeys.forEach((key) => {
        const description = parameters[key];
        if (description) {
            notifications.push(`<li>${description}</li>`);
        }
    });

    return notifications.join('');
}

/**
 * @description 更新考生志願列表指定列的志願資料
 * @param {number} row - 列號
 * @param {Array[]} values - 二維陣列志願資料
 */
function updateSpecificRow(row, values) {
    if (!Array.isArray(values[0])) {
        values = [values];
    }

    const headers = studentChoiceSheet
        .getRange(1, 1, 1, studentChoiceSheet.getLastColumn())
        .getValues()[0];
    const startColumn = headers.indexOf('是否參加集體報名') + 1; // 「是否參加集體報名」欄位的索引
    const range = studentChoiceSheet.getRange(
        row,
        startColumn,
        1,
        limitsOfChoices + 1 // 包含「是否參加」欄位
    );
    range.setValues(values);

    Logger.log(
        '更新考生志願列表的第 %s 列，更新的值為: %s',
        row,
        JSON.stringify(values)
    );
}

// 此函數原作者為彰化高商李政燁老師，用於尋找符合的文字所在的列號(row number)
// 此處修改或增加部分如下：
// (1)變數名稱以切合其用途
// (2)加上 JSDoc 註解以便於維護和使用
// (3)增加參數設定工作表的讀取
// (4)加上 Logger.log() 記錄以利於偵錯
/**
 * @description 在指定範圍以文字搜尋取得列號
 * @param {Range} targetRange - 搜尋範圍
 * @param {string} keyword - 關鍵字
 * @returns {number} 找到的列號，未找到回傳 0
 */
function findValueRow(targetRange, keyword) {
    if (!keyword) return 0;

    const sheet =
        typeof targetRange.getSheet === 'function'
            ? targetRange.getSheet()
            : targetRange;

    const foundCell = targetRange
        .createTextFinder(keyword)
        .matchEntireCell(true)
        .matchCase(false)
        .findNext();

    Logger.log(
        '在 %s 工作表的第 %s 列找到關鍵字: %s',
        sheet.getName(),
        foundCell ? foundCell.getRow() : 0,
        keyword
    );
    return foundCell ? foundCell.getRow() : 0; // 有找到傳回 row number，否則傳回 0
}

/**
 * @description 依目前登入電子郵件從統測報名資料取得使用者資料
 * @returns {Object<string, any>|null} 使用者資料或 null
 */
function getUserData() {
    if (!examDataSheet) {
        Logger.log('考生報名資料表不存在');
        return null;
    }
    const loginEmail = Session.getActiveUser().getEmail();
    if (!loginEmail) {
        Logger.log('User not logged in');
        return null;
    }

    // 檢查快取中的考生資料
    const cachedExamData = getCacheData(CACHE_KEYS.EXAM_DATA);
    let examData;

    if (cachedExamData) {
        examData = cachedExamData;
    } else {
        examData = {
            headers: examDataSheet
                .getRange(1, 1, 1, examDataSheet.getLastColumn())
                .getValues()[0],
            data: examDataSheet.getDataRange().getValues().slice(1),
        };
        setCacheData(CACHE_KEYS.EXAM_DATA, examData);
    }

    const { headers, data } = examData;

    // 尋找使用者資料
    const userRow = data.find(
        (row) => row[headers.indexOf('信箱')] === loginEmail
    );

    if (!userRow) {
        return null;
    }

    // 檢查必要欄位
    const requiredFields = ['學號', '班級名稱', '考生姓名', '報考群(類)名稱'];
    const missingFields = requiredFields.filter(
        (field) => headers.indexOf(field) === -1
    );

    if (missingFields.length > 0) {
        Logger.log('缺少必要欄位：' + missingFields.join(', '));
        return null;
    }

    const user = headers.reduce((acc, key, idx) => {
        acc[key] = userRow[idx];
        return acc;
    }, {});

    Logger.log('getUserData() 取得使用者資料：%s', JSON.stringify(user));
    return user;
}

/**
 * @description 取得使用者的參加狀態、已選擇志願及可選擇志願列表
 * @param {Object<string, any>} [user] - 使用者資訊，若為 null 則自動取得
 * @returns {{isJoined: boolean, selectedChoices: any[], choices: string[]}}
 */
function getOptionData(user = null) {
    if (!user) {
        user = getUserData();
    }

    // 取得志願選項資料（使用快取）
    let choicesData = getCacheData(CACHE_KEYS.CHOICES_DATA);
    if (!choicesData) {
        choicesData = {
            headers: choicesSheet
                .getRange(1, 1, 1, choicesSheet.getLastColumn())
                .getValues()[0],
            data: choicesSheet
                .getRange(
                    2,
                    1,
                    choicesSheet.getLastRow() - 1,
                    choicesSheet.getLastColumn()
                )
                .getValues(),
        };
        setCacheData(CACHE_KEYS.CHOICES_DATA, choicesData);
    }

    const groupIndex =
        choicesData.headers.indexOf(
            user['報考群(類)代碼'].padStart(2, '0') + user['報考群(類)名稱']
        ) + 1;

    // 取得學生選擇（不使用快取，因為這是會變動的資料）
    const studentData = studentChoiceSheet
        .getRange(
            1,
            1,
            studentChoiceSheet.getLastRow(),
            studentChoiceSheet.getLastColumn()
        )
        .getValues();
    const startColumn = studentData[0].indexOf('是否參加集體報名');
    const studentRow =
        studentData[findValueRow(studentChoiceSheet, user['信箱']) - 1] || [];

    const isJoined = studentRow[startColumn]?.toString() === '是';
    const selectedChoices =
        studentRow.slice(startColumn + 1, startColumn + limitsOfChoices + 1) ||
        Array(limitsOfChoices).fill('');

    // 取得科系選項
    const departmentOptions = choicesData.data
        .map((row) => row[groupIndex - 1])
        .filter((item) => item?.toString().trim() !== '');

    Logger.log(
        'getOptionData() 返回的資料：%s',
        JSON.stringify({ isJoined, selectedChoices, departmentOptions })
    );

    return { isJoined, selectedChoices, departmentOptions };
}

function getLimitOfSchools() {
    const cachedData = getCacheData(CACHE_KEYS.LIMIT_OF_SCHOOLS);
    if (cachedData) {
        return cachedData;
    }

    const [headers, ...data] = limitOfSchoolsheet.getDataRange().getValues();
    const limits = data
        .filter((row) => row.length > 0 && row[0] !== '')
        .reduce((acc, row) => {
            acc[row[0]] = { schoolName: row[1], limitsOfSchool: row[2] };
            return acc;
        }, {});

    setCacheData(CACHE_KEYS.LIMIT_OF_SCHOOLS, limits);
    Logger.log(
        'getLimitOfSchools() 返回的限制資料：%s',
        JSON.stringify(limits)
    );
    return limits;
}

function doPost(e) {
    try {
        Logger.log('doPost 請求參數：%s', JSON.stringify(e.parameters));

        // 檢查是否已過截止時間
        const user = getUserData();
        const parameters = getParameters();
        const endTime = new Date(parameters['系統關閉時間']);
        const now = new Date();

        const timeDiff = now - endTime;
        const tolerance = 60000; // 1 分鐘的容忍時間
        if (timeDiff > tolerance) {
            Logger.log('現在時間：%s', now);
            Logger.log('截止時間：%s', endTime);
            Logger.log('doPost 已過截止時間 1 分鐘，不再更新資料！');
            return ContentService.createTextOutput('志願調查已結束');
        }

        const joinedParam = e.parameters.isJoinedInput?.[0] || '否';
        const isJoined = joinedParam === '是';
        let departmentChoices = [];
        if (isJoined) {
            departmentChoices = [
                e.parameters.departmentChoices_1?.[0] || '',
                e.parameters.departmentChoices_2?.[0] || '',
                e.parameters.departmentChoices_3?.[0] || '',
                e.parameters.departmentChoices_4?.[0] || '',
                e.parameters.departmentChoices_5?.[0] || '',
                e.parameters.departmentChoices_6?.[0] || '',
            ].sort((a, b) => {
                if (a === '' && b === '') return 0;
                if (a === '') return 1;
                if (b === '') return -1;
                return Number(a) - Number(b);
            });
        }
        const row = findValueRow(
            studentChoiceSheet,
            Session.getActiveUser().getEmail()
        );
        if (!isJoined) {
            updateSpecificRow(row, [joinedParam, '', '', '', '', '', '']);
        } else {
            updateSpecificRow(row, [joinedParam, ...departmentChoices]);
        }

        const template = HtmlService.createTemplateFromFile('output');
        template.loginEmail = Session.getActiveUser().getEmail();
        template.serviceUrl = getServiceUrl();
        template.user = user;
        template.parameters = parameters;
        template.notifications = getNotifications(parameters);
        template.limitOfSchools = getLimitOfSchools();
        template.isJoined = getOptionData(user).isJoined;
        template.selectedChoices = getOptionData(user).selectedChoices;
        template.departmentOptions = getOptionData(user).departmentOptions;

        return template.evaluate().setTitle('四技二專甄選入學志願調查系統');
    } catch (err) {
        Logger.log('doPost 發生錯誤：%s\n%s', err.message, err.stack);
        return ContentService.createTextOutput(
            '系統錯誤，請稍後再試。<br><pre>' + err.message + '</pre>'
        );
    }
}

function exportCsv() {
    const parameters = getParameters();
    const now = new Date();
    const nowString = Utilities.formatDate(
        now,
        'Asia/Taipei',
        'yyyy-MM-dd_HHmm'
    );

    const [headers, ...data] = forImportSheet.getDataRange().getValues();

    // 將所有資料轉換成文字格式，並過濾掉完全空白的列
    const processedData = data
        .filter((row) =>
            row.some(
                (cell) =>
                    cell !== null &&
                    cell !== undefined &&
                    cell.toString().trim() !== ''
            )
        )
        .map((row) =>
            row.map((cell) => {
                // 如果是空值，回傳空字串
                if (cell === null || cell === undefined) return '';
                // 將所有值轉換成字串
                return String(cell);
            })
        );

    // 如果沒有有效的資料列，顯示錯誤訊息
    if (processedData.length === 0) {
        const ui = SpreadsheetApp.getUi();
        ui.alert('錯誤', '沒有可匯出的資料，請確認資料內容。', ui.ButtonSet.OK);
        return null;
    }

    const csvContent = [
        headers.map(String).join(','),
        ...processedData.map((row) => row.join(',')),
    ].join('\n');

    // 取得試算表所在的資料夾
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    const parentFolder = spreadsheetFile.getParents().hasNext()
        ? spreadsheetFile.getParents().next()
        : DriveApp.getRootFolder();

    // 在相同資料夾中建立 CSV 檔案
    const fileName =
        parameters['報名學校代碼'] + 'StudQuota_' + nowString + '.csv';
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    const file = parentFolder.createFile(blob);

    // 取得下載連結
    const fileUrl = file.getDownloadUrl();
    Logger.log('CSV 檔案已建立在與試算表相同的資料夾中：%s', fileUrl);

    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(
        `
            <div style="padding: 10px; font-family: Arial, sans-serif;">
                <p>CSV 檔案已建立完成！</p>
                <p><a href="${fileUrl}" target="_blank" download>點此下載檔案</a></p>
                <p style="color: #666; font-size: 0.9em;">檔案已儲存在與試算表相同的資料夾中</p>
            </div>
        `
    )
        .setWidth(300)
        .setHeight(150);

    ui.showModalDialog(htmlOutput, '檔案匯出完成');
    return fileUrl;
}
