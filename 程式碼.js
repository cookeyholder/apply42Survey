const ss = SpreadsheetApp.getActiveSpreadsheet();
const configSheet = ss.getSheetByName('參數設定');
const examDataSheet = ss.getSheetByName('統測報名資料');
const choicesSheet = ss.getSheetByName('志願選項');
const studentChoiceSheet = ss.getSheetByName('考生志願列表');
const limitOfSchoolsheet = ss.getSheetByName('可報名之系科組學程數');
const limitsOfChoices = 6; // 最多可填的志願數量

/**
 * @description 建立自訂功能表「志願調查系統」
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('志願調查系統')
        .addItem('開啟表單', 'startForm')
        .addToUi();
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
        if (!user) {
            return HtmlService.createHtmlOutput(
                '請先登入學校的信箱帳號，並使用 Chrome 瀏覽器。'
            );
        } else {
            const template = HtmlService.createTemplateFromFile('index');
            template.parameters = getParameters();
            template.serviceUrl = getServiceUrl();
            template.loginEmail = Session.getActiveUser().getEmail();
            template.user = user;
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const examDataSheet = ss.getSheetByName('統測報名資料');
    if (!examDataSheet) {
        Logger.log('考生報名資料表不存在');
        return null;
    }
    const loginEmail = Session.getActiveUser().getEmail();
    if (!loginEmail) {
        Logger.log('User not logged in');
        return null;
    }

    const headers = examDataSheet
        .getRange(1, 1, 1, examDataSheet.getLastColumn())
        .getValues()[0];

    const userData = examDataSheet
        .getRange(findValueRow(examDataSheet, loginEmail), 1, 1, headers.length)
        .getValues()[0];

    const studentIdIndex = headers.indexOf('學號');
    const classIndex = headers.indexOf('班級名稱');
    const studentNameIndex = headers.indexOf('考生姓名');
    const groupIndex = headers.indexOf('報考群(類)名稱');

    if (
        classIndex < 0 ||
        studentNameIndex < 0 ||
        groupIndex < 0 ||
        studentIdIndex < 0
    ) {
        Logger.log('Header not found');
        return null;
    }

    if (userData) {
        const user = headers.reduce((acc, key, idx) => {
            acc[key] = userData[idx];
            return acc;
        }, {});
        Logger.log('getUserData() 取得使用者資料：%s', JSON.stringify(user));
        return user;
    }
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
    const headers = choicesSheet
        .getRange(1, 1, 1, choicesSheet.getLastColumn())
        .getValues()[0];

    const groupIndex =
        headers.indexOf(
            user['報考群(類)代碼'].padStart(2, '0') + user['報考群(類)名稱']
        ) + 1;

    const startColumn =
        studentChoiceSheet
            .getRange(1, 1, 1, studentChoiceSheet.getLastColumn())
            .getValues()[0]
            .indexOf('是否參加集體報名') + 1;

    let [isJoined, ...selectedChoices] = studentChoiceSheet
        .getRange(
            findValueRow(studentChoiceSheet, user['信箱']),
            startColumn,
            1,
            limitsOfChoices + 1
        )
        .getValues()[0];
    isJoined = isJoined.toString() === '是' ? true : false;

    const departmentOptions = choicesSheet
        .getRange(2, groupIndex, choicesSheet.getLastRow(), 1)
        .getValues()
        .map((row) => row[0])
        .filter((item) => item.toString().trim() !== '');

    Logger.log(
        'getOptionData() 返回的資料：%s',
        JSON.stringify({ isJoined, selectedChoices, departmentOptions })
    );

    return { isJoined, selectedChoices, departmentOptions };
}

function getLimitOfSchools() {
    const [headers, ...data] = limitOfSchoolsheet.getDataRange().getValues();
    const limits = data
        .filter((row) => row.length > 0 && row[0] !== '')
        .reduce((acc, row) => {
            acc[row[0]] = { schoolName: row[1], limitsOfSchool: row[2] };
            return acc;
        }, {});

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

        return ContentService.createTextOutput('成功更新志願選擇資料。');
    } catch (err) {
        Logger.log('doPost 發生錯誤：%s\n%s', err.message, err.stack);
        return ContentService.createTextOutput(
            '系統錯誤，請稍後再試。<br><pre>' + err.message + '</pre>'
        );
    }
}
