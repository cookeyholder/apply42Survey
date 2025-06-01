const ss = SpreadsheetApp.getActiveSpreadsheet();
const configSheet = ss.getSheetByName('參數設定');
const examDataSheet = ss.getSheetByName('統測報名資料');
const choicesSheet = ss.getSheetByName('志願選項');
const studentChoiceSheet = ss.getSheetByName('考生志願列表');
const limitOfSchoolsheet = ss.getSheetByName('可報名之系科組學程數');
const forImportSheet = ss.getSheetByName('匯入報名系統');
const mentorSheet = ss.getSheetByName('導師名單');
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

        // 如果使用者未登入或登入的不在允許名單之中
        if (!user) {
            return HtmlService.createHtmlOutput(
                '請先登入學校的信箱帳號，並使用 Chrome 瀏覽器。'
            );
        }

        // 如果使用者資料中有「統一入學測驗報名序號」，則表示他是學生
        if (user['統一入學測驗報名序號']) {
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

        // 如果使用者資料中沒有「統一入學測驗報名序號」，則表示他是導師
        if (!user['統一入學測驗報名序號']) {
            const { headers, data } = getTraineesDepartmentChoices(user);
            const template = HtmlService.createTemplateFromFile('mentorView');
            template.loginEmail = Session.getActiveUser().getEmail();
            template.serviceUrl = getServiceUrl();
            template.user = user;
            template.parameters = parameters;
            template.headers = headers;
            template.data = data;
            return template.evaluate().setTitle('導師查詢班級學生志願');
        }
    } catch (err) {
        Logger.log('doGet 發生錯誤：%s\n%s', err.message, err.stack);
        return HtmlService.createHtmlOutput(
            '系統錯誤，請稍後再試。<br><pre>' + err.message + '</pre>'
        );
    }
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

        const template = HtmlService.createTemplateFromFile('success');
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
