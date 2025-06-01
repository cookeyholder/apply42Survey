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
 * @description 依目前登入電子郵件從統測報名資料取得使用者資料
 * @returns {Object<string, any>|null} 使用者資料或 null
 */
function getUserData() {
    if (!examDataSheet || !mentorSheet) {
        Logger.log('統測報名資料或導師名單不存在');
        return null;
    }
    const loginEmail = Session.getActiveUser().getEmail();
    if (!loginEmail) {
        Logger.log('User not logged in');
        return null;
    }

    // 檢查快取中的考生資料
    const cachedExamData = getCacheData(CACHE_KEYS.EXAM_DATA);
    if (findValueRow(examDataSheet, loginEmail)) {
        examData = {
            headers: examDataSheet
                .getRange(1, 1, 1, examDataSheet.getLastColumn())
                .getValues()[0],
            data: examDataSheet.getDataRange().getValues().slice(1),
        };
    } else if (findValueRow(mentorSheet, loginEmail)) {
        examData = {
            headers: mentorSheet
                .getRange(1, 1, 1, mentorSheet.getLastColumn())
                .getValues()[0],
            data: mentorSheet.getDataRange().getValues().slice(1),
        };
    } else {
        examData = { headers: [], data: [] };
    }

    if (!cachedExamData) {
        setCacheData(CACHE_KEYS.EXAM_DATA, examData);
    }

    const { headers, data } = examData;
    Logger.log('Headers: %s', headers);
    Logger.log('Data: %s', data);

    // 尋找使用者資料
    const userRow = data.filter((row) => row[0].toString() === loginEmail)[0];
    Logger.log('userRow: %s', userRow);

    if (!userRow) {
        Logger.log('找不到使用者資料，信箱：%s', loginEmail);
        return null;
    }

    const user = headers.reduce((acc, key, idx) => {
        acc[key] = userRow[idx];
        return acc;
    }, {});

    Logger.log('getUserData() 取得使用者資料：%s', JSON.stringify(user));
    return user;
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
