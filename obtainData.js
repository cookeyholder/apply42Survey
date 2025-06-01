// 新增常數用於資料驗證
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const REQUIRED_EXAM_HEADERS = [
    '信箱',
    '統一入學測驗報名序號',
    '班級名稱',
    '考生姓名',
];
const REQUIRED_STUDENT_HEADERS = ['信箱', '是否參加集體報名'];

/**
 * @description 取得執行個體 URL
 * @returns {string} serviceUrl
 */
function getServiceUrl() {
    try {
        return ScriptApp.getService().getUrl();
    } catch (error) {
        Logger.log('取得服務 URL 失敗：%s', error.message);
        return '';
    }
}

/**
 * @description 從參數設定工作表取得所有參數
 * @returns {Object<string, any>} 參數鍵值對
 */
function getParameters() {
    try {
        if (!configSheet) {
            throw new Error('參數設定工作表不存在');
        }

        const dataRange = configSheet.getDataRange();
        if (dataRange.getNumRows() > MAX_SHEET_ROWS) {
            throw new Error('參數設定資料過大');
        }

        const data = dataRange.getValues();
        const parameters = {};

        for (const row of data) {
            if (row.length >= 2 && row[0] && typeof row[0] === 'string') {
                const key = row[0].toString().trim();
                const value = row[1];

                if (key.length > 0 && key.length <= 100) {
                    parameters[key] = value;
                }
            }
        }

        // 驗證必要參數
        const requiredParams = ['系統名稱', '系統關閉時間'];
        for (const param of requiredParams) {
            if (!parameters[param]) {
                Logger.log('缺少必要參數：%s', param);
            }
        }

        return parameters;
    } catch (error) {
        Logger.log('取得參數失敗：%s', error.message);
        return {};
    }
}

/**
 * @description 驗證使用者電子郵件格式
 * @param {string} email - 電子郵件地址
 * @returns {boolean} 是否為有效格式
 */
function isValidEmail(email) {
    return typeof email === 'string' && EMAIL_REGEX.test(email);
}

/**
 * @description 安全地取得工作表資料
 * @param {Sheet} sheet - 工作表物件
 * @param {Array<string>} requiredHeaders - 必要的標頭
 * @returns {Object|null} 包含 headers 和 data 的物件
 */
function getSheetDataSafely(sheet, requiredHeaders = []) {
    try {
        if (!sheet) {
            return null;
        }

        const dataRange = sheet.getDataRange();
        if (
            dataRange.getNumRows() === 0 ||
            dataRange.getNumRows() > MAX_SHEET_ROWS
        ) {
            Logger.log(
                '工作表 %s 資料列數異常：%d',
                sheet.getName(),
                dataRange.getNumRows()
            );
            return null;
        }

        const allData = dataRange.getValues();
        const headers = allData[0] || [];
        const data = allData.slice(1);

        // 驗證必要標頭
        for (const requiredHeader of requiredHeaders) {
            if (!headers.includes(requiredHeader)) {
                Logger.log(
                    '工作表 %s 缺少必要標頭：%s',
                    sheet.getName(),
                    requiredHeader
                );
                return null;
            }
        }

        return { headers, data };
    } catch (error) {
        Logger.log(
            '取得工作表資料失敗 (%s)：%s',
            sheet?.getName() || 'unknown',
            error.message
        );
        return null;
    }
}

/**
 * @description 依目前登入電子郵件從統測報名資料取得使用者資料（安全版本）
 * @returns {Object<string, any>|null} 使用者資料或 null
 */
function getUserData() {
    try {
        if (!examDataSheet || !mentorSheet) {
            Logger.log('統測報名資料或導師名單工作表不存在');
            return null;
        }

        const loginEmail = Session.getActiveUser().getEmail();
        if (!loginEmail || !isValidEmail(loginEmail)) {
            Logger.log('使用者未登入或電子郵件格式無效');
            return null;
        }

        // 嘗試從快取獲取使用者資料
        const cacheKey = `user_${loginEmail}`;
        let userData = getCacheData(cacheKey);
        if (userData) {
            Logger.log('從快取取得使用者資料：%s', loginEmail);
            return userData;
        }

        // 檢查是否為學生（在統測報名資料中）
        let userRow = null;
        let headers = null;
        let isStudent = false;

        const examRow = findValueRow(examDataSheet, loginEmail);
        if (examRow && examRow > 0) {
            headers = examDataSheet
                .getRange(1, 1, 1, examDataSheet.getLastColumn())
                .getValues()[0];
            const dataRow = examDataSheet
                .getRange(examRow, 1, 1, examDataSheet.getLastColumn())
                .getValues()[0];
            userRow = dataRow;
            isStudent = true;
            Logger.log('在統測報名資料中找到使用者：%s', loginEmail);
        } else {
            // 檢查是否為導師
            const mentorRow = findValueRow(mentorSheet, loginEmail);
            if (mentorRow && mentorRow > 0) {
                headers = mentorSheet
                    .getRange(1, 1, 1, mentorSheet.getLastColumn())
                    .getValues()[0];
                const dataRow = mentorSheet
                    .getRange(mentorRow, 1, 1, mentorSheet.getLastColumn())
                    .getValues()[0];
                userRow = dataRow;
                isStudent = false;
                Logger.log('在導師名單中找到使用者：%s', loginEmail);
            }
        }

        if (!userRow || !headers) {
            Logger.log('找不到使用者資料，信箱：%s', loginEmail);
            return null;
        }

        // 建立使用者物件
        userData = headers.reduce((acc, key, idx) => {
            if (key && idx < userRow.length) {
                acc[String(key)] = userRow[idx] !== null ? userRow[idx] : '';
            }
            return acc;
        }, {});

        // 標記使用者類型
        userData._isStudent = isStudent;
        userData._userType = isStudent ? 'student' : 'mentor';

        // 快取使用者資料（較短的快取時間）
        setCacheData(cacheKey, userData, 1800); // 30 分鐘

        Logger.log(
            'getUserData() 成功取得使用者資料：%s（類型：%s）',
            loginEmail,
            userData._userType
        );
        return userData;
    } catch (error) {
        Logger.log('getUserData() 發生錯誤：%s', error.message);
        return null;
    }
}

/**
 * @description 在工作表資料中尋找使用者
 * @param {string} email - 使用者電子郵件
 * @param {Object} sheetData - 工作表資料物件
 * @returns {Object|null} 使用者資料或 null
 */
function findUserInSheetData(email, sheetData) {
    try {
        const { headers, data } = sheetData;
        const emailIndex = headers.indexOf('信箱');

        if (emailIndex === -1) {
            return null;
        }

        const userRow = data.find(
            (row) =>
                row[emailIndex] &&
                row[emailIndex].toString().toLowerCase() === email.toLowerCase()
        );

        if (!userRow) {
            return null;
        }

        // 建立使用者資料物件
        const userData = {};
        headers.forEach((header, index) => {
            if (header && typeof header === 'string') {
                userData[header] = userRow[index] || '';
            }
        });

        return userData;
    } catch (error) {
        Logger.log('在工作表資料中尋找使用者時發生錯誤：%s', error.message);
        return null;
    }
}

/**
 * @description 取得通知訊息列表（安全版本）
 * @param {Object} parameters - 系統參數
 * @returns {string} HTML 格式的通知列表
 */
function getNotifications(parameters) {
    if (!parameters || typeof parameters !== 'object') {
        Logger.log('getNotifications: 參數無效');
        return '';
    }

    try {
        const notifications = [];
        const descriptionKeys = Object.keys(parameters)
            .filter((key) => key && key.startsWith('說明'))
            .sort(); // 排序確保順序一致

        descriptionKeys.forEach((key) => {
            const description = parameters[key];
            if (description && typeof description === 'string') {
                // 清理 HTML 內容以防 XSS
                const cleanDescription = sanitizeHtml(description);
                notifications.push(`<li>${cleanDescription}</li>`);
            }
        });

        return notifications.join('');
    } catch (error) {
        Logger.log('getNotifications() 發生錯誤：%s', error.message);
        return '';
    }
}

/**
 * @description 取得使用者的參加狀態、已選擇志願及可選擇志願列表（安全版本）
 * @param {Object<string, any>} [user] - 使用者資訊，若為 null 則自動取得
 * @returns {{isJoined: boolean, selectedChoices: any[], departmentOptions: string[]}}
 */
function getOptionData(user = null) {
    try {
        if (!user) {
            user = getUserData();
        }

        if (!user || !user['報考群(類)代碼'] || !user['報考群(類)名稱']) {
            Logger.log('getOptionData: 使用者資料不完整');
            return {
                isJoined: false,
                selectedChoices: [],
                departmentOptions: [],
            };
        }

        // 取得志願選項資料（使用快取）
        let choicesData = getCacheData(CACHE_KEYS.CHOICES_DATA);
        if (!choicesData && choicesSheet) {
            try {
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
            } catch (error) {
                Logger.log('讀取志願選項資料時發生錯誤：%s', error.message);
                return {
                    isJoined: false,
                    selectedChoices: [],
                    departmentOptions: [],
                };
            }
        }

        if (!choicesData) {
            Logger.log('getOptionData: 志願選項資料不可用');
            return {
                isJoined: false,
                selectedChoices: [],
                departmentOptions: [],
            };
        }

        // 尋找對應的群類欄位
        const groupCode = String(user['報考群(類)代碼']).padStart(2, '0');
        const groupName = String(user['報考群(類)名稱']);
        const targetColumn = groupCode + groupName;

        const groupIndex = choicesData.headers.indexOf(targetColumn);
        if (groupIndex === -1) {
            Logger.log('找不到對應的群類欄位：%s', targetColumn);
            return {
                isJoined: false,
                selectedChoices: [],
                departmentOptions: [],
            };
        }

        // 取得學生選擇資料
        let studentData = null;
        const userEmail = user['信箱'] || Session.getActiveUser().getEmail();

        if (studentChoiceSheet) {
            try {
                studentData = studentChoiceSheet
                    .getRange(
                        1,
                        1,
                        studentChoiceSheet.getLastRow(),
                        studentChoiceSheet.getLastColumn()
                    )
                    .getValues();
            } catch (error) {
                Logger.log('讀取學生選擇資料時發生錯誤：%s', error.message);
                return {
                    isJoined: false,
                    selectedChoices: [],
                    departmentOptions: [],
                };
            }
        }

        if (!studentData || studentData.length < 2) {
            Logger.log('學生選擇資料不可用');
            return {
                isJoined: false,
                selectedChoices: [],
                departmentOptions: [],
            };
        }

        const studentHeaders = studentData[0];
        const startColumnIndex = studentHeaders.indexOf('是否參加集體報名');

        if (startColumnIndex === -1) {
            Logger.log('找不到「是否參加集體報名」欄位');
            return {
                isJoined: false,
                selectedChoices: [],
                departmentOptions: [],
            };
        }

        // 尋找學生資料列
        const studentRowIndex = studentData.findIndex(
            (row, index) => index > 0 && row[0] === userEmail
        );

        let isJoined = false;
        let selectedChoices = Array(limitsOfChoices).fill('');

        if (studentRowIndex > 0) {
            const studentRow = studentData[studentRowIndex];
            isJoined = String(studentRow[startColumnIndex]).trim() === '是';

            // 取得已選擇的志願
            for (let i = 0; i < limitsOfChoices; i++) {
                const choiceIndex = startColumnIndex + 1 + i;
                if (
                    choiceIndex < studentRow.length &&
                    studentRow[choiceIndex]
                ) {
                    selectedChoices[i] = String(studentRow[choiceIndex]).trim();
                }
            }
        }

        // 取得科系選項
        const departmentOptions = choicesData.data
            .map((row) => row[groupIndex])
            .filter((item) => item && String(item).trim() !== '')
            .map((item) => String(item));

        const result = { isJoined, selectedChoices, departmentOptions };
        Logger.log(
            'getOptionData() 返回資料：%s',
            JSON.stringify({
                isJoined: result.isJoined,
                selectedChoicesCount: result.selectedChoices.filter((c) => c)
                    .length,
                departmentOptionsCount: result.departmentOptions.length,
            })
        );

        return result;
    } catch (error) {
        Logger.log('getOptionData() 發生錯誤：%s', error.message);
        return { isJoined: false, selectedChoices: [], departmentOptions: [] };
    }
}

/**
 * @description 取得學校志願數限制（安全版本）
 * @returns {Object} 學校限制資料
 */
function getLimitOfSchools() {
    try {
        const cachedData = getCacheData(CACHE_KEYS.LIMIT_OF_SCHOOLS);
        if (cachedData) {
            return cachedData;
        }

        if (!limitOfSchoolsSheet) {
            Logger.log('學校限制工作表不存在');
            return {};
        }

        const sheetData = getSheetDataSafely(limitOfSchoolsSheet);
        if (!sheetData) {
            Logger.log('無法取得學校限制資料');
            return {};
        }

        const { headers, data } = sheetData;
        const limitData = {};

        data.forEach((row) => {
            if (row.length >= 3 && row[0] && row[1] && row[2]) {
                const schoolCode = String(row[0]).trim();
                const schoolName = String(row[1]).trim();
                const limitsOfSchool = parseInt(row[2]);

                if (
                    schoolCode &&
                    !isNaN(limitsOfSchool) &&
                    limitsOfSchool > 0
                ) {
                    limitData[schoolCode] = {
                        schoolName: schoolName,
                        limitsOfSchool: limitsOfSchool,
                    };
                }
            }
        });

        // 快取資料
        setCacheData(CACHE_KEYS.LIMIT_OF_SCHOOLS, limitData);

        Logger.log(
            'getLimitOfSchools() 成功取得 %d 筆學校限制資料',
            Object.keys(limitData).length
        );
        return limitData;
    } catch (error) {
        Logger.log('getLimitOfSchools() 發生錯誤：%s', error.message);
        return {};
    }
}

/**
 * @description 清理 HTML 內容以防 XSS 攻擊
 * @param {string} html - 要清理的 HTML 字串
 * @returns {string} 清理後的 HTML
 */
function sanitizeHtml(html) {
    if (!html || typeof html !== 'string') {
        return '';
    }

    try {
        // 移除危險的標籤和屬性
        return html
            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
            .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
            .replace(/javascript:/gi, '')
            .replace(/on\w+\s*=/gi, '')
            .trim();
    } catch (error) {
        Logger.log('sanitizeHtml() 發生錯誤：%s', error.message);
        return '';
    }
}
