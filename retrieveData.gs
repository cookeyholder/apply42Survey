const cache = CacheService.getScriptCache();
// 新增常數用於資料驗證
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const REQUIRED_EXAM_HEADERS = [
  "信箱",
  "統一入學測驗報名序號",
  "班級名稱",
  "考生姓名",
];
const REQUIRED_STUDENT_HEADERS = ["信箱", "是否參加集體報名"];

/**
 * @description 取得執行個體 URL
 * @returns {string} serviceUrl
 */
function getServiceUrl() {
  try {
    const effectiveUserEmail = Session.getEffectiveUser().getEmail();
    const effectiveUserDomain = effectiveUserEmail.split("@")[1];
    const url = ScriptApp.getService().getUrl();
    const regex = /\/s\/(.+?)\//;
    const serviceId = url.match(regex)[1];
    const mode = url.split("/").pop();
    return (
      "https://script.google.com/a/macros/" +
      effectiveUserDomain +
      "/s/" +
      serviceId +
      "/" +
      mode
    );
  } catch (error) {
    Logger.log("取得服務 URL 失敗：%s", error.message);
    return "";
  }
}

/**
 * @description 從參數設定工作表取得所有參數
 * @returns {Object<string, any>} 參數鍵值對
 */
function getConfigs() {
  try {
    if (!configSheet) {
      throw new Error("(getConfigs)參數設定工作表不存在");
    }

    const dataRange = configSheet.getDataRange();
    if (dataRange.getNumRows() > MAX_SHEET_ROWS) {
      throw new Error("(getConfigs)參數設定資料過大");
    }

    const data = dataRange.getValues();
    const configs = {};

    for (const row of data) {
      if (row.length >= 2 && row[0] && typeof row[0] === "string") {
        const key = row[0].toString().trim();
        const value = row[1];

        if (key.length > 0 && key.length <= 100) {
          configs[key] = value;
        }
      }
    }

    // 驗證必要參數
    const requiredParams = ["系統名稱", "系統關閉時間"];
    for (const param of requiredParams) {
      if (!configs[param]) {
        Logger.log("(getConfigs)缺少必要參數：%s", param);
      }
    }

    return configs;
  } catch (error) {
    Logger.log("(getConfigs)取得參數失敗：%s", error.message);
    return {};
  }
}

/**
 * @description 驗證使用者電子郵件格式
 * @param {string} email - 電子郵件地址
 * @returns {boolean} 是否為有效格式
 */
function isValidEmail(email) {
  return typeof email === "string" && EMAIL_REGEX.test(email);
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
        "(getSheetDataSafely)工作表 %s 資料列數異常：%d",
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
          "(getSheetDataSafely)工作表 %s 缺少必要標頭：%s",
          sheet.getName(),
          requiredHeader
        );
        return null;
      }
    }

    return { headers, data };
  } catch (error) {
    Logger.log(
      "(getSheetDataSafely)取得工作表資料失敗 (%s)：%s",
      sheet?.getName() || "unknown",
      error.message
    );
    return null;
  }
}

/**
 * @description 取得指定標頭在工作表中的索引（1-based）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 工作表物件
 * @param {string} headerName 標頭名稱
 * @returns {number} 標頭的欄位索引，如果找不到則回傳 -1
 */
function getHeaderIndex(sheet, headerName) {
  if (!sheet || !headerName) {
    Logger.log("(getHeaderIndex)無效的參數");
    return -1;
  }
  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const index = headers.indexOf(headerName);
    return index !== -1 ? index + 1 : -1; // Apps Script 的欄位索引是 1-based
  } catch (error) {
    Logger.log("(getHeaderIndex)執行時發生錯誤: %s", error.message);
    return -1;
  }
}

/**
 * @description 從快取中取得使用者資料
 * @param {string} email - 使用者電子郵件
 * @returns {Object<string, any>|null} 快取的使用者資料或 null
 */
function getUserFromCache(email) {
  if (!email) return null;

  try {
    const validEmailCacheKey = getSafeKeyFromEmail(email);
    const cacheKey = CACHE_KEYS.USER_DATA_PREFIX + validEmailCacheKey;

    if (!isValidCacheKey(cacheKey)) {
      Logger.log("(getUserFromCache)生成的快取鍵值無效：%s", cacheKey);
      return null;
    }

    const cached = getCacheData(cacheKey);

    if (cached) {
      Logger.log("(getUserFromCache)從快取取得使用者資料：%s", email);
      return cached;
    }
  } catch (error) {
    Logger.log(
      "(getUserFromCache)從快取讀取使用者資料時發生錯誤：%s",
      error.message
    );
  }

  return null;
}

/**
 * @description 在統測報名資料表中搜尋使用者
 * @param {string} email - 使用者電子郵件
 * @returns {Object|null} 包含使用者列號、工作表、欄位索引和使用者類型的物件，或 null
 */
function findStudentUser(email) {
  if (!examDataSheet) {
    Logger.log(`(findStudentUser)examDataSheet 不存在`);
    return null;
  }

  const userRow = findValueRow(email, examDataSheet);

  if (userRow && userRow > 0) {
    Logger.log(
      "(findStudentUser)找到學生使用者，信箱：%s，行號：%d",
      email,
      userRow
    );
    return {
      userRow,
      targetSheet: examDataSheet,
      idColumnIndex: getHeaderIndex(examDataSheet, "信箱"),
      userType: "學生",
    };
  }

  return null;
}

/**
 * @description 在老師資料表中搜尋使用者
 * @param {string} email - 使用者電子郵件
 * @returns {Object|null} 包含使用者列號、工作表、欄位索引和使用者類型的物件，或 null
 */
function findTeacherUser(email) {
  if (!teacherSheet) {
    Logger.log(`(findTeacherUser)teacherSheet 不存在`);
    return null;
  }

  const userRow = findValueRow(email, teacherSheet);

  if (userRow && userRow > 0) {
    Logger.log(
      "(findTeacherUser)找到老師使用者，信箱：%s，行號：%d",
      email,
      userRow
    );
    return {
      userRow,
      targetSheet: teacherSheet,
      idColumnIndex: getHeaderIndex(teacherSheet, "信箱"),
      userType: "老師",
    };
  }

  return null;
}

/**
 * @description 從工作表行數據建立使用者資料物件
 * @param {Sheet} targetSheet - 目標工作表
 * @param {number} userRow - 使用者資料所在的列號
 * @param {string} userType - 使用者類型
 * @returns {Object<string, any>|null} 使用者資料物件或 null
 */
function buildUserDataObject(targetSheet, userRow, userType) {
  try {
    const headers = targetSheet
      .getRange(1, 1, 1, targetSheet.getLastColumn())
      .getValues()[0];
    const dataRow = targetSheet
      .getRange(userRow, 1, 1, targetSheet.getLastColumn())
      .getValues()[0];

    const userData = headers.reduce((acc, key, idx) => {
      if (key && idx < dataRow.length) {
        acc[String(key)] = dataRow[idx] !== null ? dataRow[idx] : "";
      }
      return acc;
    }, {});

    userData.userType = userType;
    Logger.log(
      "(buildUserDataObject)成功建立使用者資料物件：%s",
      JSON.stringify(userData)
    );
    return userData;
  } catch (error) {
    Logger.log(
      "(buildUserDataObject)建立使用者資料物件時發生錯誤：%s",
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
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  const cached = getUserFromCache(email);
  if (cached) return cached;

  // 從工作表取得資料並快取
  let userData = findStudentUser(email);

  // 若未找到學生，嘗試在老師資料表搜尋
  if (!userData) {
    Logger.log(
      `(getUserData)使用者 ${email} 在統測報名資料表中未找到，嘗試在老師資料表搜尋`
    );
    userData = findTeacherUser(email);
  } else {
    Logger.log(`(getUserData)使用者 ${email} 在統測報名資料表中找到`);
  }

  // 若均未找到，回傳 null
  if (!userData) {
    Logger.log(`(getUserData)使用者 ${email} 在老師及統測報名資料表中均未找到`);
    return null;
  }

  // 建立使用者資料物件
  const { userRow, targetSheet, userType } = userData;
  const userDataObject = buildUserDataObject(targetSheet, userRow, userType);

  if (userDataObject) {
    // 快取使用者資料（24 小時）
    const validEmailCacheKey = getSafeKeyFromEmail(email);

    // 將使用者添加到快取索引，方便日後清除
    addUserToIndex(email);

    // 存儲使用者資料
    setCacheData(
      CACHE_KEYS.USER_DATA_PREFIX + validEmailCacheKey,
      userDataObject,
      86400
    );
    Logger.log("(getUserData)成功取得並快取使用者資料：%s", email);
  }

  return userDataObject;
}

/**
 * @description 在工作表資料中尋找使用者
 * @param {string} email - 使用者電子郵件
 * @param {Object} target - 工作表資料物件
 * @returns {Object|null} 使用者資料或 null
 */
function findUserInSheet(email, target) {
  try {
    const userRow = findValueRow(email, target);
    if (!userRow || userRow === 0) {
      Logger.log("(findUserInSheet)找不到使用者資料，信箱：%s", email);
      return null;
    } else {
      Logger.log(
        "(findUserInSheet)找到使用者資料，信箱：%s，行號：%d",
        email,
        userRow
      );
    }

    const headers = target
      .getRange(1, 1, 1, target.getLastColumn())
      .getValues()[0];
    const dataRow = target
      .getRange(userRow, 1, 1, target.getLastColumn())
      .getValues()[0];

    const userData = headers.reduce((acc, key, idx) => {
      if (key && idx < dataRow.length) {
        acc[String(key)] = dataRow[idx] !== null ? dataRow[idx] : "";
      }
      return acc;
    }, {});

    return userData ? userData : null;
  } catch (error) {
    Logger.log(
      "(findUserInSheet)在工作表資料中尋找使用者時發生錯誤：%s",
      error.message
    );
    return null;
  }
}

/**
 * @description 取得通知訊息列表（安全版本）
 * @param {Object} configs - 系統參數
 * @returns {string} HTML 格式的通知列表
 */
function getNotifications(configs) {
  if (!configs || typeof configs !== "object") {
    Logger.log("(getNotifications)參數無效");
    return "";
  }

  try {
    const notifications = [];
    const descriptionKeys = Object.keys(configs)
      .filter((key) => key && key.startsWith("說明"))
      .sort(); // 排序確保順序一致

    descriptionKeys.forEach((key) => {
      const description = configs[key];
      if (description && typeof description === "string") {
        // 清理 HTML 內容以防 XSS
        const cleanDescription = sanitizeHtml(description);
        notifications.push(`<li>${cleanDescription}</li>`);
      }
    });

    Logger.log(
      "(getNotifications)成功取得通知訊息：%s",
      notifications.join("")
    );

    return notifications.join("");
  } catch (error) {
    Logger.log("getNotifications() 發生錯誤：%s", error.message);
    return "";
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

    if (!user || !user["報考群(類)代碼"] || !user["報考群(類)名稱"]) {
      Logger.log("getOptionData: 使用者資料不完整");
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
        Logger.log(
          "(getOptionData)讀取志願選項資料時發生錯誤：%s",
          error.message
        );
        return {
          isJoined: false,
          selectedChoices: [],
          departmentOptions: [],
        };
      }
    }

    if (!choicesData) {
      Logger.log("(getOptionData)志願選項資料不可用");
      return {
        isJoined: false,
        selectedChoices: [],
        departmentOptions: [],
      };
    }

    // 尋找對應的群類欄位
    const groupCode = String(user["報考群(類)代碼"]).padStart(2, "0");
    const groupName = String(user["報考群(類)名稱"]);
    const targetColumn = groupCode + groupName;

    const groupIndex = choicesData.headers.indexOf(targetColumn);
    if (groupIndex === -1) {
      Logger.log("(getOptionData)找不到對應的群類欄位：%s", targetColumn);
      return {
        isJoined: false,
        selectedChoices: [],
        departmentOptions: [],
      };
    }

    // 取得學生選擇資料
    let studentData = null;
    const userEmail = user["信箱"] || Session.getActiveUser().getEmail();

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
        Logger.log(
          "(getOptionData)讀取學生選擇資料時發生錯誤：%s",
          error.message
        );
        return {
          isJoined: false,
          selectedChoices: [],
          departmentOptions: [],
        };
      }
    }

    if (!studentData || studentData.length < 2) {
      Logger.log("(getOptionData)學生選擇資料不可用");
      return {
        isJoined: false,
        selectedChoices: [],
        departmentOptions: [],
      };
    }

    const studentHeaders = studentData[0];
    const startColumnIndex = studentHeaders.indexOf("是否參加集體報名");

    if (startColumnIndex === -1) {
      Logger.log("找不到「是否參加集體報名」欄位");
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
    let selectedChoices = Array(limitOfChoices).fill("");

    if (studentRowIndex > 0) {
      const studentRow = studentData[studentRowIndex];
      isJoined = String(studentRow[startColumnIndex]).trim() === "是";

      // 取得已選擇的志願
      for (let i = 0; i < limitOfChoices; i++) {
        const choiceIndex = startColumnIndex + 1 + i;
        if (choiceIndex < studentRow.length && studentRow[choiceIndex]) {
          selectedChoices[i] = String(studentRow[choiceIndex]).trim();
        }
      }
    }

    // 取得科系選項
    const departmentOptions = choicesData.data
      .map((row) => row[groupIndex])
      .filter((item) => item && String(item).trim() !== "")
      .map((item) => String(item));

    const result = { isJoined, selectedChoices, departmentOptions };
    Logger.log(
      "(getOptionData)返回資料：%s",
      JSON.stringify({
        是否參加集體報名: result.isJoined,
        選擇志願數: result.selectedChoices.filter((c) => c).length,
        志願選項數: result.departmentOptions.length,
      })
    );

    return result;
  } catch (error) {
    Logger.log("(getOptionData)發生錯誤：%s", error.message);
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
      Logger.log("(getLimitOfSchools)學校限制工作表不存在");
      return {};
    }

    const sheetData = getSheetDataSafely(limitOfSchoolsSheet);
    if (!sheetData) {
      Logger.log("(getLimitOfSchools)無法取得學校限制資料");
      return {};
    }

    const { headers, data } = sheetData;
    const limitData = {};

    data.forEach((row) => {
      if (row.length >= 3 && row[0] && row[1] && row[2]) {
        const schoolCode = String(row[0]).trim();
        const schoolName = String(row[1]).trim();
        const limitsOfSchool = parseInt(row[2]);

        if (schoolCode && !isNaN(limitsOfSchool) && limitsOfSchool > 0) {
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
      "(getLimitOfSchools)成功取得 %d 筆學校限制資料",
      Object.keys(limitData).length
    );
    return limitData;
  } catch (error) {
    Logger.log("(getLimitOfSchools)發生錯誤：%s", error.message);
    return {};
  }
}

/**
 * @description 清理 HTML 內容以防 XSS 攻擊
 * @param {string} html - 要清理的 HTML 字串
 * @returns {string} 清理後的 HTML
 */
function sanitizeHtml(html) {
  if (!html || typeof html !== "string") {
    return "";
  }

  try {
    // 移除危險的標籤和屬性
    return html
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
      .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, "")
      .replace(/javascript:/gi, "")
      .replace(/on\w+\s*=/gi, "")
      .trim();
  } catch (error) {
    Logger.log("sanitizeHtml() 發生錯誤：%s", error.message);
    return "";
  }
}
