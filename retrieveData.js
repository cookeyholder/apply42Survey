const cache = CacheService.getScriptCache(); // Add this line to define cache
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
    return ScriptApp.getService().getUrl();
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
      throw new Error("參數設定工作表不存在");
    }

    const dataRange = configSheet.getDataRange();
    if (dataRange.getNumRows() > MAX_SHEET_ROWS) {
      throw new Error("參數設定資料過大");
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
        Logger.log("缺少必要參數：%s", param);
      }
    }

    return configs;
  } catch (error) {
    Logger.log("取得參數失敗：%s", error.message);
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
        "工作表 %s 資料列數異常：%d",
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
          "工作表 %s 缺少必要標頭：%s",
          sheet.getName(),
          requiredHeader
        );
        return null;
      }
    }

    return { headers, data };
  } catch (error) {
    Logger.log(
      "取得工作表資料失敗 (%s)：%s",
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
    Logger.log("getHeaderIndex: 無效的參數");
    return -1;
  }
  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const index = headers.indexOf(headerName);
    return index !== -1 ? index + 1 : -1; // Apps Script 的欄位索引是 1-based
  } catch (error) {
    Logger.log("getHeaderIndex 執行時發生錯誤: %s", error.message);
    return -1;
  }
}

/**
 * @description 依目前登入電子郵件從統測報名資料取得使用者資料（安全版本）
 * @returns {Object<string, any>|null} 使用者資料或 null
 */
function getUserData() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  // 快取鍵值
  const cacheKey = CACHE_KEYS.USER_DATA_PREFIX + email;
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  let targetSheet, idColumnIndex, userType;

  // 檢查是否為導師
  if (mentorSheet) {
    // const mentorRow = findValueRow(mentorSheet, email); // Original
    const mentorRow = findValueRow(email, mentorSheet); // Changed
    if (mentorRow && mentorRow > 0) {
      targetSheet = mentorSheet;
      idColumnIndex = getHeaderIndex(targetSheet, "信箱");
      userType = "導師";
    }
  }

  // 如果不是導師，或導師表不存在，則檢查是否為學生
  if (!targetSheet && examDataSheet) {
    // const studentRow = findValueRow(examDataSheet, email); // Original
    const studentRow = findValueRow(email, examDataSheet); // Changed
    if (studentRow && studentRow > 0) {
      targetSheet = examDataSheet;
      idColumnIndex = getHeaderIndex(targetSheet, "信箱");
      userType = "學生";
    } else {
      // 如果在學生資料中也找不到，則回傳 null
      Logger.log(`使用者 ${email} 在導師及學生名單中均未找到`);
      return null;
    }
  } else if (!targetSheet) {
    // 如果兩個工作表都不存在
    Logger.log("導師名單和統測報名資料工作表均不存在");
    return null;
  }

  // const userRow = findValueRow(target, email); // Original
  // const userRow = findValueRow(email, target); // Changed
  const userRow = findValueRow(email, targetSheet); // Corrected: Use targetSheet

  if (!userRow || userRow === 0) {
    Logger.log("找不到使用者資料，信箱：%s", email);
    return null;
  }

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

  // 快取使用者資料（較長的快取時間）
  setCacheData(cacheKey, userData, 86400); // 24 小時

  Logger.log("getUserData() 成功取得使用者資料：%s", email);
  return userData;
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
      Logger.log("找不到使用者資料，信箱：%s", email);
      return null;
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
    Logger.log("在工作表資料中尋找使用者時發生錯誤：%s", error.message);
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
    Logger.log("getNotifications: 參數無效");
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
        Logger.log("讀取志願選項資料時發生錯誤：%s", error.message);
        return {
          isJoined: false,
          selectedChoices: [],
          departmentOptions: [],
        };
      }
    }

    if (!choicesData) {
      Logger.log("getOptionData: 志願選項資料不可用");
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
      Logger.log("找不到對應的群類欄位：%s", targetColumn);
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
        Logger.log("讀取學生選擇資料時發生錯誤：%s", error.message);
        return {
          isJoined: false,
          selectedChoices: [],
          departmentOptions: [],
        };
      }
    }

    if (!studentData || studentData.length < 2) {
      Logger.log("學生選擇資料不可用");
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
      "getOptionData() 返回資料：%s",
      JSON.stringify({
        isJoined: result.isJoined,
        selectedChoicesCount: result.selectedChoices.filter((c) => c).length,
        departmentOptionsCount: result.departmentOptions.length,
      })
    );

    return result;
  } catch (error) {
    Logger.log("getOptionData() 發生錯誤：%s", error.message);
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
      Logger.log("學校限制工作表不存在");
      return {};
    }

    const sheetData = getSheetDataSafely(limitOfSchoolsSheet);
    if (!sheetData) {
      Logger.log("無法取得學校限制資料");
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
      "getLimitOfSchools() 成功取得 %d 筆學校限制資料",
      Object.keys(limitData).length
    );
    return limitData;
  } catch (error) {
    Logger.log("getLimitOfSchools() 發生錯誤：%s", error.message);
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
