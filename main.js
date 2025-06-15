// 安全性和效能常數
const MAX_SHEET_ROWS = 10000;
const REQUEST_TIMEOUT = 30000; // 30 秒
const MAX_PARAMETER_LENGTH = 500;
const ALLOWED_PARAMETERS = ["系統名稱", "系統關閉時間", "報名學校代碼"];

const ss = SpreadsheetApp.getActiveSpreadsheet();
const configSheet = ss.getSheetByName("參數設定");
const examDataSheet = ss.getSheetByName("統測報名資料");
const choicesSheet = ss.getSheetByName("志願選項");
const studentChoiceSheet = ss.getSheetByName("考生志願列表");
const limitOfSchoolsSheet = ss.getSheetByName("可報名之系科組學程數");
const forImportSheet = ss.getSheetByName("匯入報名系統");
const teacherSheet = ss.getSheetByName("老師名單");
const logSheet = ss.getSheetByName("日誌");
const limitOfChoices = 6; // 最多可填的志願數量

/**
 * @description 建立自訂功能表「志願調查系統」
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("志願調查系統")
      .addItem("匯出報名用CSV", "exportCsv")
      .addItem("各志願選填人數統計", "showStatisticsPage")
      .addItem("清除快取", "clearAllCacheInternal")
      .addToUi();
    Logger.log("(onOpen)功能表建立成功");
  } catch (error) {
    Logger.log("(onOpen)建立功能表時發生錯誤：%s", error.message);
  }
}

/**
 * @description 處理 GET 請求，回傳表單頁面或錯誤訊息（安全版本）
 * @param {Object} request - 請求參數
 * @returns {HtmlOutput} HTML 輸出
 */
function doGet(request) {
  try {
    // 驗證請求參數
    if (!validateRequestParameters(request.parameters)) {
      Logger.log("請求參數驗證失敗");
      return HtmlService.createHtmlOutput(
        '<div style="padding: 20px; color: red;">無效的請求參數</div>'
      );
    }

    Logger.log("(doGet)請求參數：%s", JSON.stringify(request.parameters));

    const user = getUserData();
    Logger.log("(doGet)取得的使用者資料：%s", JSON.stringify(user));

    const configs = getConfigs();
    Logger.log("(doGet)取得的系統設定資訊：%s", JSON.stringify(configs));

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
    if (!configs || !configs["系統名稱"]) {
      Logger.log("(doGet)系統參數不完整");
      return HtmlService.createHtmlOutput(
        '<div style="padding: 20px; color: red;">系統設定錯誤，請聯絡管理員</div>'
      );
    }

    // 如果是學生就顯示學生頁面
    if (user["userType"] === "學生") {
      return renderStudentPage(user, configs);
    }

    // 如果是老師就顯示老師頁面
    if (user["userType"] === "老師") {
      return renderTeacherPage(user, configs);
    }
  } catch (err) {
    Logger.log("(doGet)發生錯誤：%s\n%s", err.message, err.stack);

    // 轉義 HTML 特殊字符以防止 XSS
    const escapeHtml = (text) => {
      return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#x27;");
    };

    return HtmlService.createHtmlOutput(`
            <div style="padding: 20px; text-align: center; font-family: Arial, sans-serif;">
                <h2 style="color: #d32f2f;">系統錯誤</h2>
                <p>很抱歉，系統發生錯誤，請稍後再試。</p>
                <p style="color: #666; font-size: 0.9em;">錯誤時間：${new Date().toLocaleString(
                  "zh-TW"
                )}</p>
                <div style="margin-top: 20px; padding: 10px; background-color: #f5f5f5; border-left: 4px solid #d32f2f; text-align: left;">
                    <h3 style="color: #d32f2f; margin-top: 0;">錯誤詳情：</h3>
                    <p><strong>錯誤訊息：</strong> ${escapeHtml(
                      err.message || "未知錯誤"
                    )}</p>
                    <details style="margin-top: 10px;">
                        <summary style="cursor: pointer; color: #666;">顯示詳細堆疊資訊</summary>
                        <pre style="background-color: #fff; padding: 10px; border: 1px solid #ddd; margin-top: 10px; font-size: 12px; overflow: auto;">${escapeHtml(
                          err.stack || "無堆疊資訊"
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
 * @param {Object} configs - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderStudentPage(user, configs) {
  try {
    const template = HtmlService.createTemplateFromFile("index");
    template.loginEmail = Session.getActiveUser().getEmail();
    template.serviceUrl = getServiceUrl();
    template.user = user;
    template.configs = configs;
    template.notifications = getNotifications(configs);
    template.limitOfSchools = getLimitOfSchools();

    const optionData = getOptionData(user);
    template.isJoined = optionData.isJoined;
    template.selectedChoices = optionData.selectedChoices;
    template.departmentOptions = optionData.departmentOptions;

    return setXFrameOptionsSafely(
      template.evaluate().setTitle("四技二專甄選入學志願調查系統")
    );
  } catch (error) {
    Logger.log("(renderStudentPage)渲染學生頁面時發生錯誤：%s", error.message);
    throw error;
  }
}

/**
 * @description 渲染老師頁面
 * @param {Object} user - 使用者資料
 * @param {Object} configs - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderTeacherPage(user, configs) {
  try {
    const studentData = getTraineesDepartmentChoices(user);
    const template = HtmlService.createTemplateFromFile("teacherView");

    template.loginEmail = Session.getActiveUser().getEmail();
    template.serviceUrl = getServiceUrl();
    template.user = user;
    template.configs = configs;
    template.headers = studentData.headers;
    template.data = studentData.data;

    return setXFrameOptionsSafely(
      template.evaluate().setTitle("老師查詢班級學生志願")
    );
  } catch (error) {
    Logger.log("(renderTeacherPage)渲染老師頁面時發生錯誤：%s", error.message);
    throw error;
  }
}

/**
 * @description 處理 POST 請求（安全版本）
 * @param {Object} request - 請求參數
 * @returns {HtmlOutput|ContentService.TextOutput} 回應內容
 */
function doPost(request) {
  try {
    // 驗證請求參數
    if (!validateRequestParameters(request.parameters)) {
      Logger.log("(doPost)POST 請求參數驗證失敗");
      return ContentService.createTextOutput("無效的請求參數").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    Logger.log("(doPost)請求參數：%s", JSON.stringify(request.parameters));

    const user = getUserData();
    if (!user || !user["統一入學測驗報名序號"]) {
      Logger.log("(doPost)無效的使用者或非學生帳號嘗試提交");
      return ContentService.createTextOutput("存取被拒絕").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    const configs = getConfigs();
    if (!configs || !configs["系統關閉時間"]) {
      Logger.log("(doPost)系統參數不完整");
      return ContentService.createTextOutput("系統設定錯誤").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // 檢查截止時間
    const endTime = new Date(configs["系統關閉時間"]);
    const now = new Date();
    const tolerance = 60000; // 1 分鐘容忍時間

    if (isNaN(endTime.getTime())) {
      Logger.log("(doPost)系統關閉時間格式錯誤：%s", configs["系統關閉時間"]);
      return ContentService.createTextOutput("系統時間設定錯誤").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    if (now - endTime > tolerance) {
      Logger.log(
        "(doPost)提交時間已過截止時間，現在：%s，截止：%s",
        now,
        endTime
      );
      return ContentService.createTextOutput("志願調查已結束").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // 驗證和清理輸入資料
    const joinedParam = String(
      request.parameters.isJoinedInput?.[0] || "否"
    ).trim();
    const isJoined = joinedParam === "是";

    let departmentChoices = [];
    if (isJoined) {
      // 取得並驗證志願選擇
      for (let i = 1; i <= limitOfChoices; i++) {
        const choice = String(
          request.parameters[`departmentChoices_${i}`]?.[0] || ""
        ).trim();
        // 驗證志願格式（應為6位數字）
        if (choice && !/^\d{6}$/.test(choice)) {
          Logger.log("(doPost)無效的志願格式：%s", choice);
          return ContentService.createTextOutput("無效的志願格式").setMimeType(
            ContentService.MimeType.TEXT
          );
        }
        departmentChoices.push(choice);
      }

      // 排序志願（空值排到後面）
      departmentChoices.sort((a, b) => {
        if (a === "" && b === "") return 0;
        if (a === "") return 1;
        if (b === "") return -1;
        return Number(a) - Number(b);
      });
    }

    // 更新資料
    const userEmail = Session.getActiveUser().getEmail();
    const row = findValueRow(userEmail, studentChoiceSheet);

    if (!row || row === 0) {
      Logger.log("(doPost)找不到使用者資料列：%s", userEmail);
      return ContentService.createTextOutput("找不到使用者資料").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // 準備更新的資料
    const updateData = isJoined
      ? [joinedParam, ...departmentChoices]
      : [joinedParam, "", "", "", "", "", ""];

    if (updateSpecificRow(row, updateData)) {
      Logger.log("(doPost)成功更新使用者 %s 的志願資料", JSON.stringify(user));
    }

    // 建立日誌記錄
    record = {
      isJoined: isJoined,
      departmentChoices_1: departmentChoices[0],
      departmentChoices_2: departmentChoices[1],
      departmentChoices_3: departmentChoices[2],
      departmentChoices_4: departmentChoices[3],
      departmentChoices_5: departmentChoices[4],
      departmentChoices_6: departmentChoices[5],
    };
    logAdder(user, record);

    // 如果有設定要寄送選填結果通知信，才會寄送
    if (configs["是否寄送選填內容通知信"] === "是") {
      sendResultNotificationEmail(
        user,
        userEmail,
        departmentChoices,
        new Date().toLocaleString("zh-TW", {
          timeZone: "Asia/Taipei",
        }),
        configs
      );
      Logger.log("(doPost)寄送選填內容通知信給使用者 %s 成功", userEmail);
    } else {
      Logger.log("(doPost)未設定寄送選填內容通知信，跳過寄送步驟");
    }

    // 渲染成功頁面
    return renderStudentPage(user, configs);
  } catch (err) {
    Logger.log("(doPost)發生錯誤：%s\n%s", err.message, err.stack);
    return ContentService.createTextOutput("系統錯誤，請稍後再試").setMimeType(
      ContentService.MimeType.TEXT
    );
  }
}

/**
 * @description 渲染成功頁面
 * @param {Object} user - 使用者資料
 * @param {Object} configs - 系統參數
 * @returns {HtmlOutput} HTML 輸出
 */
function renderStudentPage(user, configs) {
  try {
    const template = HtmlService.createTemplateFromFile("success");
    template.loginEmail = Session.getActiveUser().getEmail();
    template.serviceUrl = getServiceUrl();
    template.user = user;
    template.configs = configs;
    template.notifications = getNotifications(configs);
    template.limitOfSchools = getLimitOfSchools();

    const optionData = getOptionData(user);
    template.isJoined = optionData.isJoined;
    template.selectedChoices = optionData.selectedChoices;
    template.departmentOptions = optionData.departmentOptions;

    return setXFrameOptionsSafely(
      template.evaluate().setTitle("四技二專甄選入學志願調查系統")
    );
  } catch (error) {
    Logger.log("(renderStudentPage)渲染成功頁面時發生錯誤：%s", error.message);
    throw error;
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
      Logger.log("XFrameOptionsMode.SAMEORIGIN 未定義，跳過設定");
      return htmlOutput;
    }
  } catch (error) {
    Logger.log("設定 XFrameOptionsMode 時發生錯誤：%s", error.message);
    return htmlOutput;
  }
}

function logAdder(user, record) {
  const departmentOptions = getOptionData(user)["departmentOptions"];
  const departmentName = (option) => {
    if (!option || option === "") return "";

    const match = departmentOptions.filter((dept) => dept.startsWith(option));
    return match.length > 0 ? match[0] : "未知志願";
  };

  logSheet.appendRow([
    Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss"),
    user["信箱"],
    user["班級名稱"],
    user["學號"],
    user["考生姓名"],
    user["統一入學測驗報名序號"],
    user["報考群(類)名稱"],
    record["isJoined"] ? "是" : "否",
    departmentName(record["departmentChoices_1"]),
    departmentName(record["departmentChoices_2"]),
    departmentName(record["departmentChoices_3"]),
    departmentName(record["departmentChoices_4"]),
    departmentName(record["departmentChoices_5"]),
    departmentName(record["departmentChoices_6"]),
  ]);
}
