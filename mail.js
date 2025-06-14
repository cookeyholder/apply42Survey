/**
 * @description 寄送志願選填結果通知信給學生
 * @param {Object} user - 學生資料物件
 * @param {string} toEmail - 收件人信箱
 * @param {Array} wishes - 已選填的志願陣列 (志願名稱字串)
 * @param {string} submissionTime - 選填時間
 * @param {Object} configs - 系統設定
 * @returns {boolean} - 寄信成功與否
 */
function sendResultNotificationEmail(user, toEmail, wishes, submissionTime, configs) {
  try {
    const departmentOptions = getOptionData(user)["departmentOptions"];
    const departmentName = (option) => {
        if (!option || option === "") return "";

        const match = departmentOptions.filter((dept) => dept.startsWith(option));
        return match.length > 0 ? match[0] : "未知志願";
    };


    // 建立 HTML 範本
    const template = HtmlService.createTemplateFromFile("resultNotificationEamil");
    
    // 注入變數到範本中
    template.className = user['班級名稱'] || "";
    template.studentId = user['學號'] || "";
    template.studentName = user['考生姓名'] || "";
    template.wishes = wishes.filter(w => w && w.trim() !== "").map((wish)=>departmentName(wish));
    template.submissionTime = submissionTime;
    template.configs = configs;
    
    // 產生 HTML 內容
    const htmlBody = template.evaluate().getContent();
    
    // 寄送郵件
    MailApp.sendEmail({
      to: toEmail,
      subject: configs['通知信主旨'],
      htmlBody: htmlBody,
      name: configs['通知信寄件人名稱']
    });

    Logger.log("成功寄送通知信給 %s (%s)", user['考生姓名'], toEmail);
    return true;
  } catch (error) {
    Logger.log("寄送通知信發生錯誤: %s", error.message);
    return false;
  }
}