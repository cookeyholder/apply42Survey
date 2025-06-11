// 新增安全性常數
const MAX_CLASS_SIZE = 200; // 最大班級人數
const REQUIRED_MENTOR_FIELDS = ["班級", "姓名"];
const REQUIRED_STUDENT_FIELDS = [
  "班級名稱",
  "考生姓名",
  "統一入學測驗報名序號",
];

/**
 * @description 驗證導師資料的完整性
 * @param {Object} mentor - 導師資料
 * @returns {boolean} 資料是否有效
 */
function validateMentorData(mentor) {
  if (!mentor || typeof mentor !== "object") {
    Logger.log("導師資料無效或為空");
    return false;
  }

  for (const field of REQUIRED_MENTOR_FIELDS) {
    if (
      !mentor[field] ||
      typeof mentor[field] !== "string" ||
      mentor[field].trim() === ""
    ) {
      Logger.log("導師資料缺少必要欄位：%s", field);
      return false;
    }
  }

  return true;
}

/**
 * @description 清理和驗證學生資料
 * @param {Array} row - 學生資料列
 * @param {Array} headers - 標頭陣列
 * @returns {Array|null} 清理後的資料或 null
 */
function sanitizeStudentRow(row, headers) {
  if (!Array.isArray(row) || !Array.isArray(headers)) {
    return null;
  }

  // 清理每個欄位的資料
  return row.map((cell, index) => {
    if (cell === null || cell === undefined) {
      return "";
    }

    const str = cell.toString();

    // 對敏感欄位進行額外清理
    const header = headers[index];
    if (header && (header.includes("姓名") || header.includes("序號"))) {
      // 移除可能的危險字符，但保留中文字符
      return str
        .replace(/[<>="'\\]/g, "")
        .trim()
        .substring(0, 50);
    }

    return str
      .replace(/[<>="'\\]/g, "")
      .trim()
      .substring(0, 100);
  });
}

/**
 * @description 取得導師班級學生的科系志願選擇（安全版本）
 * @param {Object} mentor - 導師資料
 * @returns {{headers: Array, data: Array}} 學生資料
 */
function getTraineesDepartmentChoices(mentor) {
  try {
    // 驗證導師資料
    if (!validateMentorData(mentor)) {
      throw new Error("導師資料驗證失敗");
    }

    if (!studentChoiceSheet) {
      throw new Error("考生志願列表工作表不存在");
    }

    const className = mentor["班級"].toString().trim();
    if (!className) {
      throw new Error("班級名稱不能為空");
    }

    // 安全地取得工作表資料
    const sheetData = getSheetDataSafely(
      studentChoiceSheet,
      REQUIRED_STUDENT_FIELDS
    );
    if (!sheetData) {
      Logger.log("無法取得考生志願列表資料");
      return { headers: [], data: [] };
    }

    const { headers, data } = sheetData;

    // 尋找各欄位的索引
    const fieldIndexes = {
      classIndex: headers.indexOf("班級名稱"),
      serialIndex: headers.indexOf("統一入學測驗報名序號"),
      studentNameIndex: headers.indexOf("考生姓名"),
      groupIndex: headers.indexOf("報考群(類)名稱"),
      paymentTypeIndex: headers.indexOf("繳費身分"),
      feeIndex: headers.indexOf("報名費"),
      isJoinedIndex: headers.indexOf("是否參加集體報名"),
      choiceIndexes: [],
    };

    // 找出所有志願欄位的索引
    for (let i = 1; i <= limitOfChoices; i++) {
      const choiceHeader = `志願${i}校系名稱`;
      const index = headers.indexOf(choiceHeader);
      fieldIndexes.choiceIndexes.push(index);
    }

    // 驗證必要欄位是否存在
    const requiredIndexes = [
      fieldIndexes.classIndex,
      fieldIndexes.studentNameIndex,
      fieldIndexes.serialIndex,
    ];

    if (requiredIndexes.some((index) => index === -1)) {
      throw new Error("工作表缺少必要的欄位");
    }

    // 篩選出該班級的學生資料
    const classStudents = data.filter((row) => {
      const rowClassName = row[fieldIndexes.classIndex];
      return rowClassName && rowClassName.toString().trim() === className;
    });

    // 檢查班級大小
    if (classStudents.length > MAX_CLASS_SIZE) {
      Logger.log("班級人數過多，可能有安全問題：%d 人", classStudents.length);
      return { headers: [], data: [] };
    }

    // 建立輸出標頭
    const outputHeaders = [
      "考生姓名",
      "統一入學測驗報名序號",
      "班級名稱",
      "報考群(類)名稱",
      "繳費身分",
      "報名費",
      "是否參加集體報名",
    ];

    // 加入志願標頭
    for (let i = 1; i <= limitOfChoices; i++) {
      outputHeaders.push(`志願${i}校系名稱`);
    }

    // 處理學生資料
    const processedData = classStudents
      .map((row) => {
        // 基本資料
        const studentData = [
          row[fieldIndexes.studentNameIndex] || "",
          row[fieldIndexes.serialIndex] || "",
          row[fieldIndexes.classIndex] || "",
          row[fieldIndexes.groupIndex] || "",
          row[fieldIndexes.paymentTypeIndex] || "",
          row[fieldIndexes.feeIndex] || "",
          row[fieldIndexes.isJoinedIndex] || "",
        ];

        // 志願資料
        fieldIndexes.choiceIndexes.forEach((choiceIndex) => {
          studentData.push(choiceIndex !== -1 ? row[choiceIndex] || "" : "");
        });

        return sanitizeStudentRow(studentData, outputHeaders);
      })
      .filter((row) => row !== null) // 移除無效資料
      .sort((a, b) => {
        // 按學生姓名排序
        const nameA = a[0] || "";
        const nameB = b[0] || "";
        return nameA.localeCompare(nameB, "zh-TW");
      });

    Logger.log(
      "成功取得班級 %s 的學生資料，共 %d 人",
      className,
      processedData.length
    );

    return {
      headers: outputHeaders,
      data: processedData,
    };
  } catch (error) {
    Logger.log("getTraineesDepartmentChoices 發生錯誤：%s", error.message);
    return { headers: [], data: [] };
  }
}

/**
 * @description 取得班級統計資訊
 * @param {Array} studentData - 學生資料陣列
 * @param {Array} headers - 標頭陣列
 * @returns {Object} 統計資訊
 */
function getClassStatistics(studentData, headers) {
  try {
    if (!Array.isArray(studentData) || !Array.isArray(headers)) {
      return {
        totalStudents: 0,
        respondedStudents: 0,
        joinedStudents: 0,
        averageFee: 0,
      };
    }

    const isJoinedIndex = headers.indexOf("是否參加集體報名");
    const feeIndex = headers.indexOf("報名費");

    if (isJoinedIndex === -1) {
      return {
        totalStudents: studentData.length,
        respondedStudents: 0,
        joinedStudents: 0,
        averageFee: 0,
      };
    }

    let respondedCount = 0;
    let joinedCount = 0;
    let totalFee = 0;
    let feeCount = 0;

    for (const student of studentData) {
      const joinStatus = student[isJoinedIndex];

      if (joinStatus && joinStatus.toString().trim() !== "") {
        respondedCount++;

        if (joinStatus.toString().trim() === "是") {
          joinedCount++;
        }
      }

      if (feeIndex !== -1 && student[feeIndex]) {
        const fee = parseFloat(student[feeIndex]);
        if (!isNaN(fee) && fee >= 0) {
          totalFee += fee;
          feeCount++;
        }
      }
    }

    return {
      totalStudents: studentData.length,
      respondedStudents: respondedCount,
      joinedStudents: joinedCount,
      averageFee: feeCount > 0 ? Math.round(totalFee / feeCount) : 0,
    };
  } catch (error) {
    Logger.log("計算班級統計時發生錯誤：%s", error.message);
    return {
      totalStudents: 0,
      respondedStudents: 0,
      joinedStudents: 0,
      averageFee: 0,
    };
  }
}
