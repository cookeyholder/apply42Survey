/**
 * @description 顯示統計頁面
 */
function showStatisticsPage() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile(
    "statisticsTemplate.html"
  )
    .setWidth(900)
    .setHeight(700);
  htmlOutput = setXFrameOptionsSafely(htmlOutput); // Use the existing safe wrapper
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "志願選填統計");
}

/**
 * @description 取得原始統計資料供前端使用
 * @returns {Object} 依類群分類的志願統計資料或包含錯誤訊息的物件
 */
function getRawStatisticsData() {
  try {
    if (!studentChoiceSheet) {
      Logger.log("「考生志願列表」工作表不存在");
      return { error: "「考生志願列表」工作表不存在，無法產生統計資料。" };
    }

    const studentData = studentChoiceSheet.getDataRange().getValues();

    if (studentData.length < 2) {
      Logger.log("「考生志願列表」沒有足夠的資料");
      return { error: "「考生志願列表」沒有足夠的資料可供統計。" };
    }

    const studentHeaders = studentData[0];

    // 取得「是否參加集體報名」、「報考群(類)代碼」、「報考群(類)名稱」和志願欄位的索引
    const isJoinedIndex = studentHeaders.indexOf("是否參加集體報名");
    const groupCodeColumnIndex = studentHeaders.indexOf("報考群(類)代碼"); // Added for group code
    const groupNameColumnIndex = studentHeaders.indexOf("報考群(類)名稱");
    const choiceStartIndex = studentHeaders.indexOf("志願1校系代碼");
    const choiceNameStartIndex = studentHeaders.indexOf("志願1校系名稱");

    if (isJoinedIndex === -1) {
      return { error: "找不到「是否參加集體報名」欄位。" };
    }
    if (groupCodeColumnIndex === -1) {
      // Added error check for group code column
      return { error: "找不到「報考群(類)代碼」欄位。" };
    }
    if (groupNameColumnIndex === -1) {
      return { error: "找不到「報考群(類)名稱」欄位。" };
    }
    if (choiceStartIndex === -1 && choiceNameStartIndex === -1) {
      return {
        error:
          "找不到志願代碼或志願名稱相關欄位。請確認欄位名稱是否為「志願[數字]校系代碼」或「志願[數字]校系名稱」。",
      };
    }
    const useCode = choiceStartIndex !== -1;
    const actualChoiceStartIndex = useCode
      ? choiceStartIndex
      : choiceNameStartIndex;

    const statistics = {};

    // Removed choiceToGroupMap creation logic

    // 遍歷「考生志願列表」的每一列資料 (跳過標頭)
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      // 只統計參加集體報名的學生
      if (row[isJoinedIndex] !== "是") {
        continue;
      }

      const studentGroupCode = String(row[groupCodeColumnIndex] || "").trim();
      const studentGroupName = String(row[groupNameColumnIndex] || "").trim();
      let effectiveGroupName;

      if (studentGroupCode && studentGroupName) {
        effectiveGroupName = `${studentGroupCode}${studentGroupName}`;
      } else if (studentGroupName) {
        // Fallback if code is missing
        effectiveGroupName = studentGroupName;
      } else if (studentGroupCode) {
        // Fallback if name is missing
        effectiveGroupName = studentGroupCode;
      } else {
        effectiveGroupName = "其他未分類";
      }

      for (let k = 0; k < limitOfChoices; k++) {
        const choiceValue = row[actualChoiceStartIndex + (useCode ? k * 2 : k)];
        const currentChoiceValue = row[actualChoiceStartIndex + k];

        if (currentChoiceValue && String(currentChoiceValue).trim() !== "") {
          const currentChoice = String(currentChoiceValue).trim();
          if (!statistics[effectiveGroupName]) {
            statistics[effectiveGroupName] = {};
          }
          statistics[effectiveGroupName][currentChoice] =
            (statistics[effectiveGroupName][currentChoice] || 0) + 1;
        }
      }
    }

    // 將統計結果轉換為排序後的陣列
    const sortedStatistics = {};
    for (const groupName in statistics) {
      sortedStatistics[groupName] = Object.entries(statistics[groupName])
        .map(([name, count]) => ({ name, count }))
        .sort((a, b) => b.count - a.count);
    }
    Logger.log("產生的統計資料：%s", JSON.stringify(sortedStatistics));
    return sortedStatistics;
  } catch (err) {
    Logger.log("getRawStatisticsData 發生錯誤: %s", err.message);
    Logger.log("錯誤堆疊: %s", err.stack);
    return { error: "產生統計資料時發生未預期的錯誤：" + err.message };
  }
}

/**
 * @description 取得「考生志願列表」中「報考群(類)名稱」的所有唯一值
 * @returns {Object} 包含唯一群類名稱陣列或錯誤訊息的物件
 */
function getUniqueGroupNames() {
  try {
    if (!studentChoiceSheet) {
      Logger.log("「考生志願列表」工作表不存在");
      return { error: "「考生志願列表」工作表不存在，無法取得群類名稱。" };
    }
    const studentData = studentChoiceSheet.getDataRange().getValues();
    if (studentData.length < 2) {
      Logger.log("「考生志願列表」沒有足夠的資料");
      return { error: "「考生志願列表」沒有足夠的資料可供讀取群類。" };
    }

    const studentHeaders = studentData[0];
    const groupCodeIndex = studentHeaders.indexOf("報考群(類)代碼"); // Added for group code
    const groupNameIndex = studentHeaders.indexOf("報考群(類)名稱");

    if (groupCodeIndex === -1) {
      // Added error check for group code column
      return { error: "在「考生志願列表」中找不到「報考群(類)代碼」欄位。" };
    }
    if (groupNameIndex === -1) {
      return { error: "在「考生志願列表」中找不到「報考群(類)名稱」欄位。" };
    }

    const groupNames = new Set();
    for (let i = 1; i < studentData.length; i++) {
      const groupCode = String(studentData[i][groupCodeIndex] || "").trim();
      const groupNameValue = String(
        studentData[i][groupNameIndex] || ""
      ).trim();
      let combinedGroupName;

      if (groupCode && groupNameValue) {
        combinedGroupName = `${groupCode}${groupNameValue}`;
      } else if (groupNameValue) {
        combinedGroupName = groupNameValue;
      } else if (groupCode) {
        combinedGroupName = groupCode;
      }

      if (combinedGroupName) {
        groupNames.add(combinedGroupName);
      }
    }
    Logger.log(
      "取得的唯一群類名稱：%s",
      JSON.stringify(Array.from(groupNames))
    );
    return { groupNames: Array.from(groupNames).sort() }; // 返回排序後的陣列
  } catch (err) {
    Logger.log("getUniqueGroupNames 發生錯誤: %s", err.message);
    Logger.log("錯誤堆疊: %s", err.stack);
    return { error: "取得唯一群類名稱時發生未預期的錯誤：" + err.message };
  }
}
