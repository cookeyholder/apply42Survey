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

    if (foundCell) {
        Logger.log(
            '在 %s 工作表的第 %s 列找到關鍵字: %s',
            sheet.getName(),
            foundCell.getRow(),
            keyword
        );
    } else {
        Logger.log(
            '在 %s 工作表中，沒有找到關鍵字： %s',
            sheet.getName(),
            keyword
        );
    }

    return foundCell ? foundCell.getRow() : 0; // 有找到傳回 row number，否則傳回 0
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

function exportCsv() {
    const parameters = getParameters();
    const now = new Date();
    const nowString = Utilities.formatDate(
        now,
        'Asia/Taipei',
        'yyyy-MM-dd_HHmm'
    );

    const [headers, ...data] = forImportSheet.getDataRange().getValues();

    // 將所有資料轉換成文字格式，並過濾掉完全空白的列
    const processedData = data
        .filter((row) =>
            row.some(
                (cell) =>
                    cell !== null &&
                    cell !== undefined &&
                    cell.toString().trim() !== ''
            )
        )
        .map((row) =>
            row.map((cell) => {
                // 如果是空值，回傳空字串
                if (cell === null || cell === undefined) return '';
                // 將所有值轉換成字串
                return String(cell);
            })
        );

    // 如果沒有有效的資料列，顯示錯誤訊息
    if (processedData.length === 0) {
        const ui = SpreadsheetApp.getUi();
        ui.alert('錯誤', '沒有可匯出的資料，請確認資料內容。', ui.ButtonSet.OK);
        return null;
    }

    const csvContent = [
        headers.map(String).join(','),
        ...processedData.map((row) => row.join(',')),
    ].join('\n');

    // 取得試算表所在的資料夾
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    const parentFolder = spreadsheetFile.getParents().hasNext()
        ? spreadsheetFile.getParents().next()
        : DriveApp.getRootFolder();

    // 在相同資料夾中建立 CSV 檔案
    const fileName =
        parameters['報名學校代碼'] + 'StudQuota_' + nowString + '.csv';
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    const file = parentFolder.createFile(blob);

    // 取得下載連結
    const fileUrl = file.getDownloadUrl();
    Logger.log('CSV 檔案已建立在與試算表相同的資料夾中：%s', fileUrl);

    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(
        `
            <div style="padding: 10px; font-family: Arial, sans-serif;">
                <p>CSV 檔案已建立完成！</p>
                <p><a href="${fileUrl}" target="_blank" download>點此下載檔案</a></p>
                <p style="color: #666; font-size: 0.9em;">檔案已儲存在與試算表相同的資料夾中</p>
            </div>
        `
    )
        .setWidth(300)
        .setHeight(150);

    ui.showModalDialog(htmlOutput, '檔案匯出完成');
    return fileUrl;
}
