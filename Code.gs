// 🆕 追加：初回のみ実行してIDを保存、実行後は削除推奨
// function setConfig() {
//   PropertiesService.getScriptProperties().setProperties({
//     "DST_ID": "ここにファイルIDを入れる" //納品控帳ファイルID
//   });
//   console.log("設定完了");
// }


function onEditHandler(e) {

  const SRC_SHEET_NAME = "入力用";
  const FLAG_HEADER = "転記済み";
  const SHIP_DATE_HEADER = "出荷日";

  const DST_ID = PropertiesService.getScriptProperties().getProperty("DST_ID");
  if (!DST_ID) {
    console.error("DST_IDが設定されていません。setConfig()を実行してください。");
    return;
  }

  const TEMPLATE_SHEET_NAME = "原紙";
  const DST_HEADER_ROW = 2;
  const DST_START_ROW = 3;

  if (!e) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SRC_SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 2) return;

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const flagCol = headers.indexOf(FLAG_HEADER) + 1;
  const shipDateCol = headers.indexOf(SHIP_DATE_HEADER) + 1;

  if (flagCol === 0 || shipDateCol === 0) return;
  if (col !== flagCol) return;

  const rowData = sheet
    .getRange(row, 1, 1, headers.length)
    .getValues()[0];

  const shipDate = rowData[shipDateCol - 1];
  if (!(shipDate instanceof Date)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const dstSS = SpreadsheetApp.openById(DST_ID);
    const sheetName = Utilities.formatDate(shipDate, "Asia/Tokyo", "yy/MM/dd");
    let dstSheet = dstSS.getSheetByName(sheetName);
    let sheetCreated = false;

    /* ===============================
      チェックOFF → 削除処理
    =============================== */
    if (e.value !== "TRUE") {

      if (!dstSheet) return;

      const lastRow = dstSheet.getLastRow();
      if (lastRow < DST_START_ROW) return;

      const colAValues = dstSheet
        .getRange(DST_START_ROW, 1, lastRow - DST_START_ROW + 1, 1)
        .getValues()
        .flat();

      const index = colAValues.indexOf(row);

      if (index !== -1) {
        const deleteRowNumber = DST_START_ROW + index;
        const currentLastColumn = dstSheet.getLastColumn();
        dstSheet
          .getRange(deleteRowNumber, 1, 1, currentLastColumn)
          .clearContent();
      }

      return;
    }

    /* ===============================
      チェックON → 転記処理
    =============================== */

    if (!dstSheet) {
      const template = dstSS.getSheetByName(TEMPLATE_SHEET_NAME);
      if (!template) {
        console.error(`テンプレートシート "${TEMPLATE_SHEET_NAME}" が見つかりません`);
        return;
      }

      const newSheet = template.copyTo(dstSS);
      newSheet.setName(sheetName);
      dstSheet = dstSS.getSheetByName(sheetName);
      sheetCreated = true;
    }

    if (!(shipDate instanceof Date)) return;

    // すでに同じ元行があるか確認（A列）
    if (dstSheet) {
      const lastRowCheck = dstSheet.getLastRow();

      if (lastRowCheck >= DST_START_ROW) {
        const existingRows = dstSheet
          .getRange(DST_START_ROW, 1, lastRowCheck - DST_START_ROW + 1, 1)
          .getValues()
          .flat();

        if (existingRows.includes(row)) return;
      }
    }

    const lastRow = dstSheet.getLastRow();
    const lastColumn = dstSheet.getLastColumn();

    const dataRows = Math.max(lastRow - DST_START_ROW + 1, 0);
    const checkRows = Math.min(50, dataRows);
    const startCheckRow = checkRows > 0 ? lastRow - checkRows + 1 : DST_START_ROW;

    const dstHeadersForExclude = dstSheet
      .getRange(DST_HEADER_ROW, 1, 1, lastColumn)
      .getValues()[0];

    const excludeHeaders = ["別請求", "出荷準備", "在庫管理出庫"];

    const excludeIndexes = dstHeadersForExclude.reduce((acc, header, index) => {
      if (excludeHeaders.includes(header)) acc.push(index);
      return acc;
    }, []);

    let startRow = DST_START_ROW;

    const dataRange = checkRows > 0
      ? dstSheet.getRange(startCheckRow, 1, checkRows, lastColumn).getValues()
      : [];

    for (let i = dataRange.length - 1; i >= 0; i--) {
      const rowHasData = dataRange[i].some((cell, index) => {
        if (index === 0) return false;
        if (excludeIndexes.includes(index)) return false;
        return cell !== "" && cell !== null && cell !== false;
      });

      if (rowHasData) {
        startRow = startCheckRow + i + 1;
        break;
      }
    }

    /* ===============================
      ✅ 追加：結合セル対応ヘッダーマップ作成
    =============================== */
    const headerColMap = buildHeaderColumnMap(dstSheet, DST_HEADER_ROW);
    const totalCols = dstSheet.getLastColumn();
    const writeRow = new Array(totalCols).fill("");

    writeRow[0] = row; // A列：元行番号管理

    for (const [headerName, dstColIndex] of Object.entries(headerColMap)) {
      const srcIndex = headers.indexOf(headerName);
      if (srcIndex !== -1) {
        writeRow[dstColIndex] = rowData[srcIndex];
      }
    }

    dstSheet
      .getRange(startRow, 1, 1, writeRow.length)
      .setValues([writeRow]);

    const namedRanges = dstSheet.getNamedRanges();
    const shipDateRange = namedRanges.find(nr => nr.getName().endsWith("出荷日セル"));
    if (shipDateRange && !shipDateRange.getRange().getValue()) {
      shipDateRange.getRange().setValue(
        Utilities.formatDate(shipDate, "Asia/Tokyo", "yy/MM/dd")
      );
    }

    if (sheetCreated) {
      sortSheetsByDate();
    }

  } finally {
    lock.releaseLock();
  }
}


/* ===============================
  ✅ 追加：結合セル対応ヘッダーマップ
=============================== */
function buildHeaderColumnMap(dstSheet, headerRow) {
  const lastCol = dstSheet.getLastColumn();
  const headers = dstSheet
    .getRange(headerRow, 1, 1, lastCol)
    .getValues()[0];

  const map = {}; // { ヘッダー名: 列インデックス(0始まり) }

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];

    // 空白はスキップ（チェックボックス列など）
    if (header === "" || header === null || header === false) continue;

    // 右隣が空 → 結合セルと判断 → 右隣(i+1)をデータ列とする
    const nextHeader = headers[i + 1];
    if (i + 1 < headers.length && (nextHeader === "" || nextHeader === null || nextHeader === false)) {
      map[header] = i + 1; // データ列は右隣
      i++;                  // チェックボックス列（i+1）はスキップ
    } else {
      // 結合なし → そのままの列
      map[header] = i;
    }
  }

  return map;
}


// =================================
// 　　　　　並べ替え用関数
// =================================
function sortSheetsByDate() {

  // ✅ 変更：直書き → PropertiesServiceから取得
  const DST_ID = PropertiesService.getScriptProperties().getProperty("DST_ID");
  if (!DST_ID) {
    console.error("DST_IDが設定されていません。setConfig()を実行してください。");
    return;
  }

  const ss = SpreadsheetApp.openById(DST_ID);
  const sheets = ss.getSheets();

  const dateSheets = sheets
    .filter(s => /^\d{2}\/\d{1,2}\/\d{1,2}$/.test(s.getName()))
    .map(s => {

      const parts = s.getName().split("/");
      const year = 2000 + Number(parts[0]);
      const month = Number(parts[1]) - 1;
      const day = Number(parts[2]);

      return {
        sheet: s,
        date: new Date(year, month, day)
      };
    });

  dateSheets.sort((a, b) => a.date - b.date);

  dateSheets.forEach((obj, index) => {
    ss.setActiveSheet(obj.sheet);
    ss.moveActiveSheet(index + 1);
  });
}







// ================================
//           自動ロック
// ================================
// function protectPastSheets() {

//   const lock = LockService.getScriptLock();
//   lock.waitLock(30000);

//   try {

//     // ✅ 変更：直書き → PropertiesServiceから取得
//     const DST_ID = PropertiesService.getScriptProperties().getProperty("DST_ID");
//     if (!DST_ID) {
//       console.error("DST_IDが設定されていません。setConfig()を実行してください。");
//       return;
//     }

//     const ss = SpreadsheetApp.openById(DST_ID);

//     const permissionSheet = ss.getSheetByName("権限管理");
//     const lastRow = permissionSheet.getLastRow();
//     if (lastRow < 2) return;

//     const permissionData = permissionSheet
//       .getRange(2, 1, lastRow - 1, 2)
//       .getValues();

//     const admins = [];
//     const staff = [];

//     permissionData.forEach(row => {
//       const email = row[0];
//       const role = row[1];
//       if (!email) return;

//       if (role === "管理者") admins.push(email);
//       if (role === "担当者") staff.push(email);
//     });

//     const sheets = ss.getSheets();
//     const today = new Date();
//     today.setHours(0,0,0,0);

//     sheets.forEach(sheet => {

//       const name = sheet.getName();
//       if (!/^\d{2}\/\d{1,2}\/\d{1,2}$/.test(name)) return;

//       const parts = name.split("/");
//       const year = 2000 + Number(parts[0]);
//       const month = Number(parts[1]) - 1;
//       const day = Number(parts[2]);

//       const sheetDate = new Date(year, month, day);
//       sheetDate.setHours(0,0,0,0);

//       const dayOfWeek = today.getDay();
//       const daysToAdd = dayOfWeek === 5 ? 3 : 1;

//       const limitDate = new Date(today);
//       limitDate.setDate(today.getDate() + daysToAdd);

//       if (sheetDate > limitDate) return;

//       const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
//       if (protections.length > 0) return;

//       const protection = sheet.protect();
//       protection.setWarningOnly(false);

//       protection.removeEditors(protection.getEditors());
//       protection.removeEditor(Session.getEffectiveUser());

//       protection.addEditors(admins);

//       if (protection.canDomainEdit()) {
//         protection.setDomainEdit(false);
//       }

//       const editableRangesForStaff = [
//         sheet.getRange("K:K"),
//         sheet.getRange("L:L"),
//         sheet.getRange("P:S")
//       ];

//       protection.setUnprotectedRanges(editableRangesForStaff);
//       protection.addEditors(staff);

//     });
//   } finally {
//     lock.releaseLock();
//   }
// }
