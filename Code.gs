function onEditHandler(e) {


  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 最大30秒待機

  try{
    const SRC_SHEET_NAME = "入力用";
    const FLAG_HEADER = "転記済み";
    const SHIP_DATE_HEADER = "出荷日";

    const DST_ID = "出荷管理ファイルID"; //出荷管理ファイルID
    const TEMPLATE_SHEET_NAME = "原紙";

    const DST_HEADER_ROW = 2;
    const DST_START_ROW = 3;

    if (!e) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SRC_SHEET_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row === 1) return;

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

    const dstSS = SpreadsheetApp.openById(DST_ID);
    const sheetName = Utilities.formatDate(shipDate, "Asia/Tokyo", "yyyy-MM-dd");
    let dstSheet = dstSS.getSheetByName(sheetName);

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

        dstSheet
          .getRange(deleteRowNumber, 1, 1, dstSheet.getLastColumn())
          .clearContent();  // ← deleteRowではなくこれ
      }

      sortSheetsByDate();

      return;
    }
    

    /* ===============================
      チェックON → 転記処理
    =============================== */

    if (!dstSheet) {
      const template = dstSS.getSheetByName(TEMPLATE_SHEET_NAME);

      // 念のため再確認
      dstSheet = dstSS.getSheetByName(sheetName);

      if (!dstSheet) {
        const newSheet = template.copyTo(dstSS);
        newSheet.setName(sheetName);
        dstSheet = dstSS.getSheetByName(sheetName);
      }
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

        if (existingRows.includes(row)) {
          return; // すでに転記済みなら終了
        }
      }
    }

    let startRow = DST_START_ROW;

    const maxRows = dstSheet.getMaxRows();

    // B〜M列（2〜13列目）をまとめて取得
    const dataRange = dstSheet.getRange(
      DST_START_ROW,
      2,
      maxRows - DST_START_ROW + 1,
      12
    ).getValues();

    // 下から探す
    for (let i = dataRange.length - 1; i >= 0; i--) {

      const rowHasData = dataRange[i].some(
        cell => cell !== "" && cell !== null
      );

      if (rowHasData) {
        startRow = DST_START_ROW + i + 1;
        break;
      }
    }

    const dstHeaders = dstSheet
      .getRange(DST_HEADER_ROW, 1, 1, dstSheet.getLastColumn())
      .getValues()[0];

    const writeRow = dstHeaders.map(h => {
      const index = headers.indexOf(h);
      return index !== -1 ? rowData[index] : "";
    });

    /* ★ A列に元行番号を書き込む */
    writeRow[0] = row;

    dstSheet
      .getRange(startRow, 1, 1, writeRow.length)
      .setValues([writeRow]);

    dstSheet.getRange("C1").setValue(shipDate);

    const formattedDate = Utilities.formatDate(shipDate, "Asia/Tokyo", "yyyy-MM-dd");

  // D1が空のときだけ入力
  if (!dstSheet.getRange("D1").getValue()) {
    dstSheet.getRange("D1").setValue(formattedDate);
  }

    sortSheetsByDate();
  } finally {
    lock.releaseLock();
  }
}


// =================================
// 　　　　　並べ替え用関数
// =================================
function sortSheetsByDate() {

  const DST_ID = "出荷管理ファイルID";  // ＜出荷管理ファイルID＞
  const ss = SpreadsheetApp.openById(DST_ID);

  const sheets = ss.getSheets();

  const dateSheets = sheets
    .filter(s => /^\d{4}-\d{2}-\d{2}$/.test(s.getName()))
    .map(s => ({
      sheet: s,
      date: new Date(s.getName())
    }));

  // 古い → 新しい（右が最新）
  dateSheets.sort((a, b) => a.date - b.date);

  // 左から順に並べる
  dateSheets.forEach((obj, index) => {
    ss.setActiveSheet(obj.sheet);
    ss.moveActiveSheet(index + 1);
  });

}


// ================================
//           自動ロック
// ================================
function protectPastSheets() {

  const DST_ID = "1oOhR4rDG396X59jNJWqLUHA8L0lXUluVek4YkqhYgMc";
  const ss = SpreadsheetApp.openById(DST_ID);

  const permissionSheet = ss.getSheetByName("権限管理");
  const permissionData = permissionSheet
    .getRange(2, 1, permissionSheet.getLastRow() - 1, 2)
    .getValues();

  const admins = [];
  const staff = [];

  permissionData.forEach(row => {
    const email = row[0];
    const role = row[1];
    if (!email) return;

    if (role === "管理者") admins.push(email);
    if (role === "担当者") staff.push(email);
  });

  const sheets = ss.getSheets();
  const today = new Date();
  today.setHours(0,0,0,0);

  sheets.forEach(sheet => {

    const name = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(name)) return;

    const sheetDate = new Date(name);
    sheetDate.setHours(0,0,0,0);
    if (sheetDate > today) return;

    // 既存保護削除
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
         .forEach(p => p.remove());

    const protection = sheet.protect();
    protection.setWarningOnly(false);

    // まず全員削除
    protection.removeEditors(protection.getEditors());

    // 管理者はフル編集可
    protection.addEditors(admins);

    // ドメイン編集禁止
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    // 担当者は特定列のみ編集可
    const editableRangesForStaff = [
      sheet.getRange("J:J"),
      sheet.getRange("K:K"),
      sheet.getRange("O:R")
    ];

    protection.setUnprotectedRanges(editableRangesForStaff);

    // 担当者も編集者として追加
    protection.addEditors(staff);

  });
}
