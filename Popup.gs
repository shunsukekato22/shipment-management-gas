// ============================================================
// ⚠️ 設定箇所（ポップアップを表示しないヘッダー名を実際の名前に変更してください）
// ============================================================
const IGNORE_HEADER_NAMES = new Set([
  "納品数",  // J列
  "ロット№",  // K列
  "別請求",  // O列
  "出荷準備",  // P列
  "在庫管理出庫",  // Q列
  "備考",  // R列
]);
const FLAG_HEADER_NAME = "※消さないでください（ポップアップ管理列）"; // （フラグ保存列）
// ============================================================

// キャッシュ
const sheetCache = {};
const columnCache = {};

function onEditHandler(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // 対象シートでなければ終了
  if (!isPastSheet(sheetName)) return;

  // ヘッダー行は除外
  if (row <= 2) return;

  // 列マップを取得
  const columnMap = getColumnMap(sheet);

  // フラグ列番号を取得
  const flagColumn = columnMap[FLAG_HEADER_NAME];
  if (!flagColumn) return;

  // 判定しない列を除外（IGNORE列 + フラグ列）
  const ignoreColumns = new Set(
    [...IGNORE_HEADER_NAMES, FLAG_HEADER_NAME]
      .map(name => columnMap[name])
      .filter(Boolean)
  );
  if (ignoreColumns.has(col)) return;

  // フラグ確認（フラグがあればスキップ）
  if (sheet.getRange(row, flagColumn).getValue() === "warned") return;

  // 警告を表示
  SpreadsheetApp.getUi().alert(
    "⚠️ 警告",
    "追加・変更する場合は出荷担当者へ連絡してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  // フラグを保存
  sheet.getRange(row, flagColumn).setValue("warned");
}

// ヘッダー行から列名と列番号のマップを作成
function getColumnMap(sheet) {
  const sheetName = sheet.getName();

  // キャッシュがあれば返す
  if (columnCache[sheetName]) return columnCache[sheetName];

  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[header] = index + 1;
  });

  columnCache[sheetName] = map;
  return map;
}

function isPastSheet(sheetName) {
  // キャッシュに結果があればそれを返す
  if (sheetName in sheetCache) return sheetCache[sheetName];

  // シート名が YY/MM/DD 形式か確認
  const regex = /^\d{2}\/\d{1,2}\/\d{1,2}$/;
  if (!regex.test(sheetName)) {
    sheetCache[sheetName] = false;
    return false;
  }

  const parts = sheetName.split("/");
  const year = 2000 + parseInt(parts[0]);
  const month = parseInt(parts[1]) - 1;
  const day = parseInt(parts[2]);

  // シートの日付
  const sheetDate = new Date(year, month, day);
  sheetDate.setHours(0, 0, 0, 0);

  // 明後日の日付
  const dayAfterTomorrow = new Date();
  dayAfterTomorrow.setHours(0, 0, 0, 0);
  dayAfterTomorrow.setDate(dayAfterTomorrow.getDate() + 2);

  // 結果をキャッシュに保存して返す
  const result = sheetDate < dayAfterTomorrow;
  sheetCache[sheetName] = result;
  return result;
}


// ==============================
function debugFull() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  Logger.log("=== シート確認 ===");
  Logger.log("シート名: " + sheetName);
  Logger.log("対象シートか: " + isPastSheet(sheetName));

  Logger.log("=== ヘッダー確認 ===");
  const columnMap = getColumnMap(sheet);
  Logger.log("列マップ: " + JSON.stringify(columnMap));

  Logger.log("=== フラグ列確認 ===");
  const flagColumn = columnMap[FLAG_HEADER_NAME];
  Logger.log("FLAG_HEADER_NAME: " + FLAG_HEADER_NAME);
  Logger.log("フラグ列番号: " + flagColumn);

  Logger.log("=== 除外列確認 ===");
  const ignoreColumns = [...IGNORE_HEADER_NAMES].map(name => ({
    name,
    col: columnMap[name]
  }));
  Logger.log("除外列: " + JSON.stringify(ignoreColumns));

  Logger.log("=== onEdit確認 ===");
  Logger.log("onEdit関数が存在するか手動で確認してください");
}
