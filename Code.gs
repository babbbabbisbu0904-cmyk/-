/**
 * 引き継ぎノート デジタル化プロトタイプ（複数建物対応版）
 * Google Apps Script（コンテナバインド版：スプレッドシートに紐づけて使用）
 *
 * シート構成:
 *   職員マスタ       : 氏名 / メールアドレス / 建物(未使用・legacy) / 管理者フラグ(TRUE/FALSE)
 *   職員建物マッピング : 氏名 / メールアドレス / 建物   ※1人が複数建物に出勤する場合は複数行
 *   建物マスタ       : 建物名
 *   シフト表         : 日付 / 建物 / 氏名 / メールアドレス
 *   引き継ぎログ     : ID / 登録日時 / 建物 / 記入者名 / 記入者メール / 重要度 / 内容
 *   既読ログ         : 引き継ぎID / 確認者名 / 確認者メール / 確認日時
 */

const SHEET_STAFF = '職員マスタ';
const SHEET_STAFF_BUILDINGS = '職員建物マッピング';
const SHEET_BUILDINGS = '建物マスタ';
const SHEET_SHIFT = 'シフト表';
const SHEET_LOG = '引き継ぎログ';
const SHEET_READ = '既読ログ';

const UNREAD_FALLBACK_DAYS = 30; // シフト実績が無い職員向けのフォールバック期間
const ADMIN_LOOKBACK_DAYS = 14;  // 管理者ビューで表示する期間

/* =========================================
 * メニュー / 初期セットアップ
 * ========================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('引き継ぎシステム')
    .addItem('① 初期セットアップ（シート作成）', 'setupSpreadsheet')
    .addItem('② サンプルデータを入力（テスト用）', 'insertSampleData')
    .addItem('③ 毎朝のリマインドを設定', 'createDailyReminderTrigger')
    .addItem('リマインドを今すぐ送信（テスト用）', 'sendMorningReminders')
    .addSeparator()
    .addItem('【移行】職員マスタの建物列→職員建物マッピングへ', 'migrateStaffBuildings')
    .addToUi();
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheetWithHeaders_(ss, SHEET_STAFF, ['氏名', 'メールアドレス', '建物(未使用)', '管理者フラグ']);
  createSheetWithHeaders_(ss, SHEET_STAFF_BUILDINGS, ['氏名', 'メールアドレス', '建物']);
  createSheetWithHeaders_(ss, SHEET_BUILDINGS, ['建物名']);
  createSheetWithHeaders_(ss, SHEET_SHIFT, ['日付', '建物', '氏名', 'メールアドレス']);
  createSheetWithHeaders_(ss, SHEET_LOG, ['ID', '登録日時', '建物', '記入者名', '記入者メール', '重要度', '内容']);
  createSheetWithHeaders_(ss, SHEET_READ, ['引き継ぎID', '確認者名', '確認者メール', '確認日時']);
  SpreadsheetApp.getUi().alert(
    'シート構成を作成しました。「職員マスタ」「職員建物マッピング」「建物マスタ」「シフト表」にデータを入力してください。\n' +
    '1人のスタッフが複数建物に出勤する場合は、職員建物マッピングに氏名・メールアドレスが同じ行を建物の数だけ追加してください。'
  );
}

function createSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// 既存の「1人1建物」版（職員マスタの建物列）から職員建物マッピングへデータを移行する
function migrateStaffBuildings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) {
    SpreadsheetApp.getUi().alert('職員マスタが見つかりません。');
    return;
  }
  const data = staffSheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf('氏名');
  const emailIdx = headers.indexOf('メールアドレス');
  const buildingIdx = headers.indexOf('建物') !== -1 ? headers.indexOf('建物') : headers.indexOf('建物(未使用)');

  if (buildingIdx === -1) {
    SpreadsheetApp.getUi().alert('職員マスタに建物列が見つかりません（既に移行済みの可能性があります）。');
    return;
  }

  let mappingSheet = ss.getSheetByName(SHEET_STAFF_BUILDINGS);
  if (!mappingSheet) mappingSheet = ss.insertSheet(SHEET_STAFF_BUILDINGS);

  const existing = mappingSheet.getLastRow() > 1
    ? mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 3).getValues()
    : [];
  const existingKeys = new Set(existing.map(r => r[1] + '|' + r[2])); // メール|建物

  const newRows = data
    .slice(1)
    .filter(r => r[emailIdx] && r[buildingIdx])
    .map(r => [r[nameIdx], r[emailIdx], r[buildingIdx]])
    .filter(r => !existingKeys.has(r[1] + '|' + r[2]));

  if (mappingSheet.getLastRow() < 1) {
    mappingSheet.getRange(1, 1, 1, 3).setValues([['氏名', 'メールアドレス', '建物']]).setFontWeight('bold');
    mappingSheet.setFrozenRows(1);
  }

  if (newRows.length > 0) {
    mappingSheet.getRange(mappingSheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
  }

  SpreadsheetApp.getUi().alert(
    newRows.length + '件のデータを職員建物マッピングへ移行しました。\n' +
    '複数建物に出勤するスタッフは、職員建物マッピングシートに同じ氏名・メールアドレスの行を建物の数だけ追加してください。\n' +
    '職員マスタの「建物」列は今後使用しません（残しておいても動作に影響はありません）。'
  );
}

function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const buildingSheet = ss.getSheetByName(SHEET_BUILDINGS);
  buildingSheet.getRange(2, 1, 2, 1).setValues([['第一ホーム'], ['第二ホーム']]);

  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  staffSheet.getRange(2, 1, 3, 4).setValues([
    ['山田太郎', 'yamada@example.com', '第一ホーム', 'TRUE'],
    ['佐藤花子', 'sato@example.com', '第一ホーム', 'FALSE'],
    ['鈴木一郎', 'suzuki@example.com', '第二ホーム', 'FALSE']
  ]);

  // 佐藤花子は第一ホームと第二ホームの両方に出勤する例（複数建物のサンプル）
  const mappingSheet = ss.getSheetByName(SHEET_STAFF_BUILDINGS);
  mappingSheet.getRange(2, 1, 4, 3).setValues([
    ['山田太郎', 'yamada@example.com', '第一ホーム'],
    ['佐藤花子', 'sato@example.com', '第一ホーム'],
    ['佐藤花子', 'sato@example.com', '第二ホーム'],
    ['鈴木一郎', 'suzuki@example.com', '第二ホーム']
  ]);

  const shiftSheet = ss.getSheetByName(SHEET_SHIFT);
  const today = new Date();
  const threeDaysAgo = new Date(today);
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  // 佐藤花子は3日前に第一ホーム、今日は第二ホームに出勤 → 建物ごとに前回出勤日が異なる例
  shiftSheet.getRange(2, 1, 2, 4).setValues([
    [threeDaysAgo, '第一ホーム', '佐藤花子', 'sato@example.com'],
    [today, '第二ホーム', '佐藤花子', 'sato@example.com']
  ]);

  SpreadsheetApp.getUi().alert(
    'サンプルデータを入力しました。「職員マスタ」「職員建物マッピング」のメールアドレスを、実際にテストで使うご自身のGoogleアカウントに書き換えてから動作確認してください。'
  );
}

/* =========================================
 * 共通ヘルパー
 * ========================================= */

function getSheetData_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName + '（先にメニューから初期セットアップを実行してください）');
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function stripTime_(date) {
  const d = new Date(date);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fmt_(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
}

function getStaffBasic_(email) {
  const row = getSheetData_(SHEET_STAFF).find(r => r['メールアドレス'] === email);
  if (!row) return null;
  return {
    name: row['氏名'],
    email: row['メールアドレス'],
    isAdmin: String(row['管理者フラグ']).toUpperCase() === 'TRUE'
  };
}

function getStaffBuildings_(email) {
  return getSheetData_(SHEET_STAFF_BUILDINGS)
    .filter(r => r['メールアドレス'] === email)
    .map(r => r['建物'])
    .filter(String);
}

function assertStaffHasBuilding_(email, building) {
  const buildings = getStaffBuildings_(email);
  if (buildings.indexOf(building) === -1) {
    throw new Error('この建物への権限がありません: ' + building);
  }
}

/* =========================================
 * 未読判定ロジック（建物ごと）
 * ========================================= */

// この職員の、指定した建物での「前回出勤日（本日より前で直近の出勤日）」を取得
function getLastShiftDate_(email, building) {
  const today = stripTime_(new Date());
  const shiftDates = getSheetData_(SHEET_SHIFT)
    .filter(r => r['メールアドレス'] === email && r['建物'] === building)
    .map(r => stripTime_(r['日付']))
    .filter(d => d < today);
  if (shiftDates.length === 0) return null;
  return new Date(Math.max.apply(null, shiftDates));
}

// 指定メールアドレス・建物の未読引き継ぎ一覧（サーバー内部用）
function getUnreadHandoversFor_(email, building) {
  const lastShift = getLastShiftDate_(email, building);
  const cutoff = lastShift || new Date(Date.now() - UNREAD_FALLBACK_DAYS * 24 * 60 * 60 * 1000);

  const readIds = new Set(
    getSheetData_(SHEET_READ)
      .filter(r => r['確認者メール'] === email)
      .map(r => r['引き継ぎID'])
  );

  return getSheetData_(SHEET_LOG)
    .filter(r => r['建物'] === building && new Date(r['登録日時']) > cutoff)
    .filter(r => !readIds.has(r['ID']))
    .sort((a, b) => new Date(b['登録日時']) - new Date(a['登録日時']));
}

// 本日出勤予定の建物、無ければ直近の出勤建物、それも無ければマッピング上の先頭を返す
function getDefaultBuilding_(email, buildings) {
  const todayStr = Utilities.formatDate(stripTime_(new Date()), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const shifts = getSheetData_(SHEET_SHIFT).filter(r => r['メールアドレス'] === email);

  const todayShift = shifts.find(
    r => Utilities.formatDate(stripTime_(r['日付']), Session.getScriptTimeZone(), 'yyyy-MM-dd') === todayStr
  );
  if (todayShift) return todayShift['建物'];

  const pastShifts = shifts
    .filter(r => stripTime_(r['日付']) < stripTime_(new Date()))
    .sort((a, b) => new Date(b['日付']) - new Date(a['日付']));
  if (pastShifts.length > 0) return pastShifts[0]['建物'];

  return buildings[0];
}

/* =========================================
 * クライアント（Index.html）から呼び出す関数
 * ========================================= */

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const basic = getStaffBasic_(email);
  if (!basic) {
    throw new Error('職員マスタに登録されていないアカウントです: ' + (email || '(取得不可)'));
  }
  const buildings = getStaffBuildings_(email);
  if (buildings.length === 0) {
    throw new Error(basic.name + 'さんに建物が割り当てられていません。職員建物マッピングシートを確認してください。');
  }
  return {
    name: basic.name,
    email: basic.email,
    isAdmin: basic.isAdmin,
    buildings: buildings,
    defaultBuilding: getDefaultBuilding_(email, buildings)
  };
}

function getUnreadHandovers(building) {
  const email = Session.getActiveUser().getEmail();
  assertStaffHasBuilding_(email, building);
  return getUnreadHandoversFor_(email, building).map(r => ({
    ID: r['ID'],
    登録日時: fmt_(r['登録日時']),
    建物: r['建物'],
    記入者名: r['記入者名'],
    重要度: r['重要度'],
    内容: r['内容']
  }));
}

function getRecentHandovers(building) {
  const email = Session.getActiveUser().getEmail();
  assertStaffHasBuilding_(email, building);

  const readIds = new Set(
    getSheetData_(SHEET_READ)
      .filter(r => r['確認者メール'] === email)
      .map(r => r['引き継ぎID'])
  );

  return getSheetData_(SHEET_LOG)
    .filter(r => r['建物'] === building)
    .sort((a, b) => new Date(b['登録日時']) - new Date(a['登録日時']))
    .slice(0, 30)
    .map(r => ({
      ID: r['ID'],
      登録日時: fmt_(r['登録日時']),
      記入者名: r['記入者名'],
      重要度: r['重要度'],
      内容: r['内容'],
      isRead: readIds.has(r['ID'])
    }));
}

function markAsRead(handoverId) {
  const email = Session.getActiveUser().getEmail();
  const basic = getStaffBasic_(email);
  if (!basic) throw new Error('職員マスタに登録されていないアカウントです: ' + email);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_READ);
  sheet.appendRow([handoverId, basic.name, email, new Date()]);
  return true;
}

function addHandover(building, content, importance) {
  const email = Session.getActiveUser().getEmail();
  assertStaffHasBuilding_(email, building);
  const basic = getStaffBasic_(email);
  if (!basic) throw new Error('職員マスタに登録されていないアカウントです: ' + email);
  if (!content || !content.trim()) throw new Error('内容が空です');

  const id = Utilities.getUuid();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  sheet.appendRow([id, new Date(), building, basic.name, email, importance, content]);

  notifyNewHandover_(building, basic.name, content, email);
  return true;
}

function getAdminOverview() {
  const email = Session.getActiveUser().getEmail();
  const basic = getStaffBasic_(email);
  if (!basic || !basic.isAdmin) throw new Error('管理者権限がありません: ' + email);

  const buildings = getSheetData_(SHEET_BUILDINGS).map(r => r['建物名']);
  const logs = getSheetData_(SHEET_LOG);
  const reads = getSheetData_(SHEET_READ);
  const mapping = getSheetData_(SHEET_STAFF_BUILDINGS);
  const cutoff = new Date(Date.now() - ADMIN_LOOKBACK_DAYS * 24 * 60 * 60 * 1000);

  return buildings.map(building => {
    const buildingStaffNames = mapping.filter(m => m['建物'] === building).map(m => m['氏名']);
    const entries = logs
      .filter(l => l['建物'] === building && new Date(l['登録日時']) > cutoff)
      .sort((a, b) => new Date(b['登録日時']) - new Date(a['登録日時']))
      .map(l => {
        const readers = reads.filter(r => r['引き継ぎID'] === l['ID']).map(r => r['確認者名']);
        const unreaders = buildingStaffNames.filter(n => readers.indexOf(n) === -1);
        return {
          id: l['ID'],
          date: fmt_(l['登録日時']),
          author: l['記入者名'],
          importance: l['重要度'],
          content: l['内容'],
          readers: readers,
          unreaders: unreaders
        };
      });
    return { building: building, entries: entries };
  });
}

/* =========================================
 * 通知 / リマインド
 * ========================================= */

function notifyNewHandover_(building, authorName, content, authorEmail) {
  const recipients = getSheetData_(SHEET_STAFF_BUILDINGS)
    .filter(r => r['建物'] === building && r['メールアドレス'] !== authorEmail)
    .map(r => r['メールアドレス'])
    .filter(String);
  const uniqueRecipients = Array.from(new Set(recipients));
  if (uniqueRecipients.length === 0) return;

  MailApp.sendEmail({
    to: uniqueRecipients.join(','),
    subject: '【引き継ぎ】' + building + ' - 新しい引き継ぎ事項があります',
    body:
      authorName + 'さんより新しい引き継ぎが登録されました。\n\n' +
      '内容:\n' + content + '\n\n' +
      '確認はこちら:\n' + ScriptApp.getService().getUrl()
  });
}

// 時間主導トリガーで毎朝実行し、本日出勤予定の建物ごとに未読が残っている職員へリマインド
function sendMorningReminders() {
  const todayStr = Utilities.formatDate(stripTime_(new Date()), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const staffToday = getSheetData_(SHEET_SHIFT).filter(r => {
    return Utilities.formatDate(stripTime_(r['日付']), Session.getScriptTimeZone(), 'yyyy-MM-dd') === todayStr;
  });

  staffToday.forEach(shift => {
    const email = shift['メールアドレス'];
    const building = shift['建物'];
    if (!email || !building) return;
    const unread = getUnreadHandoversFor_(email, building);
    if (unread.length > 0) {
      MailApp.sendEmail({
        to: email,
        subject: '【要確認】' + building + 'の未読の引き継ぎが' + unread.length + '件あります',
        body:
          '本日「' + building + '」に出勤予定です。出勤前に引き継ぎ内容をご確認ください。\n\n' +
          '確認はこちら:\n' + ScriptApp.getService().getUrl()
      });
    }
  });
}

function createDailyReminderTrigger() {
  deleteExistingTriggers_('sendMorningReminders');
  ScriptApp.newTrigger('sendMorningReminders').timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert('毎朝6時に未読リマインドを送信するトリガーを設定しました。');
}

function deleteExistingTriggers_(functionName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(t);
  });
}

/* =========================================
 * Webアプリのエントリーポイント
 * ========================================= */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('引き継ぎノート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
