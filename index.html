// ============================================================
// スキルアップ管理システム - Google Apps Script (GAS)
// バージョン2：Googleドライブ写真保存対応
// ============================================================

const SHEET_NAME_FEEDBACK = 'フィードバック';
const SHEET_NAME_PROGRESS = '進捗';
const SHEET_NAME_STAFF    = 'スタッフ';
const DRIVE_FOLDER_NAME   = 'TSP練習写真';  // ドライブに作られるフォルダ名

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getFeedback') return getJSON(getFeedback(e.parameter.staffName));
  if (action === 'getProgress') return getJSON(getProgress(e.parameter.staffName));
  if (action === 'getAllStaff') return getJSON(getAllStaff());
  return getJSON({ error: 'unknown action' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  if (action === 'saveFeedback') return getJSON(saveFeedback(data));
  if (action === 'updateProgress') return getJSON(updateProgress(data));
  if (action === 'syncStaff') return getJSON(syncStaff(data));
  return getJSON({ error: 'unknown action' });
}

// ---- スタッフ一覧取得 ----
function getAllStaff() {
  const sheet = getOrCreateSheet(SHEET_NAME_STAFF, ['名前', '役職', '入社日', 'アバター', '現在のステップ']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ---- フィードバック保存（写真対応） ----
function saveFeedback(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_FEEDBACK, [
    '日付', 'スタッフ名', 'スキル', '対象ステップ', '練習時間',
    '評価', 'コメント', '次回ポイント', '記録者', '本人コメント', '写真URL'
  ]);

  // 写真をGoogleドライブに保存
  let photoUrls = '';
  if (data.photos && data.photos.length > 0) {
    const urls = [];
    data.photos.forEach((base64, idx) => {
      try {
        const url = savePhotoToDrive(base64, data.staffName, data.step, idx + 1);
        if (url) urls.push(url);
      } catch(e) {}
    });
    photoUrls = urls.join('\n');
  }

  sheet.appendRow([
    new Date().toLocaleDateString('ja-JP'),
    data.staffName   || '',
    data.skill       || '',
    data.step        || '',
    data.duration    || '',
    data.rating      || '',
    data.comment     || '',
    data.nextPoint   || '',
    data.recorder    || '',
    data.selfComment || '',
    photoUrls,
  ]);

  return { success: true, message: '保存しました', photoUrls };
}

// ---- 写真をGoogleドライブに保存 ----
function savePhotoToDrive(base64DataUrl, staffName, stepName, index) {
  const matches = base64DataUrl.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
  if (!matches || matches.length !== 3) return null;

  const mimeType = matches[1];
  const base64   = matches[2];
  const ext      = mimeType.includes('png') ? 'png' : 'jpg';

  // ルートフォルダ「TSP練習写真」を取得または作成
  const rootFolder  = getOrCreateFolder(DriveApp.getRootFolder(), DRIVE_FOLDER_NAME);

  // スタッフ名フォルダ
  const staffFolder = getOrCreateFolder(rootFolder, staffName);

  // 日付フォルダ（例：2026-05-05）
  const today      = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const dateFolder = getOrCreateFolder(staffFolder, today);

  // ファイル名（例：STEP2_振袖着付_1.jpg）
  const safeName = stepName.replace(/[\/\\:*?"<>|]/g, '_');
  const fileName = `${safeName}_${index}.${ext}`;

  // ファイル保存
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const file = dateFolder.createFile(blob);

  // 閲覧リンクを返す
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ---- フォルダ取得または作成 ----
function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

// ---- フィードバック取得 ----
function getFeedback(staffName) {
  const sheet = getOrCreateSheet(SHEET_NAME_FEEDBACK, [
    '日付', 'スタッフ名', 'スキル', '対象ステップ', '練習時間',
    '評価', 'コメント', '次回ポイント', '記録者', '本人コメント', '写真URL'
  ]);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => !staffName || row[1] === staffName)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    })
    .reverse();
}

// ---- 進捗更新 ----
function updateProgress(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_PROGRESS, [
    'スタッフ名', 'ステップ番号', 'ステップ名', 'ステータス', '更新日'
  ]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.staffName && rows[i][1] === data.stepIndex) {
      sheet.getRange(i + 1, 4).setValue(data.status);
      sheet.getRange(i + 1, 5).setValue(new Date().toLocaleDateString('ja-JP'));
      return { success: true };
    }
  }
  sheet.appendRow([data.staffName, data.stepIndex, data.stepTitle, data.status, new Date().toLocaleDateString('ja-JP')]);
  return { success: true };
}

// ---- 進捗取得 ----
function getProgress(staffName) {
  const sheet = getOrCreateSheet(SHEET_NAME_PROGRESS, [
    'スタッフ名', 'ステップ番号', 'ステップ名', 'ステータス', '更新日'
  ]);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => !staffName || row[0] === staffName)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

// ---- ユーティリティ ----
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- スタッフ一覧をシートに同期 ----
function syncStaff(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_STAFF, ['名前', '役職', '入社日', 'アバター', '現在のステップ']);
  // 既存データをクリアしてヘッダー行以外を書き直す
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  data.staff.forEach(s => {
    const stepStr = JSON.stringify(s.skills || {});
    sheet.appendRow([s.name, '新人スタッフ', s.joined, s.avatar, stepStr]);
  });
  return { success: true };
}
