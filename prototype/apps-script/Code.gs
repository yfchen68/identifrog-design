/**
 * 蛙抵家 2.0 原型｜階段 0 實測錄音上傳端點（Google Apps Script）
 *
 * 設定方式（約 10 分鐘）：
 * 1. 在 Google 雲端硬碟新增一份 Google 試算表，例如「蛙抵家 2.0 實測紀錄」。
 * 2. 試算表選單「擴充功能 → Apps Script」，把這份程式全部貼上，取代原本的內容，按儲存。
 * 3. 上方函式選單選「setup」，按「執行」。第一次會要求授權：
 *    選你的帳號 →「進階」→「前往（專案名稱）（不安全）」→「允許」。
 *    執行完，下方「執行記錄」會顯示上傳金鑰和錄音資料夾的連結。
 * 4. 右上角「部署 → 新增部署作業」，類型選「網頁應用程式」：
 *    執行身分選「我」，誰可以存取選「所有人」，按「部署」，複製網頁應用程式網址（結尾是 /exec）。
 * 5. 把網址和上傳金鑰交給負責原型的人，做成測試連結。
 *
 * 每段錄音會存到試算表旁邊的「蛙抵家 2.0 實測錄音」資料夾，並在「實測錄音」工作表記一列。
 * 要給泰元老師比對，分享這個資料夾和試算表給他即可。
 * 之後如果修改了這份程式，要用「部署 → 管理部署作業 → 編輯 → 版本：新版本」更新，網址不變。
 */

const SHEET_NAME = '實測錄音';
const FOLDER_NAME = '蛙抵家 2.0 實測錄音';
const MAX_BYTES = 20 * 1024 * 1024;
const HEADERS = ['收到時間', '上傳編號', '配對編號', '場域', '備註', '錄音時間（到小時）', '序號', '裝置', '瀏覽環境', '取樣率 Hz', '長度 秒',
  '回音消除', '降噪', '自動增益', '底噪 dBFS', '訊噪差 dB', '99% 頻寬 Hz', '−60 dB 截止 Hz', '原型版本', '檔名', '檔案連結', '診斷摘要'];

/** 第一次使用時執行：建立錄音資料夾、上傳金鑰與工作表。重複執行不會覆蓋既有設定。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從試算表的「擴充功能 → Apps Script」打開這份程式，再執行 setup。');
  const props = PropertiesService.getScriptProperties();

  let folderId = props.getProperty('FOLDER_ID');
  if (!folderId) {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    const folder = parents.hasNext() ? parents.next().createFolder(FOLDER_NAME) : DriveApp.createFolder(FOLDER_NAME);
    folderId = folder.getId();
    props.setProperty('FOLDER_ID', folderId);
  }

  let key = props.getProperty('UPLOAD_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('UPLOAD_KEY', key);
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }

  Logger.log('上傳金鑰：' + key);
  Logger.log('錄音資料夾：' + DriveApp.getFolderById(folderId).getUrl());
}

/** 想換一把上傳金鑰時執行（舊的測試連結會失效）。 */
function resetKey() {
  const key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty('UPLOAD_KEY', key);
  Logger.log('新的上傳金鑰：' + key);
}

/** 原型上傳一段錄音：{ key, id, fileName, data(base64 WAV), meta, summary } */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const props = PropertiesService.getScriptProperties();
    const key = props.getProperty('UPLOAD_KEY');
    if (!key || body.key !== key) return json_({ ok: false, error: 'key' });

    const id = String(body.id || '');
    if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) return json_({ ok: false, error: 'id' });
    if (typeof body.data !== 'string' || !body.data) return json_({ ok: false, error: 'data' });

    lock.waitLock(30000);
    locked = true;

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const folderId = props.getProperty('FOLDER_ID');
    if (!sh || !folderId) return json_({ ok: false, error: 'setup' });

    // 同一段錄音重送時不重複存檔
    const last = sh.getLastRow();
    if (last > 1) {
      const found = sh.getRange(2, 2, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
      if (found) return json_({ ok: true, duplicate: true });
    }

    const bytes = Utilities.base64Decode(body.data);
    if (bytes.length > MAX_BYTES) return json_({ ok: false, error: 'size' });

    const name = cleanName_(body.fileName, id);
    const file = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(bytes, 'audio/wav', name));
    const m = body.meta || {};
    sh.appendRow([
      new Date(), id, m.pairId, m.site, m.note, m.hour, m.seq, m.device, m.environment, m.sampleRate, m.durationSec,
      m.echoCancellation, m.noiseSuppression, m.autoGainControl, m.noiseFloorDbfs, m.snrDb, m.bandwidth99Hz, m.rolloff60Hz,
      m.app, name, file.getUrl(), String(body.summary || '').slice(0, 45000)
    ].map(safe_));
    return json_({ ok: true, file: name });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/** 用瀏覽器打開網址時，確認端點有在運作。 */
function doGet() {
  return ContentService.createTextOutput('蛙抵家 2.0 實測錄音上傳端點：運作中。').setMimeType(ContentService.MimeType.TEXT);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cleanName_(name, id) {
  const s = String(name || '').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '_').slice(0, 120);
  return /\.wav$/i.test(s) ? s : (s || id) + '.wav';
}

/** 避免試算表把以 = + - @ 開頭的文字當成公式 */
function safe_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}
