/**
 * E-yan Coin App - 大阪支社 (v6.4 Auto-Repair Edition)
 * 改修点：残高やランクが空欄の場合、自動的に初期値(100/素浪人)を適用してエラーを防ぐ
 */

// --- ★設定エリア (Config) ---
const APP_CONFIG = {
  INITIAL_COIN: 100,           // 月初の所持コイン
  MULTIPLIER_DIFF_DEPT: 1.2,   // 他部署倍率
  MESSAGE_MAX_LENGTH: 100,     // メッセージ文字数上限
  ECONOMY_THRESHOLD_L2: 6500,  // 景気Lv2閾値
  ECONOMY_THRESHOLD_L3: 13500, // 景気Lv3閾値
  REMINDER_THRESHOLD: 50,      // リマインド閾値

  // ID設定
  SS_ID: '1E0qf3XM-W8TM5HZ_SrPPoGAV4kwObvS6FmQdaFR3Bpw', // メインSS
  ARCHIVE_SS_ID: '1Gk3B_yd0q-sqskmQwHBsWk0PfYbSqfD0UdzYYiMhN5w', // アーカイブSS
  JSON_FILE_ID: '1K-9jVyC8SK9_g8AS87WxuI1Ax_IeiX7X'
};

const SHEET_NAMES = {
  USERS: 'Users',
  TRANSACTIONS: 'Transactions',
  DEPARTMENTS: 'Departments',
  ARCHIVE_LOG: 'Archive_Log',
  MVP_HISTORY: 'MVP_History',
  CONFIG: 'Config'
};

// --- Web App Entry Points ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('E-yan Coin - 大阪支社')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Config取得 ---
function getSharedEmailConfig() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('CONFIG_SHARED_EMAIL');
  if (cached) return cached;

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return 'kusano@race-tech.co.jp'; 

  const data = sheet.getDataRange().getValues();
  let val = '';
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]).trim() === 'SHARED_EMAIL') {
      val = data[i][1];
      break;
    }
  }
  if(val) cache.put('CONFIG_SHARED_EMAIL', val, 21600);
  return val || 'kusano@race-tech.co.jp';
}

// --- 起動直後のデータ取得 ---
function getInitialData() {
  const activeEmail = Session.getActiveUser().getEmail();
  const sharedEmail = getSharedEmailConfig();
  
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = userSheet.getDataRange().getValues();
    const header = data.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[String(h).trim()] = i); // 空白除去

    const emailColIndex = colIdx['login_email']; 
    const targetColIdx = (emailColIndex !== undefined) ? emailColIndex : colIdx['user_id'];

    // --- 分岐1: 共有アカウントの場合 ---
    if (activeEmail === sharedEmail) {
      const userList = [];
      data.forEach(row => {
        const rowEmail = (emailColIndex !== undefined) ? row[emailColIndex] : '';
        if (!rowEmail || rowEmail === '' || rowEmail === sharedEmail) {
          userList.push({
            user_id: row[colIdx.user_id],    
            name: row[colIdx.name],          
            department: row[colIdx.department] 
          });
        }
      });
      
      return {
        success: true,
        mode: 'SHARED_LOGIN', 
        userList: userList,
        config: APP_CONFIG,
        economy: getEconomyStateCached() 
      };
    }

    // --- 分岐2: 個人アカウントの場合 ---
    let myRow = null;
    for(let i=0; i<data.length; i++) {
      if(data[i][targetColIdx] === activeEmail) {
        myRow = data[i];
        break;
      }
    }

    if (!myRow) {
      return {
        error: 'NOT_REGISTERED',
        departments: getDepartmentsCached()
      };
    }

    const currentUser = buildUserData(myRow, colIdx);
    
    return {
      success: true,
      mode: 'PERSONAL_LOGIN',
      user: currentUser,
      economy: getEconomyStateCached(),
      config: APP_CONFIG,
      dataVersion: new Date().getTime().toString()
    };

  } catch (e) {
    console.error('Error:', e);
    throw new Error('起動エラー: ' + e.message);
  }
}

// --- 共有モードでのログイン処理 ---
function loginAsSharedUser(targetUserId) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = userSheet.getDataRange().getValues();
  const header = data.shift();
  const colIdx = {};
  header.forEach((h, i) => colIdx[String(h).trim()] = i); // 空白除去

  let myRow = null;
  for(let i=0; i<data.length; i++) {
    if(String(data[i][colIdx.user_id]) === String(targetUserId)) {
      myRow = data[i];
      break;
    }
  }

  if (!myRow) throw new Error('ユーザーが見つかりません: ' + targetUserId);
  return buildUserData(myRow, colIdx);
}

// ★修正: データ構築ヘルパー（空データ時の自動補正を追加）
function buildUserData(row, colIdx) {
  // 残高が空なら初期値を設定
  let balance = row[colIdx.wallet_balance];
  if (balance === '' || balance === null || balance === undefined) {
    balance = APP_CONFIG.INITIAL_COIN;
  }

  // ランクが空なら初期値を設定
  let rank = row[colIdx.rank];
  if (!rank || rank === '') {
    rank = '素浪人';
  }

  const currentUser = {
    user_id: row[colIdx.user_id],
    name: row[colIdx.name],
    department: row[colIdx.department],
    rank: rank,
    wallet_balance: balance,
    lifetime_received: row[colIdx.lifetime_received],
    memo: row[colIdx.memo]
  };

  const lastMonthMVP = getLastMonthMVP();
  currentUser.isMVP = (lastMonthMVP === currentUser.user_id);

  let dailySent = 0;
  try {
    if (currentUser.memo) {
      const memoObj = JSON.parse(currentUser.memo);
      const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      if (memoObj.last_sent_date === todayStr) {
        dailySent = memoObj.daily_total || 0;
      }
    }
  } catch (e) {}
  currentUser.dailySent = dailySent;
  return currentUser;
}

// --- ユーザーリスト取得 ---
function getUserListData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    const header = data.shift();
    const c = {};
    header.forEach((h,i) => c[String(h).trim()]=i);
    
    const usersArray = data.map(row => [
      row[c.user_id], 
      row[c.name], 
      row[c.department]
    ]);
    return { success: true, list: usersArray, from: 'Sheet' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// --- 送金処理 (★修正: 残高読み取り時の空対策を追加) ---
function sendAirCoin(receiverId, comment, amountInput, isHidden, senderIdOverride) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: '混雑中。再試行してください。' };

  try {
    const amount = Number(amountInput);
    if (amount > 10) throw new Error('1回10枚までです。');

    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    
    const activeEmail = Session.getActiveUser().getEmail();
    const sharedEmail = getSharedEmailConfig();
    const isSharedMode = (activeEmail === sharedEmail);

    // ★セキュリティ強化: senderIdOverrideの妥当性検証
    if (!isSharedMode && senderIdOverride) {
      return { success: false, message: '不正な操作です。' };
    }
    if (isSharedMode && !senderIdOverride) {
      return { success: false, message: '利用者情報が不足しています。ログインし直してください。' };
    }

    const data = userSheet.getDataRange().getValues();
    const header = data.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[String(h).trim()] = i); // 空白除去
    const emailColIndex = colIdx['login_email'] !== undefined ? colIdx['login_email'] : colIdx['user_id'];

    // --- 送信者の特定 ---
    let senderRowIndex = -1;
    let senderData = null;

    if (isSharedMode) {
      if (!senderIdOverride) throw new Error('利用者情報が不足しています(共有モード)');
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][colIdx.user_id]) === String(senderIdOverride)) {
          senderRowIndex = i;
          senderData = data[i];
          break;
        }
      }
    } else {
      for (let i = 0; i < data.length; i++) {
        if (data[i][emailColIndex] === activeEmail) {
          senderRowIndex = i;
          senderData = data[i];
          break;
        }
      }
    }

    // --- 受信者の特定 ---
    let receiverRowIndex = -1;
    let receiverData = null;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][colIdx.user_id]) === String(receiverId)) {
        receiverRowIndex = i;
        receiverData = data[i];
        break;
      }
    }

    if (senderRowIndex === -1) throw new Error('送信者が見つかりません');
    if (receiverRowIndex === -1) throw new Error('受信者が見つかりません');
    if (String(senderData[colIdx.user_id]) === String(receiverData[colIdx.user_id])) {
      throw new Error('自分には送れません');
    }

    const memoJsonStr = senderData[colIdx.memo] || "{}";
    let memoObj = {};
    try { memoObj = JSON.parse(memoJsonStr); } catch(e) {}

    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (memoObj.last_sent_date !== todayStr) {
      memoObj.last_sent_date = todayStr;
      memoObj.daily_total = 0;
    }
    if (!memoObj.monthly_log) memoObj.monthly_log = {};

    if ((memoObj.daily_total + amount) > 20) throw new Error(`1日上限(20枚)を超えます。`);
    
    const currentTargetCount = memoObj.monthly_log[receiverId] || 0;
    if ((currentTargetCount + amount) > 30) throw new Error(`この人への月間上限(30枚)を超えます。`);

    // ★修正: 残高読み取り時の空対策
    let currentBalanceVal = senderData[colIdx.wallet_balance];
    if (currentBalanceVal === '' || currentBalanceVal === null) {
        currentBalanceVal = APP_CONFIG.INITIAL_COIN; // 空なら初期値とみなす
    }
    const currentBalance = Number(currentBalanceVal);
    
    if (currentBalance < amount) throw new Error('コイン不足');

    const isSameDept = senderData[colIdx.department] === receiverData[colIdx.department];
    const multiplier = isSameDept ? 1 : Number(APP_CONFIG.MULTIPLIER_DIFF_DEPT);
    const valueGained = Math.floor(amount * multiplier);

    const newBal = currentBalance - amount;
    const newLife = Number(receiverData[colIdx.lifetime_received]) + valueGained;

    memoObj.daily_total += amount;
    memoObj.monthly_log[receiverId] = currentTargetCount + amount;

    let newRank = receiverData[colIdx.rank];
    if (newLife >= 120) newRank = '天下人';
    else if (newLife >= 90) newRank = '豪商';
    else if (newLife >= 45) newRank = '商人';
    else if (newLife >= 12) newRank = '丁稚';

    const now = new Date();
    userSheet.getRange(senderRowIndex + 2, colIdx.wallet_balance + 1).setValue(newBal);
    userSheet.getRange(senderRowIndex + 2, colIdx.memo + 1).setValue(JSON.stringify(memoObj));
    userSheet.getRange(senderRowIndex + 2, colIdx.last_updated + 1).setValue(now);

    userSheet.getRange(receiverRowIndex + 2, colIdx.lifetime_received + 1).setValue(newLife);
    if (newRank !== receiverData[colIdx.rank]) {
      userSheet.getRange(receiverRowIndex + 2, colIdx.rank + 1).setValue(newRank);
    }

    const shareFlag = isHidden ? 1 : 0;

    transSheet.appendRow([
      Utilities.getUuid(), now, 
      senderData[colIdx.user_id],   
      receiverData[colIdx.user_id], 
      senderData[colIdx.department], receiverData[colIdx.department],
      amount, multiplier, amount, valueGained, comment,
      shareFlag
    ]);

    const cache = CacheService.getScriptCache();
    cache.remove('HISTORY_' + senderData[colIdx.user_id]); 
    cache.remove('HISTORY_' + receiverData[colIdx.user_id]);
    cache.remove('RANKINGS_v6');
    cache.remove('ECONOMY_STATE_v5');

    return {
      success: true, message: '送信完了！',
      newBalance: newBal, dailySent: memoObj.daily_total
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 履歴取得 (共有モード対応版) ---
function getUserHistory(userIdOverride) {
  const activeEmail = Session.getActiveUser().getEmail();
  const sharedEmail = getSharedEmailConfig();
  const isSharedMode = (activeEmail === sharedEmail);

  // ★共有モード時はuserIdOverrideを使用、個人モード時はactiveEmailを使用
  let targetIdentifier;
  let cacheKey;

  if (isSharedMode) {
    if (!userIdOverride) return { success: true, history: [] }; // user_idがない場合は空
    targetIdentifier = userIdOverride;
    cacheKey = 'HISTORY_' + userIdOverride;
  } else {
    targetIdentifier = activeEmail;
    cacheKey = 'HISTORY_' + activeEmail;
  }

  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if(cached) return { success: true, history: JSON.parse(cached) };

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {success:true, history:[]};

  const history = [];
  const CHUNK = 200;
  let curr = lastRow;

  // ★共有モード時はuser_idで、個人モード時はemailで検索
  while(curr >= 2 && history.length < 20) {
    const start = Math.max(2, curr - CHUNK + 1);
    const numRows = curr - start + 1;

    const data = sheet.getRange(start, 1, numRows, 11).getValues();

    for(let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      // ★修正: row[2], row[3]はuser_id列
      if(row[2] === targetIdentifier || row[3] === targetIdentifier) {
        history.push({
          timestamp: row[1],
          sender_id: row[2],
          receiver_id: row[3],
          sender_dept: row[4],
          amount: row[6],
          value: row[9],
          message: row[10],
          type: row[2] === targetIdentifier ? 'sent' : 'received'
        });
        if(history.length >= 20) break;
      }
    }

    curr -= CHUNK;
    if(lastRow - curr > 3000) break; // 最大3000行まで
  }

  cache.put(cacheKey, JSON.stringify(history), 21600); // 6時間キャッシュ
  return { success: true, history: history };
}

// --- ヘルパー ---
function getSpreadsheet() { return SpreadsheetApp.openById(APP_CONFIG.SS_ID); }

function getDepartmentsCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('DEPT_LIST');
  if (cached) return JSON.parse(cached);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (!sheet) return [];
  const list = sheet.getRange(2, 1, sheet.getLastRow()-1 || 1, 1).getValues().flat().filter(String);
  cache.put('DEPT_LIST', JSON.stringify(list), 3600);
  return list;
}

function getEconomyStateCached(){
    const cache = CacheService.getScriptCache();
    let economyState = cache.get('ECONOMY_STATE_v5');
    if (!economyState) {
      economyState = analyzeEconomyState();
      cache.put('ECONOMY_STATE_v5', economyState, 600);
    }
    return economyState;
}

function analyzeEconomyState() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!transSheet) return 'level2';
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return 'level2';
  const startRow = 2;
  const data = transSheet.getRange(startRow, 10, lastRow - startRow + 1, 1).getValues();
  let totalValue = 0;
  for (let i = 0; i < data.length; i++) {
    totalValue += Number(data[i][0] || 0);
  }
  const l2 = APP_CONFIG.ECONOMY_THRESHOLD_L2;
  const l3 = APP_CONFIG.ECONOMY_THRESHOLD_L3;
  if (totalValue >= l3) return 'level3';
  if (totalValue >= l2) return 'level2';
  return 'level1';
}

function registerNewUser(form) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const email = Session.getActiveUser().getEmail();
  sheet.appendRow([email, form.name, form.department, '素浪人', APP_CONFIG.INITIAL_COIN, 0, '{}', new Date(), email]);
  return { success: true, message: '登録完了' };
}

function getRankings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RANKINGS_v6');
  if (cached) return { success: true, rankings: JSON.parse(cached) };

  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const userData = userSheet.getDataRange().getValues();
  userData.shift();

  const userMap = {};
  const deptHeadcount = {};
  
  userData.forEach(r => {
    const dept = r[2];
    userMap[r[0]] = { name: r[1], dept: dept }; 
    if(dept) deptHeadcount[dept] = (deptHeadcount[dept] || 0) + 1;
  });

  const mvp = userData.map(r => ({name: r[1], dept: r[2], score: Number(r[5])}))
    .sort((a,b) => b.score - a.score).slice(0, 10);

  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();
  const deptCountMap = {}; 
  const deptCoinMap = {};
  const giverMap = {};

  if(lastRow >= 2) {
    const start = 2;
    const tData = transSheet.getRange(start, 3, lastRow - start + 1, 8).getValues();

    tData.forEach(r => {
      const senderId = r[0]; 
      const senderDept = r[2];
      const amount = Number(r[4] || 0);

      if(senderDept) {
        deptCountMap[senderDept] = (deptCountMap[senderDept]||0) + 1;
        deptCoinMap[senderDept] = (deptCoinMap[senderDept]||0) + amount;
      }
      if(senderId) giverMap[senderId] = (giverMap[senderId]||0) + 1;
    });
  }

  const dept = Object.keys(deptCountMap).map(k => {
    const count = deptCountMap[k];
    const headcount = deptHeadcount[k] || 1;
    return { name: k, score: parseFloat((count / headcount).toFixed(2)) };
  }).sort((a,b) => b.score - a.score).slice(0, 5);

  const deptTotal = Object.keys(deptCoinMap).map(k => ({ name: k, score: deptCoinMap[k] })).sort((a,b) => b.score - a.score).slice(0, 5);

  const giver = Object.keys(giverMap).map(k => {
    const u = userMap[k] || { name: k, dept: '不明' }; 
    return { name: u.name, dept: u.dept, score: giverMap[k] };
  }).sort((a,b) => b.score - a.score).slice(0, 10);

  const rankings = { mvp: mvp, dept: dept, deptTotal: deptTotal, giver: giver };
  cache.put('RANKINGS_v6', JSON.stringify(rankings), 900);
  return { success: true, rankings: rankings };
}

function getArchiveMonths() {
  if (!APP_CONFIG.ARCHIVE_SS_ID) return { success: false, message: 'Archive SS Not Configured' };
  const ss = SpreadsheetApp.openById(APP_CONFIG.ARCHIVE_SS_ID);
  const sheets = ss.getSheets();
  const months = sheets.map(s => s.getName()).filter(name => name.match(/^\d{4}_\d{2}$/)).sort().reverse();
  return { success: true, months: months };
}

function getArchiveRankingData(sheetName) {
   try {
    const archiveSS = SpreadsheetApp.openById(APP_CONFIG.ARCHIVE_SS_ID);
    const sheet = archiveSS.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    return { success: true, rankings: { mvp: [], dept: [], giver: [] } }; 
   } catch(e) { return { success: false, message: e.message }; }
}
// --- MVP履歴機能 ---
function saveMVPHistory(userId, score) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!sheet) return;
  const ym = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  sheet.appendRow([ym, userId, score, new Date()]);
}

function getLastMonthMVP() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return null;
  return sheet.getRange(sheet.getLastRow(), 2).getValue(); // user_idを返す
}

// --- メール通知用ヘルパー関数 ---
function getUserMapForEmail(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const colIdx = {};
  header.forEach((h, i) => colIdx[String(h).trim()] = i);

  // ★修正: login_email列を優先的に使用、なければuser_id列を使用
  const emailCol = colIdx['login_email'] !== undefined
    ? colIdx['login_email']
    : colIdx['user_id'];

  const nameCol = colIdx['name'];

  data.forEach(row => {
    const email = row[emailCol];
    const name = row[nameCol];
    if (email) map[email] = name;
  });

  return map;
}

function escapeToEntities(text) {
  if (!text) return '';
  return String(text).split('').map(c => {
    const code = c.charCodeAt(0);
    return (code > 127) ? '&#' + code + ';' : c;
  }).join('');
}

// --- メール通知機能 ---
function sendDailyReportEmail() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No transactions to report');
    return;
  }

  // 日付設定
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 昨日の日付文字列（タイトル用：yyyy/MM/dd形式）
  const dateTitle = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');

  // データ取得（L列まで取得修正済み）
  const transData = transSheet.getRange(2, 1, lastRow - 1, 12).getValues();
  
  // ユーザーマッピング作成
  const usersData = usersSheet.getDataRange().getValues();
  const usersHeader = usersData.shift();
  const colIdx = {};
  usersHeader.forEach((h, i) => colIdx[String(h).trim()] = i);

  const userIdToEmail = {};
  const userIdToName = {};
  usersData.forEach(row => {
    const userId = row[colIdx['user_id']];
    const loginEmail = row[colIdx['login_email']] || row[colIdx['user_id']];
    const name = row[colIdx['name']];
    if (userId) {
      userIdToEmail[userId] = loginEmail;
      userIdToName[userId] = name;
    }
  });

  // フィルタリング（前日 かつ 公開のみ）
  const filteredTrans = transData.filter(row => {
    const ts = new Date(row[1]);
    const shareFlagVal = row[11]; // L列
    const shareFlag = (shareFlagVal === "" || shareFlagVal === null) ? 0 : Number(shareFlagVal);
    return ts >= yesterday && ts < today && shareFlag === 0;
  });

  if (filteredTrans.length === 0) {
    Logger.log('No public transactions yesterday');
    return;
  }

  // ---------------------------------------------------------
  // ★デザイン修正ここから：画像に合わせたカード型デザイン
  // ---------------------------------------------------------
  
  // ベースのスタイル
  const styles = {
    body: 'font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; color: #333333; line-height: 1.6;',
    container: 'max-width: 800px; margin: 0 auto; padding: 20px;',
    headerTitle: 'color: #4a4aae; font-size: 22px; font-weight: bold; margin-bottom: 5px;', // 青紫系のタイトル
    subText: 'font-size: 14px; color: #555555; margin-bottom: 25px;',
    highlight: 'background-color: #fff176; padding: 0 4px; font-weight: bold;', // 黄色マーカー
    note: 'font-size: 12px; color: #888888; display: block; margin-top: 4px;',
    
    // カード部分のスタイル
    card: 'margin-bottom: 15px; padding: 12px 15px; background-color: #fcfcfc; border-left: 4px solid #4a4aae; border-radius: 2px;',
    names: 'font-size: 14px; font-weight: bold; color: #333; margin-bottom: 6px;',
    arrow: 'color: #999; margin: 0 5px; font-weight: normal;',
    honorific: 'font-size: 12px; font-weight: normal; color: #555;',
    message: 'font-size: 14px; color: #444; white-space: pre-wrap; line-height: 1.5;'
  };

  let htmlBody = `<html><body style="${styles.body}">`;
  htmlBody += `<div style="${styles.container}">`;

  // ヘッダー部分
  htmlBody += `<div style="${styles.headerTitle}">${dateTitle} の称賛履歴</div>`;
  htmlBody += `<div style="${styles.subText}">`;
  htmlBody += `昨日の<span style="${styles.highlight}">E-yan</span> Coinメッセージをお届けします。<br>`;
  htmlBody += `<span style="${styles.note}">※共有不可のものは除外されています。</span>`;
  htmlBody += `</div>`;

  // メッセージ一覧ループ
  filteredTrans.forEach(row => {
    const senderId = row[2];
    const receiverId = row[3];
    const senderName = userIdToName[senderId] || senderId;
    const receiverName = userIdToName[receiverId] || receiverId;
    const message = row[10];

    // 各カードの生成
    htmlBody += `<div style="${styles.card}">`;
    
    // 送信者 ⇒ 受信者
    htmlBody += `<div style="${styles.names}">`;
    htmlBody += `${escapeToEntities(senderName)} <span style="${styles.honorific}">さん</span>`;
    htmlBody += `<span style="${styles.arrow}">⇒</span>`;
    htmlBody += `${escapeToEntities(receiverName)} <span style="${styles.honorific}">さん</span>`;
    htmlBody += `</div>`;
    
    // メッセージ本文（カギカッコ付き）
    htmlBody += `<div style="${styles.message}">`;
    htmlBody += `「${escapeToEntities(message)}」`;
    htmlBody += `</div>`;
    
    htmlBody += `</div>`; // card end
  });

  htmlBody += '</div></body></html>';
  // ---------------------------------------------------------
  // ★デザイン修正ここまで
  // ---------------------------------------------------------

  const subject = `【ええやんコイン】昨日の送信レポート (${dateTitle})`;

  // メール送信処理
  Object.keys(userIdToEmail).forEach(userId => {
    const email = userIdToEmail[userId];
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
      } catch (e) {
        Logger.log('Failed to send email to ' + email + ': ' + e.message);
      }
    }
  });

  Logger.log('Daily report emails sent to ' + Object.keys(userIdToEmail).length + ' users');
}
function sendDailyRecap() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return;

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const transData = transSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // ★修正: user_id to email/name マッピング作成
  const usersData = usersSheet.getDataRange().getValues();
  const usersHeader = usersData.shift();
  const colIdx = {};
  usersHeader.forEach((h, i) => colIdx[String(h).trim()] = i);

  const userIdToEmail = {};
  const userIdToName = {};
  usersData.forEach(row => {
    const userId = row[colIdx['user_id']];
    const loginEmail = row[colIdx['login_email']] || row[colIdx['user_id']];
    const name = row[colIdx['name']];
    if (userId) {
      userIdToEmail[userId] = loginEmail;
      userIdToName[userId] = name;
    }
  });

  const receivedMap = {};

  transData.forEach(row => {
    const ts = new Date(row[1]);
    if (ts >= yesterday && ts < today) {
      const receiverId = row[3];
      if (!receivedMap[receiverId]) receivedMap[receiverId] = [];
      receivedMap[receiverId].push({
        sender: row[2],
        amount: row[6],
        message: row[10]
      });
    }
  });

  Object.keys(receivedMap).forEach(receiverId => {
    const email = userIdToEmail[receiverId];
    const receiverName = userIdToName[receiverId];
    if (!email) return;

    const items = receivedMap[receiverId];
    let htmlBody = '<html><body style="font-family:sans-serif;">';
    htmlBody += '<h2 style="color:#4CAF50;">' + escapeToEntities(receiverName + ' さん、昨日のええやんコイン受信レポート') + '</h2>';
    htmlBody += '<p>' + escapeToEntities('昨日、あなたは以下のコインを受け取りました：') + '</p>';
    htmlBody += '<ul>';

    items.forEach(item => {
      const senderName = userIdToName[item.sender] || item.sender;
      htmlBody += '<li><strong>' + escapeToEntities(senderName) + '</strong> ' + escapeToEntities('から') + ' <strong>' + item.amount + '</strong> ' + escapeToEntities('枚');
      if (item.message) {
        htmlBody += '<br>' + escapeToEntities('メッセージ: "' + item.message + '"');
      }
      htmlBody += '</li>';
    });

    htmlBody += '</ul>';
    htmlBody += '<p style="margin-top:20px; color:#666;">' + escapeToEntities('今日もええやんコインで感謝を伝えましょう！') + '</p>';
    htmlBody += '</body></html>';

    const subject ='【E-yan Coin】メッセージが届いています';

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody
      });
    } catch (e) {
      Logger.log('Failed to send recap to ' + email + ': ' + e.message);
    }
  });

  Logger.log('Daily recap sent to ' + Object.keys(receivedMap).length + ' users');
}

function checkInactivity() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return;

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  sevenDaysAgo.setHours(0, 0, 0, 0);

  const transData = transSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // ★修正: user_id to email/name マッピング作成
  const usersData = usersSheet.getDataRange().getValues();
  const usersHeader = usersData.shift();
  const colIdx = {};
  usersHeader.forEach((h, i) => colIdx[String(h).trim()] = i);

  const userIdToEmail = {};
  const userIdToName = {};
  const allUserIds = [];

  usersData.forEach(row => {
    const userId = row[colIdx['user_id']];
    const loginEmail = row[colIdx['login_email']] || row[colIdx['user_id']];
    const name = row[colIdx['name']];
    if (userId) {
      userIdToEmail[userId] = loginEmail;
      userIdToName[userId] = name;
      allUserIds.push(userId);
    }
  });

  const activeUsers = {};
  transData.forEach(row => {
    const ts = new Date(row[1]);
    if (ts >= sevenDaysAgo) {
      const senderId = row[2];
      activeUsers[senderId] = true;
    }
  });

  const inactiveUsers = allUserIds.filter(userId => !activeUsers[userId]);

  inactiveUsers.forEach(userId => {
    const email = userIdToEmail[userId];
    const name = userIdToName[userId];
    if (!email) return;

    const subject = '【ええやんコイン】最近ご利用がありません';
    let htmlBody = '<html><body style="font-family:sans-serif;">';
    htmlBody += '<h2 style="color:#FF9800;">' + escapeToEntities(name + ' さん、最近ええやんコインをご利用されていませんね') + '</h2>';
    htmlBody += '<p>' + escapeToEntities('過去7日間、コインの送信がありません。') + '</p>';
    htmlBody += '<p>' + escapeToEntities('感謝の気持ちを伝えてみませんか？') + '</p>';
    htmlBody += '</body></html>';

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody
      });
    } catch (e) {
      Logger.log('Failed to send inactivity notice to ' + email + ': ' + e.message);
    }
  });

  Logger.log('Inactivity notices sent to ' + inactiveUsers.length + ' users');
}

function sendReminderEmails() {
  const ss = getSpreadsheet();
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return;

  const data = usersSheet.getDataRange().getValues();
  const header = data.shift();
  const colIdx = {};
  header.forEach((h, i) => colIdx[String(h).trim()] = i);

  const threshold = 50;

  data.forEach(row => {
    const userId = row[colIdx['user_id']];
    const name = row[colIdx['name']];
    const balance = Number(row[colIdx['balance']] || 0);
    const loginEmail = row[colIdx['login_email']] || userId;

    if (balance >= threshold) {
      const subject = '【ええやんコイン】残高のお知らせ';
      let htmlBody = '<html><body style="font-family:sans-serif;">';
      htmlBody += '<h2 style="color:#2196F3;">' + escapeToEntities(name + ' さん、ええやんコインの残高が ' + balance + ' 枚あります') + '</h2>';
      htmlBody += '<p>' + escapeToEntities('コインを使って感謝を伝えましょう！') + '</p>';
      htmlBody += '</body></html>';

      try {
        MailApp.sendEmail({
          to: loginEmail,
          subject: subject,
          htmlBody: htmlBody
        });
      } catch (e) {
        Logger.log('Failed to send reminder to ' + loginEmail + ': ' + e.message);
      }
    }
  });

  Logger.log('Reminder emails sent');
}

// --- 月次リセット機能 ---
function resetMonthlyData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(600000)) return;
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const archiveId = APP_CONFIG.ARCHIVE_SS_ID;
    const initialCoin = APP_CONFIG.INITIAL_COIN;

    const userData = userSheet.getDataRange().getValues();
    const header = userData.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);

    // ★MVP判定: user_idベースで実施
    let mvpUserId = '';
    let maxLifetime = -1;
    userData.forEach(row => {
      const lifetime = Number(row[colIdx.lifetime_received] || 0);
      if (lifetime > maxLifetime) {
        maxLifetime = lifetime;
        mvpUserId = row[colIdx.user_id];
      }
    });
    if (mvpUserId && maxLifetime > 0) saveMVPHistory(mvpUserId, maxLifetime);

    // ★アーカイブ処理: トランザクションを前月シートに移動
    if (transSheet.getLastRow() > 1 && archiveId) {
      try {
        const archiveSS = SpreadsheetApp.openById(archiveId);
        const now = new Date();
        const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        const sheetName = Utilities.formatDate(lastMonth, 'Asia/Tokyo', 'yyyy_MM');
        let targetSheet = archiveSS.getSheetByName(sheetName);
        if (!targetSheet) {
          targetSheet = archiveSS.insertSheet(sheetName);
          const headers = transSheet.getRange(1, 1, 1, transSheet.getLastColumn()).getValues();
          targetSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        }
        const transData = transSheet.getRange(2, 1, transSheet.getLastRow() - 1, transSheet.getLastColumn()).getValues();
        if (transData.length > 0) {
          targetSheet.getRange(targetSheet.getLastRow() + 1, 1, transData.length, transData[0].length).setValues(transData);
        }

        // ★修正: 固定されていない行をすべて削除することはできないエラーを回避するため、
        // 強制的に1行目を固定してから削除を実行する
        transSheet.setFrozenRows(1);

        const currentLastRow = transSheet.getLastRow();
        if (currentLastRow >= 2) {
          transSheet.deleteRows(2, currentLastRow - 1);
        }

        let logSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE_LOG);
        if(!logSheet) logSheet = ss.insertSheet(SHEET_NAMES.ARCHIVE_LOG);
        logSheet.appendRow([new Date(), `Archived to ${sheetName}`, `${transData.length} rows`]);
      } catch (e) { Logger.log('Archive failed: ' + e.message); }
    }

    // ★ユーザーデータのリセット: ランク、残高、lifetime_received、memoを初期化
    const numRows = userData.length;
    if (numRows > 0) {
      userSheet.getRange(2, colIdx.rank + 1, numRows, 1).setValue('素浪人');
      userSheet.getRange(2, colIdx.wallet_balance + 1, numRows, 1).setValue(initialCoin);
      userSheet.getRange(2, colIdx.lifetime_received + 1, numRows, 1).setValue(0);
      userSheet.getRange(2, colIdx.memo + 1, numRows, 1).setValue('{}');
    }

    // ★キャッシュのクリア
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_USERS_DATA_v4');
    cache.remove('ECONOMY_STATE_v5');
    cache.remove('RANKINGS_v6');

    Logger.log('Monthly reset completed successfully');
  } catch (e) {
    Logger.log('Monthly reset failed: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}
