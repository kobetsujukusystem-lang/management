// スプレッドシートのIDを指定
const SS_ID = "1zoef-3t3aInv53i8Ac2Z7v33KgM12pHxl8PJYPCVGkQ"; 

/**
 * GETリクエスト処理（データの取得用）
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === "getWorkers") {
      return createJsonResponse({ status: "success", data: getWorkersList() });
    }
    
    return createJsonResponse({ status: "error", message: "Invalid action" });
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * POSTリクエスト処理（ログイン、打刻、シフト希望提出など）
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    // 1. ログイン認証
    if (action === "login") {
      const result = authenticateUser(data.id, data.password);
      return createJsonResponse(result);
    }
    
    // 2. 打刻処理
    if (action === "clockIn") {
      recordAttendance(data.id, data.type);
      return createJsonResponse({ status: "success", message: `${data.type}を記録しました。` });
    }
    
    // 3. シフト希望の保存
    if (action === "saveShiftRequest") {
      saveShift(data.id, data.shifts);
      return createJsonResponse({ status: "success", message: "シフト希望を保存しました。" });
    }
    
    return createJsonResponse({ status: "error", message: "Invalid action" });
    
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * 共通のJSONレスポンス作成関数
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 「従業員一覧」シートからデータを取得する
 */
function getWorkersList() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("従業員一覧");
  const rows = sheet.getDataRange().getValues();
  
  const headers = rows[0];
  const data = rows.slice(1);
  
  return data.filter(row => row[0] === "在職" || row[0] === "有効" || !row[0]).map(row => {
    return {
      name: row[1],
      id: row[1], // IDの代わりに「名前」をそのままidとして渡す
    };
  });
}

/**
 * ログイン認証処理
 */
function authenticateUser(inputId, inputPass) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("従業員一覧");
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
const id = String(row[8]);    // ＩＤ列（I列）
const pass = String(row[9]);  // ＰＡＳＳ列（J列）
    const name = row[1];          // 従業員氏名列
    
    if (id === String(inputId) && pass === String(inputPass)) {
      // 認証成功。シート名や権限情報を返す（名前がそのまま個別シート名になっている前提）
      return {
        status: "success",
        name: name,
        sheetName: name, // 個別シート名が氏名と同じであると仮定
        isAdmin: name.includes("管理者") // 例として名前に管理者を含むかなどで判定、必要に応じて列を追加
      };
    }
  }
  
  return { status: "error", message: "IDまたはパスワードが間違っています。" };
}

/**
 * 打刻記録を個人のシートに追記する
 * 「従業員原本」のような構成で、日付, 勤務可能時間帯, シフト, 出勤, 退勤, 休憩開始, 休憩終了... がある想定
 */
function recordAttendance(id, type) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const workerName = getWorkerNameById(ss, id);
  if (!workerName) throw new Error("従業員が見つかりません。");
  
  const sheet = ss.getSheetByName(workerName);
  if (!sheet) throw new Error("個別の勤怠シートが見つかりません: " + workerName);
  
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
  const timeStr = Utilities.formatDate(now, "JST", "HH:mm:ss");
  
  // 個別シートの構造：
  // 日付(A), 勤務可能時間帯(B), シフト(C), 出勤(D), 退勤(E), 休憩開始(F), 休憩終了(F or G?), 交通費(H), 備考(I)
  // ここでは「今日の行」を探すか、新規追加する処理を行います。
  
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  
  // 下から探して今日の日付の行があればそこに書き込む、なければ新しく追加
  for (let i = data.length - 1; i >= 3; i--) { // 4行目以降がデータ行と仮定
    const rowDate = data[i][0];
    if (rowDate) {
      const formattedRowDate = Utilities.formatDate(new Date(rowDate), "JST", "yyyy/MM/dd");
      if (formattedRowDate === todayStr) {
        targetRow = i + 1; // 1-indexed
        break;
      }
    }
  }
  
  // 打刻の種類に応じて書き込む列を決定（例: 出勤=D列(4), 退勤=E列(5), 休憩開始=F列(6), 休憩終了=G列(7)）
  let colIndex = -1;
  if (type === "出勤") colIndex = 4;
  else if (type === "退勤") colIndex = 5;
  else if (type === "休憩開始") colIndex = 6;
  else if (type === "休憩終了") colIndex = 7;
  
  if (targetRow === -1) {
    // 今日分の日付行がまだなければ新しく行を追加する
    // [日付, 勤務可能時間帯, シフト, 出勤, 退勤, 休憩開始, 休憩終了, ...]
    const newRowData = [todayStr, "", "", "", "", "", ""];
    if (colIndex !== -1) {
      newRowData[colIndex - 1] = timeStr;
    }
    sheet.appendRow(newRowData);
  } else {
    // 既存の今日の行に打刻時間を書き込む
    if (colIndex !== -1) {
      sheet.getRange(targetRow, colIndex).setValue(timeStr);
    }
  }
}

/**
 * IDから従業員氏名（シート名）を引く補助関数
 */
function getWorkerNameById(ss, id) {
  const sheet = ss.getSheetByName("従業員一覧");
  const rows = sheet.getDataRange().getValues();
for (let i = 1; i < rows.length; i++) {
  if (String(rows[i][8]) === String(id)) { // I列（8）に合わせる
    return rows[i][1]; // 従業員氏名
  }
}
  return null;
}

/**
 * シフト希望を保存する
 */
function saveShift(id, shifts) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const workerName = getWorkerNameById(ss, id);
  if (!workerName) throw new Error("従業員が見つかりません。");
  
  const sheet = ss.getSheetByName(workerName);
  if (!sheet) throw new Error("個別の勤怠シートが見つかりません: " + workerName);
  
  // shifts は [{date: "2026/10/01", time: "9:00-18:00"}, ...] のような配列を想定
  shifts.forEach(s => {
    // 日付が一致する行を探して「勤務可能時間帯」(B列=2)を更新、なければ追加
    // ※実際のシートレイアウトに合わせて適宜調整してください
    sheet.appendRow([s.date, s.time, "", "", "", "", ""]);
  });
}