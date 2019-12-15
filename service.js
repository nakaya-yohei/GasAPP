/***********************************************************************************
 * サービスロジック
 * @author y.nakaya
 * @date 2018/03/01
 * @update 2018/03/01
 *
 * 履歴
 * 2018/03/01 新規リリース
 *
 ***********************************************************************************/

// =================================================================================
// リクエスト処理
// =================================================================================
function doGet(e) {
  var html = HtmlService.createTemplateFromFile('form'); 
  var htmloutput = html.evaluate();
  htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=0.38, user-scalable=no');
  htmloutput.setTitle("参加者リスト出力")
  htmloutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmloutput;  
}

/**
 * 対象イベント関連情報を取得
 * @param  {string} date - 日付
 * @param  {string} eventTitle - イベントタイトル
 *
 * @return {employeesInfo[]} 社員情報 + 分析データを含めた情報を配列で返す
 */
function getGuestData(date,eventTitle) {
   
  // 全社員情報をマスタより取得
  var employeesInfo = getEmployeMasterAll();
  
  // 全社員のイベント回答ステータスを取得
  employeesInfo = getGuestStatus(employeesInfo,date,eventTitle);
  
  // 各ゲストステータスの合計値を取得
  employeesInfo = getGuestStatusSum(employeesInfo);
  
  // 出席者の男女比情報を取得
  employeesInfo = getSexRatio(employeesInfo);
  
  // 幹事者を取得
  employeesInfo = getRepresentative(employeesInfo,date,eventTitle);
  
  return employeesInfo;
}

/**
 * スプレッドシート参加者一覧出力
 * @param  {string} date - 日付
 * @param  {string} eventTitle - イベントタイトル
 * @param  {array[]} members - 参加者
 *
 * @return {boolean} 出力結果
 */
function outputSheet(date,eventTitle,members) {
  
  // 新規スプレッドシート作成しIDを取得
  var sheetId = createSpreadSheet(date,eventTitle);
  
  // フォーマット形式を設定
  designSheet(date,eventTitle,sheetId);
  
  // 参加者を行に挿入
  insertRowMember(members,sheetId);
  
  return;  
}