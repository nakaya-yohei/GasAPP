/***********************************************************************************
 * 共通ロジック
 * @author y.nakaya
 * @date 2018/03/01
 * @update 2018/03/01
 *
 * 履歴
 * 2018/03/01 新規リリース
 * 2018/12/09 社員管理簿を最新化
 *
 ***********************************************************************************/

// =================================================================================
// 定数(社員マスタ取得用)
// =================================================================================
var EMPLOYE_MASTER_ID = '1UQEri5uBnk0vlGgBu9LThps_Oq3AjuV26ZMwETe1-zU';
var EMPLOYE_SHEET_NAME = '社員番号一覧';
var EMPLOY_COL_ID = 1;
var EMPLOYE_COL_NAME = 2;
var EMPLOYE_COL_EMAIL = 4;
var EMPLOYE_COL_SEX = 7;
var EMPLOYE_ROW_START = 2;

/**
 * 対象日のイベントタイトルを取得
 * @param  {string} selectDate - 日付
 *
 * @return {eventTitles[]} イベントタイトルを配列を返す
 */
function getEventTitles(selectDate) {
  var eventTitles = new Array();
  var cal = CalendarApp.getDefaultCalendar();
  var events = cal.getEventsForDay(new Date(selectDate));
  
  for (var i = 0; i < events.length; i++){
    eventTitles.push(events[i].getTitle());
  }
  
  return eventTitles;
}

/**
 * 対象イベントIDの取得
 * @param  {string} date - 日付
 * @param  {string} eventTitle - イベントタイトル
 *
 * @return {string} イベントIDを返す
 */
function getEventId(date,eventTitle) {
  
  var eventId = "";
  var cal = CalendarApp.getDefaultCalendar();
  var events = cal.getEventsForDay(new Date(date));
  
  for (var i = 0; i < events.length; i++){
    
    if (eventTitle != events[i].getTitle()){
      continue;
    }
    
    eventId = events[i].getId(); 
  }
  
  return eventId;
}

/**
 * 社員マスタより全社員情報を取得
 * 社員番号/名前/アドレス/性別
 *
 * @return {employeesInfo[]} 全社員情報を配列で返す
 */
function getEmployeMasterAll() {
  
  var employeesInfo = new Array();

  var empMaster = SpreadsheetApp.openById(EMPLOYE_MASTER_ID);
  if (null == empMaster) {
    return null;
  }

  var empSheet = empMaster.getSheetByName(EMPLOYE_SHEET_NAME);
  if (null == empSheet) {
    return null;
  }

  var values = empSheet.getRange(EMPLOYE_ROW_START, 1, empSheet.getLastRow(), empSheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {

    if (EMPLOYE_COL_SEX >= values[i].length) {
      continue;
    }

    var id = values[i][EMPLOY_COL_ID];
    var name = values[i][EMPLOYE_COL_NAME];
    var mail = values[i][EMPLOYE_COL_EMAIL];
    var sex = values[i][EMPLOYE_COL_SEX];

    if ('' == id || '' == name || '' == mail || '' == sex) {
      continue;
    }

    name = name.replace('　', '');
    mail = mail.trim();

    employeesInfo.push({
                        id : parseInt(id),
                        name : name,
                        mail : mail,
                        sex : sex
                       });
  }
  
  return employeesInfo;
}

/**
 * ゲストステータスの取得
 * @param  {employeesInfo[]} employeesInfo - 社員情報
 * @param  {String} date - 日付
 * @param  {String} eventTitle - イベント
 * 
 * @return {employeesInfo[]} ゲストステータス情報を追加した配列を返す
 */
function getGuestStatus(employeesInfo,date,eventTitle) {
  
  var eventId = getEventId(date,eventTitle);
  
  for(var i = 0; i < employeesInfo.length; i++){
    
    var status = "";
    
    try{
       CalendarApp.subscribeToCalendar(employeesInfo[i].mail);
      
       var event = CalendarApp.getCalendarById(employeesInfo[i].mail).getEventById(eventId);
      
       var status = event.getGuestByEmail(employeesInfo[i].mail).getGuestStatus();
      
    }catch(e){
       var status = 'ERROR';
       Logger.log(employeesInfo[i].name);
       Logger.log(employeesInfo[i].mail);
    }
    
    employeesInfo[i]['status'] = status;
    
  }
  
  return employeesInfo;
}

/**
 * ゲストステータスを文字列に変換 & 各ステータスの合計人数を取得
 * @param  {employeesInfo[]} employeesInfo - 社員情報
 * 
 * @return {employeesInfo[]} ゲストステータス情報を追加した配列を返す
 */
function getGuestStatusSum(employeesInfo) {
  
  var partCount = 0;
  var noPartCount = 0;
  var unDecCount = 0;
  var unAnsCount = 0;
  
  var maxCount = 0;
  
  for (var i = 0; i < employeesInfo.length; i++){
    if ("INVITED" == employeesInfo[i].status){
      employeesInfo[i].status = 'INVITED';
      unAnsCount += 1;
    } else if ("YES" == employeesInfo[i].status){
      employeesInfo[i].status = 'YES';
      partCount += 1;
    } else if ("NO" == employeesInfo[i].status){
      employeesInfo[i].status = 'NO';
      noPartCount += 1;
    } else if ("MAYBE" == employeesInfo[i].status){
      employeesInfo[i].status = 'MAYBE';
      unDecCount += 1;
    }
  }
  
  maxCount = Math.max(partCount, noPartCount, unDecCount, unAnsCount);
  
  employeesInfo.push({
                        partCount : partCount,
                        noPartCount : noPartCount,
                        unDecCount : unDecCount,
                        unAnsCount : unAnsCount,
                        maxCount : maxCount
                       });
  
  return employeesInfo;
}

/**
 * 男女比の取得
 * 注) lengthの長さに注意
 * @param  {employeesInfo[]} employeesInfo - 社員情報 + 合計情報
 * 
 * @return {employeesInfo[]} ゲストステータス情報を追加した配列を返す
 */
function getSexRatio(employeesInfo) {
  
  var participant = employeesInfo[employeesInfo.length-1].partCount;
  var male = 0;
  var maleRatio = 0;
  var female = 0;
  var femaleRatio = 0;
  
  for (var i = 0; i < employeesInfo.length-1; i++){
    
    if ("男" == employeesInfo[i].sex && "YES" == employeesInfo[i].status){
      male += 1;
    } else if ("女" == employeesInfo[i].sex && "YES" == employeesInfo[i].status){
      female += 1;
    }
    
  }
    
  maleRatio = Math.floor(male/participant*100);
  femaleRatio = Math.floor(female/participant*100);
  
  employeesInfo.push({
                        maleRatio : maleRatio,
                        femaleRatio : femaleRatio
                       });
  
  return employeesInfo;
}

/**
 * 代表者とアクセスユーザーの取得
 * 注)lengthの長さに注意
 * @param  {employeesInfo[]} employeesInfo - 社員情報 + 合計情報 + 男女比情報
 * 
 * @return {employeesInfo[]} ゲストステータス情報を追加した配列を返す
 */
function getRepresentative(employeesInfo,date,eventTitle) {
  
  var eventId = getEventId(date,eventTitle);
  var event = CalendarApp.getDefaultCalendar().getEventById(eventId);   
  var representatives = event.getCreators();
  var accessUser = Session.getActiveUser().getEmail();
  
  for (var i = 0; i < employeesInfo.length-2; i++){
    
    for (var j = 0; j < representatives.length; j++){
      
      if (employeesInfo[i].mail == representatives[j]){
        
          employeesInfo.push({
                        name : employeesInfo[i].name,
                        representative : representatives[j],
                        accessUser : accessUser
                       });
        
      }
      
    }
    
  }
  
  return employeesInfo;
}

/**
 * 新規スプレッドシートの作成(マイドライブに保存)
 * @param  {string} eventTitle - イベント名
 * 
 * @return {string} id - スプレッドシートID
 */
function createSpreadSheet(date,eventTitle) {
  var newSpreadSheet = SpreadsheetApp.create('【参加者一覧】' + date + '_' + eventTitle);
  var id = newSpreadSheet.getId();
  
  return id;
}

/**
 * シートデザイン処理
 * @param  {employeesInfo[]} employeesInfo
 * 
 * @return {employeesInfo[]} 
 */
function designSheet(date,eventTitle,sheetId) {  
  var values = [
                  ['開催日 : ' + date, '', 'イベント : ' + eventTitle, ''],
                  ['', '', '', ''],
                  ['参加者リスト', '', '', ''],
                  ['No', '氏名', '&', '備考'],
                  ['', '', '', ''],
                  ['その他', '', '', ''],    
                  ['No', '氏名', '&', '備考']
               ];
  var spreadSheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadSheet.getActiveSheet();
  
  // セルの幅をセット
  sheet.setColumnWidth(1,30);
  sheet.setColumnWidth(2,120);
  sheet.setColumnWidth(3,60);
  sheet.setColumnWidth(4,120);
  
  // 項目をセット
  var range1 = sheet.getRange(2, 1, values.length, values[0].length);
  range1.setValues(values);
  
  // 背景色をセット
  var range2 = sheet.getRange('A5:D5');
  range2.setBackground('cyan');
  var range3 = sheet.getRange('A8:D8');
  range3.setBackground('gray');
  
  
  // フォントスタイル(太字)をセット
  var range4 = sheet.getRange('A2:D5');
  range4.setFontWeight('bold');
  var range5 = sheet.getRange('A7:D8');
  range5.setFontWeight('bold');
  
  return;
}

/**
 * 参加者レコード挿入
 * @param  {members[]} members - 参加者
 * @param  {string} sheetId - スプレッドシートID
 * 
 * @return
 */
function insertRowMember(members,sheetId) {
  
  var initRow = 6;
  
  var spreadSheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadSheet.getActiveSheet();
  
  // 入力規則をルール化
  var rule = SpreadsheetApp.newDataValidation();
  rule.requireValueInList(["-","'"], true);
  
  for (var i = 0; i < members.length; i++){
    
    　var value = [
                   [i + 1, members[i], "-", ""]
                  ];
　　　// 1レコード挿入
     sheet.insertRowBefore(initRow);
    
     var range1 = sheet.getRange(initRow, 1, 1, 4);
     range1.setValues(value);
    
     var range2 = sheet.getRange(initRow, 3);
     range2.setDataValidation(rule);
    
     initRow +=1;
     
  }
 
  return;
}