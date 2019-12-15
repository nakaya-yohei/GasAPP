/***********************************************************************************
 * ���ʃ��W�b�N
 * @author y.nakaya
 * @date 2018/03/01
 * @update 2018/03/01
 *
 * ����
 * 2018/03/01 �V�K�����[�X
 * 2018/12/09 �Ј��Ǘ�����ŐV��
 *
 ***********************************************************************************/

// =================================================================================
// �萔(�Ј��}�X�^�擾�p)
// =================================================================================
var EMPLOYE_MASTER_ID = '1UQEri5uBnk0vlGgBu9LThps_Oq3AjuV26ZMwETe1-zU';
var EMPLOYE_SHEET_NAME = '�Ј��ԍ��ꗗ';
var EMPLOY_COL_ID = 1;
var EMPLOYE_COL_NAME = 2;
var EMPLOYE_COL_EMAIL = 4;
var EMPLOYE_COL_SEX = 7;
var EMPLOYE_ROW_START = 2;

/**
 * �Ώۓ��̃C�x���g�^�C�g�����擾
 * @param  {string} selectDate - ���t
 *
 * @return {eventTitles[]} �C�x���g�^�C�g����z���Ԃ�
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
 * �ΏۃC�x���gID�̎擾
 * @param  {string} date - ���t
 * @param  {string} eventTitle - �C�x���g�^�C�g��
 *
 * @return {string} �C�x���gID��Ԃ�
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
 * �Ј��}�X�^���S�Ј������擾
 * �Ј��ԍ�/���O/�A�h���X/����
 *
 * @return {employeesInfo[]} �S�Ј�����z��ŕԂ�
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

    name = name.replace('�@', '');
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
 * �Q�X�g�X�e�[�^�X�̎擾
 * @param  {employeesInfo[]} employeesInfo - �Ј����
 * @param  {String} date - ���t
 * @param  {String} eventTitle - �C�x���g
 * 
 * @return {employeesInfo[]} �Q�X�g�X�e�[�^�X����ǉ������z���Ԃ�
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
 * �Q�X�g�X�e�[�^�X�𕶎���ɕϊ� & �e�X�e�[�^�X�̍��v�l�����擾
 * @param  {employeesInfo[]} employeesInfo - �Ј����
 * 
 * @return {employeesInfo[]} �Q�X�g�X�e�[�^�X����ǉ������z���Ԃ�
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
 * �j����̎擾
 * ��) length�̒����ɒ���
 * @param  {employeesInfo[]} employeesInfo - �Ј���� + ���v���
 * 
 * @return {employeesInfo[]} �Q�X�g�X�e�[�^�X����ǉ������z���Ԃ�
 */
function getSexRatio(employeesInfo) {
  
  var participant = employeesInfo[employeesInfo.length-1].partCount;
  var male = 0;
  var maleRatio = 0;
  var female = 0;
  var femaleRatio = 0;
  
  for (var i = 0; i < employeesInfo.length-1; i++){
    
    if ("�j" == employeesInfo[i].sex && "YES" == employeesInfo[i].status){
      male += 1;
    } else if ("��" == employeesInfo[i].sex && "YES" == employeesInfo[i].status){
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
 * ��\�҂ƃA�N�Z�X���[�U�[�̎擾
 * ��)length�̒����ɒ���
 * @param  {employeesInfo[]} employeesInfo - �Ј���� + ���v��� + �j������
 * 
 * @return {employeesInfo[]} �Q�X�g�X�e�[�^�X����ǉ������z���Ԃ�
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
 * �V�K�X�v���b�h�V�[�g�̍쐬(�}�C�h���C�u�ɕۑ�)
 * @param  {string} eventTitle - �C�x���g��
 * 
 * @return {string} id - �X�v���b�h�V�[�gID
 */
function createSpreadSheet(date,eventTitle) {
  var newSpreadSheet = SpreadsheetApp.create('�y�Q���҈ꗗ�z' + date + '_' + eventTitle);
  var id = newSpreadSheet.getId();
  
  return id;
}

/**
 * �V�[�g�f�U�C������
 * @param  {employeesInfo[]} employeesInfo
 * 
 * @return {employeesInfo[]} 
 */
function designSheet(date,eventTitle,sheetId) {  
  var values = [
                  ['�J�Ó� : ' + date, '', '�C�x���g : ' + eventTitle, ''],
                  ['', '', '', ''],
                  ['�Q���҃��X�g', '', '', ''],
                  ['No', '����', '&', '���l'],
                  ['', '', '', ''],
                  ['���̑�', '', '', ''],    
                  ['No', '����', '&', '���l']
               ];
  var spreadSheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadSheet.getActiveSheet();
  
  // �Z���̕����Z�b�g
  sheet.setColumnWidth(1,30);
  sheet.setColumnWidth(2,120);
  sheet.setColumnWidth(3,60);
  sheet.setColumnWidth(4,120);
  
  // ���ڂ��Z�b�g
  var range1 = sheet.getRange(2, 1, values.length, values[0].length);
  range1.setValues(values);
  
  // �w�i�F���Z�b�g
  var range2 = sheet.getRange('A5:D5');
  range2.setBackground('cyan');
  var range3 = sheet.getRange('A8:D8');
  range3.setBackground('gray');
  
  
  // �t�H���g�X�^�C��(����)���Z�b�g
  var range4 = sheet.getRange('A2:D5');
  range4.setFontWeight('bold');
  var range5 = sheet.getRange('A7:D8');
  range5.setFontWeight('bold');
  
  return;
}

/**
 * �Q���҃��R�[�h�}��
 * @param  {members[]} members - �Q����
 * @param  {string} sheetId - �X�v���b�h�V�[�gID
 * 
 * @return
 */
function insertRowMember(members,sheetId) {
  
  var initRow = 6;
  
  var spreadSheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadSheet.getActiveSheet();
  
  // ���͋K�������[����
  var rule = SpreadsheetApp.newDataValidation();
  rule.requireValueInList(["-","'"], true);
  
  for (var i = 0; i < members.length; i++){
    
    �@var value = [
                   [i + 1, members[i], "-", ""]
                  ];
�@�@�@// 1���R�[�h�}��
     sheet.insertRowBefore(initRow);
    
     var range1 = sheet.getRange(initRow, 1, 1, 4);
     range1.setValues(value);
    
     var range2 = sheet.getRange(initRow, 3);
     range2.setDataValidation(rule);
    
     initRow +=1;
     
  }
 
  return;
}