/***********************************************************************************
 * �T�[�r�X���W�b�N
 * @author y.nakaya
 * @date 2018/03/01
 * @update 2018/03/01
 *
 * ����
 * 2018/03/01 �V�K�����[�X
 *
 ***********************************************************************************/

// =================================================================================
// ���N�G�X�g����
// =================================================================================
function doGet(e) {
  var html = HtmlService.createTemplateFromFile('form'); 
  var htmloutput = html.evaluate();
  htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=0.38, user-scalable=no');
  htmloutput.setTitle("�Q���҃��X�g�o��")
  htmloutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmloutput;  
}

/**
 * �ΏۃC�x���g�֘A�����擾
 * @param  {string} date - ���t
 * @param  {string} eventTitle - �C�x���g�^�C�g��
 *
 * @return {employeesInfo[]} �Ј���� + ���̓f�[�^���܂߂�����z��ŕԂ�
 */
function getGuestData(date,eventTitle) {
   
  // �S�Ј������}�X�^���擾
  var employeesInfo = getEmployeMasterAll();
  
  // �S�Ј��̃C�x���g�񓚃X�e�[�^�X���擾
  employeesInfo = getGuestStatus(employeesInfo,date,eventTitle);
  
  // �e�Q�X�g�X�e�[�^�X�̍��v�l���擾
  employeesInfo = getGuestStatusSum(employeesInfo);
  
  // �o�Ȏ҂̒j��������擾
  employeesInfo = getSexRatio(employeesInfo);
  
  // �����҂��擾
  employeesInfo = getRepresentative(employeesInfo,date,eventTitle);
  
  return employeesInfo;
}

/**
 * �X�v���b�h�V�[�g�Q���҈ꗗ�o��
 * @param  {string} date - ���t
 * @param  {string} eventTitle - �C�x���g�^�C�g��
 * @param  {array[]} members - �Q����
 *
 * @return {boolean} �o�͌���
 */
function outputSheet(date,eventTitle,members) {
  
  // �V�K�X�v���b�h�V�[�g�쐬��ID���擾
  var sheetId = createSpreadSheet(date,eventTitle);
  
  // �t�H�[�}�b�g�`����ݒ�
  designSheet(date,eventTitle,sheetId);
  
  // �Q���҂��s�ɑ}��
  insertRowMember(members,sheetId);
  
  return;  
}