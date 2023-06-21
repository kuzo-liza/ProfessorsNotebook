const vkrSettingsKey = 'vkrSettingsKey'
const vkrFirstName = '���������� ���'
const vkrSecondName = '������ ���'

function showDialogVKRCertification() {
  var html = HtmlService.createTemplateFromFile('VKRPage').evaluate().setWidth(400).setHeight(540)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ���������� � ������ ���')
}

function saveVKRSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(vkrSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ���', err.message)
  }
}

function vkrCertificationAndGraduation() {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var vkrSettings = JSON.parse(propertyService.getProperty(vkrSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ���', err.message)
    return
  }
  vkrCertification(vkrSettings)
  vkrGraduation(vkrSettings)
}

function vkrCertification(vkrSettings) {
  createSheet(vkrFirstName)

  var studentsCount = addStudents(vkrSettings.studentsDocUrl)
  var certification = SpreadsheetApp.getActiveSheet()

  certification.getRange('A1').setValue("������")
  certification.getRange('B1').setValue("�������")
  certification.getRange('C1').setValue("����� ����������")

  var vkrRule = SpreadsheetApp.newDataValidation().requireValueInList(["1", "2", "3", "4"], true).build()
  var first = certification.getRange(2, certification.getLastColumn(), studentsCount, 1)
  var second = certification.getRange(studentsCount + 2, certification.getLastColumn(), studentsCount, 1)
  var third = certification.getRange(studentsCount * 2 + 2, certification.getLastColumn(), studentsCount, 1)
  var fourth = certification.getRange(studentsCount * 3 + 2, certification.getLastColumn(), studentsCount, 1)

  first.setDataValidation(vkrRule)
  second.setDataValidation(vkrRule)
  third.setDataValidation(vkrRule)
  fourth.setDataValidation(vkrRule)

  first.setValue(1) 
  second.setValue(2) 
  third.setValue(3) 
  fourth.setValue(4)

  certification.getRange(1, certification.getLastColumn() + 1).setValue("���� ����������")

  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  for (let i = 1; i <= studentsCount * 4; i++) {
    certification.getRange(1 + i, 4).setValue(currentDate)
  }  

  certification.getRange(1, certification.getLastColumn() + 1).setValue("���� ���")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("������������")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("�����������")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("������ �� ����������")

  var ruleAttestationPoints = SpreadsheetApp.newDataValidation().requireValueInList(["�������", "�� �������", "�� ������"], true).build()
  var rangeAttestation = certification.getRange(2, certification.getLastColumn(), studentsCount * 4, 1)
  rangeAttestation.setDataValidation(ruleAttestationPoints)
  rangeAttestation.setValue(null) 
  
  if (vkrSettings.hasAttendance == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
    certification.getRange(2, certification.getLastColumn(), studentsCount * 4, 1).insertCheckboxes()
  }

  if (vkrSettings.hasCommentsAttestation == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasRemarksAttestation == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  importDocToSheet(vkrFirstName, vkrSettings.studentsDocUrl)

  if (vkrSettings.hasGroupSort == true) {
    certification.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (vkrSettings.hasSurnameSort == true) {
    certification.getRange(2, 2, studentsCount, 2).sort(2)
  }

 if (vkrSettings.hasAttendanceNumberSort == true) {
    certification.getRange(2, 3, studentsCount, 3).sort(3)
  }

  if (vkrSettings.hasFilter == true) {
    certification.getRange(1, 1, certification.getLastRow(), certification.getLastColumn()).createFilter()
  }
}

function vkrGraduation(vkrSettings) {
  createSheet(vkrSecondName)
  var graduation = SpreadsheetApp.getActiveSheet()

  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  var studentsCount = addStudents(vkrSettings.studentsDocUrl)

  graduation.getRange('A1').setValue("������")
  graduation.getRange('B1').setValue("�������")
  graduation.getRange('C1').setValue("���� ���")
  graduation.getRange('D1').setValue("������������")
  graduation.getRange('E1').setValue("�����������")
  graduation.getRange('F1').setValue("���������")
  graduation.getRange('G1').setValue("�������������")
  graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1).insertCheckboxes()
  graduation.getRange(1, graduation.getLastColumn() + 1).setValue("���� ������")

  for (let i = 1; i <= studentsCount; i++) {
    graduation.getRange(1 + i, 8).setValue(currentDate)
  }  

  if (vkrSettings.hasPointsSupervisor == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("������ ������������").mergeVertically().setHorizontalAlignment("center")

    var ruleSupervisorPoints = SpreadsheetApp.newDataValidation().requireValueInList(["�������", "������", "�����������������", "�������������������", "�� �������", "�� ������"], true).build()
    var rangePoinsS = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangePoinsS.setDataValidation(ruleSupervisorPoints)
    rangePoinsS.setValue(null) 
  }

  if (vkrSettings.hasPointsReviewer == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("������ ����������").mergeVertically().setHorizontalAlignment("center")

    var ruleReviewerPoints = SpreadsheetApp.newDataValidation().requireValueInList(["�������", "������", "�����������������", "�������������������", "�� �������", "�� ������"], true).build()
    var rangePoinsR = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangePoinsR.setDataValidation(ruleReviewerPoints)
    rangePoinsR.setValue(null) 
  }

  if (vkrSettings.hasOriginality == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("��������������").mergeVertically().setHorizontalAlignment("center")
    graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1).setValue(0)
  }

  if (vkrSettings.hasResult == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")

    var ruleResult = SpreadsheetApp.newDataValidation().requireValueInList(["�������", "������", "�����������������", "�������������������", "�� �������", "�� ������"], true).build()
    var rangeResult = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangeResult.setDataValidation(ruleResult)
    rangeResult.setValue(null) 
  }

  if (vkrSettings.hasRemarks == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasComments == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasFilter == true) {
    graduation.getRange(1, 1, graduation.getLastRow(), graduation.getLastColumn()).createFilter()
  }

  importDocToSheet(vkrSecondName, vkrSettings.studentsDocUrl)

  if (vkrSettings.hasGroupSort == true) {
    graduation.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (vkrSettings.hasSurnameSort == true) {
    graduation.getRange(2, 2, studentsCount, 2).sort(2)
  }
  
  graduation.autoResizeColumns(1, graduation.getLastColumn())
}
