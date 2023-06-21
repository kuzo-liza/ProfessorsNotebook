const homeworksSettingsKey = 'homeworksSettingsKey'
const homeworksName = '�������� �������'

function showDialogHomeworks() {
  var html = HtmlService.createTemplateFromFile('HomeworksPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� �������� �������')
}

function saveHomeworksSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(homeworksSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� �������� �������', err.message)
  }
}

function homeworks() {
  createSheet(homeworksName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var homeworksSettings = JSON.parse(propertyService.getProperty(homeworksSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� �������� �������', err.message)
  }

  var studentsCount = addStudents(homeworksSettings.studentsDocUrl)
  var homeworks = SpreadsheetApp.getActiveSheet()

  homeworks.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  homeworks.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  homeworks.getRange('C1').setValue("�����").mergeVertically().setHorizontalAlignment("center")
  homeworks.getRange('D1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  homeworks.getRange('E1').setValue("���� �����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (homeworksSettings.hasName == true) {
    homeworks.getRange(1, homeworks.getLastColumn() + 1, 1, 1).setValue("��������").mergeVertically().setHorizontalAlignment("center")
  }

  if (homeworksSettings.hasBlackboardAnswer == true) {
    homeworks.getRange(1, homeworks.getLastColumn() + 1, 1, 1).setValue("����� � �����").mergeVertically().setHorizontalAlignment("center")
    homeworks.getRange(2, homeworks.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  homeworks.getRange(1, homeworks.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (homeworksSettings.hasComments == true) {
    homeworks.getRange(1, homeworks.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (homeworksSettings.hasRemarks == true) {
    homeworks.getRange(1, homeworks.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (homeworksSettings.hasFilter == true) {
    homeworks.getRange(1, 1, homeworks.getLastRow(), homeworks.getLastColumn()).createFilter()
  }

  importDocToSheet(homeworksName, homeworksSettings.studentsDocUrl)

  if (homeworksSettings.hasGroupSort == true) {
    homeworks.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (homeworksSettings.hasSurnameSort == true) {
    homeworks.getRange(2, 2, studentsCount, 2).sort(2)
  }

  homeworks.autoResizeColumns(1, homeworks.getLastColumn())
}
