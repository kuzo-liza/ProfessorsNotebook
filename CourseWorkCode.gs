const courseWorkSettingsKey = 'courseWorkSettingsKey'
const courseWorkName = '�������� ������'

function showDialogCourseWork() {
  var html = HtmlService.createTemplateFromFile('CourseWorkPage').evaluate().setWidth(400).setHeight(370)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� �������� �����')
}

function saveCourseWorkSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(courseWorkSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� �������� �����', err.message)
  }
}

function courseWork() {
  createSheet(courseWorkName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var courseWorkSettings = JSON.parse(propertyService.getProperty(courseWorkSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� �������� �����', err.message)
  }

  var studentsCount = addStudents(courseWorkSettings.studentsDocUrl)
  var courseWork = SpreadsheetApp.getActiveSheet()

  courseWork.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  courseWork.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  courseWork.getRange('C1').setValue("��������").mergeVertically().setHorizontalAlignment("center")
  
  if (courseWorkSettings.hasVariant == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("�������").mergeVertically().setHorizontalAlignment("center")
  }
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("���� ������").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("���� �����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("����������").mergeVertically().setHorizontalAlignment("center")

  if (courseWorkSettings.hasReport == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("�����").mergeVertically().setHorizontalAlignment("center")
    courseWork.getRange(2, courseWork.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  var courseWorkRule = SpreadsheetApp.newDataValidation().requireValueInList(["�����", "�������", "�� ������"], true).build()
  var range = courseWork.getRange(2, courseWork.getLastColumn(), studentsCount, 1)
  range.setDataValidation(courseWorkRule)
  range.setValue(null) 

  if (courseWorkSettings.hasComments == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseWorkSettings.hasRemarks == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseWorkSettings.hasFilter == true) {
    courseWork.getRange(1, 1, courseWork.getLastRow(), courseWork.getLastColumn()).createFilter()
  }

  importDocToSheet(courseWorkName, courseWorkSettings.studentsDocUrl)

  if (courseWorkSettings.hasGroupSort == true) {
    courseWork.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (courseWorkSettings.hasSurnameSort == true) {
    courseWork.getRange(2, 2, studentsCount, 2).sort(2)
  }

  courseWork.autoResizeColumns(1, courseWork.getLastColumn())
}
