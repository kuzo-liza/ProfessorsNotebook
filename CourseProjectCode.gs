const courseProjectSettingsKey = 'courseProjectSettingsKey'
const courseProjectName = '�������� ������'

function showDialogCourseProject() {
  var html = HtmlService.createTemplateFromFile('CourseProjectPage').evaluate().setWidth(400).setHeight(370)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� �������� ��������')
}

function saveCourseProjectSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(courseProjectSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� �������� ��������', err.message)
  }
}

function courseProject() {
  createSheet(courseProjectName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var courseProjectSettings = JSON.parse(propertyService.getProperty(courseProjectSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� �������� ��������', err.message)
  }

  var studentsCount = addStudents(courseProjectSettings.studentsDocUrl)
  var courseProject = SpreadsheetApp.getActiveSheet()

  courseProject.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  courseProject.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  courseProject.getRange('C1').setValue("��������").mergeVertically().setHorizontalAlignment("center")
  
  if (courseProjectSettings.hasVariant == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("�������").mergeVertically().setHorizontalAlignment("center")
  }
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("���� ������").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("���� �����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("����������").mergeVertically().setHorizontalAlignment("center")

  if (courseProjectSettings.hasReport == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("�����").mergeVertically().setHorizontalAlignment("center")
    courseProject.getRange(2, courseProject.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  var courseProjectRule = SpreadsheetApp.newDataValidation().requireValueInList(["�������", "������", "�����������������", "�������������������", "�� ������"], true).build()
  var range = courseProject.getRange(2, courseProject.getLastColumn(), studentsCount, 1)
  range.setDataValidation(courseProjectRule)
  range.setValue(null) 

  if (courseProjectSettings.hasComments == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseProjectSettings.hasRemarks == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseProjectSettings.hasFilter == true) {
    courseProject.getRange(1, 1, courseProject.getLastRow(), courseProject.getLastColumn()).createFilter()
  }

  importDocToSheet(courseProjectName, courseProjectSettings.studentsDocUrl)

  if (courseProjectSettings.hasGroupSort == true) {
    courseProject.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (courseProjectSettings.hasSurnameSort == true) {
    courseProject.getRange(2, 2, studentsCount, 2).sort(2)
  }

  courseProject.autoResizeColumns(1, courseProject.getLastColumn())
}
