const testsSettingsKey = 'testsSettingsKey'
const testsName = '����'

function showDialogTests() {
  var html = HtmlService.createTemplateFromFile('TestsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ������')
}

function saveTestsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(testsSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ������', err.message)
  }
}

function tests() {
  createSheet(testsName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try { 
    var testsSettings = JSON.parse(propertyService.getProperty(testsSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ������', err.message)
  }

  var studentsCount = addStudents(testsSettings.studentsDocUrl)
  var tests = SpreadsheetApp.getActiveSheet()

  tests.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('C1').setValue("�����").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('D1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('E1').setValue("����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (testsSettings.hasName == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("��������").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasDuration == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("������������ �����").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasStudentDuration == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("������������ �������").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasPresence == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
    tests.getRange(2, tests.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (testsSettings.hasComments == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasRemarks == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasFilter == true) {
    tests.getRange(1, 1, tests.getLastRow(), tests.getLastColumn()).createFilter()
  }

  importDocToSheet(testsName, testsSettings.studentsDocUrl)

  if (testsSettings.hasGroupSort == true) {
    tests.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (testsSettings.hasSurnameSort == true) {
    tests.getRange(2, 2, studentsCount, 2).sort(2)
  }

  tests.autoResizeColumns(1, tests.getLastColumn())
}
