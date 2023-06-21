const controlsSettingsKey = 'controlsSettingsKey'
const controlsName = '����������� ������'

function showDialogControls() {
  var html = HtmlService.createTemplateFromFile('ControlsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ����������� �����')
}

function saveControlsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(controlsSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ����������� �����', err.message)
  }
}

function controls() {
  createSheet(controlsName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var controlsSettings = JSON.parse(propertyService.getProperty(controlsSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ����������� �����', err.message)
  }

  var studentsCount = addStudents(controlsSettings.studentsDocUrl)
  var controls = SpreadsheetApp.getActiveSheet()

  controls.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('C1').setValue("�����").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('D1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('E1').setValue("����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  if (controlsSettings.hasName == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("��������").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasDuration == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("������������ ������").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasStudentDuration == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("������������ �������").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasPresence == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
    controls.getRange(2, controls.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (controlsSettings.hasComments == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasRemarks == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasFilter == true) {
    controls.getRange(1, 1, controls.getLastRow(), controls.getLastColumn()).createFilter()
  }

  importDocToSheet(controlsName, controlsSettings.studentsDocUrl)

  if (controlsSettings.hasGroupSort == true) {
    controls.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (controlsSettings.hasSurnameSort == true) {
    controls.getRange(2, 2, studentsCount, 2).sort(2)
  }

  controls.autoResizeColumns(1, controls.getLastColumn())
}
