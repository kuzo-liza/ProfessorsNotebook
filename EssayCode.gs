const essaySettingsKey = 'essaySettingsKey'
const essayName = '����'

function showDialogEssay() {
  var html = HtmlService.createTemplateFromFile('EssayPage').evaluate().setWidth(400).setHeight(380)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ����')
}

function saveEssaySettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(essaySettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ����', err.message)
  }
}

function essay() {
  createSheet(essayName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var essaySettings = JSON.parse(propertyService.getProperty(essaySettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ����', err.message)
  }

  var studentsCount = addStudents(essaySettings.studentsDocUrl)
  var essay = SpreadsheetApp.getActiveSheet()

  essay.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  essay.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  essay.getRange('C1').setValue("��������").mergeVertically().setHorizontalAlignment("center")

  if (essaySettings.hasVariant == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("�������").mergeVertically().setHorizontalAlignment("center")
  }

  essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("���� �����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (essaySettings.hasDate == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("���� ������").mergeVertically().setHorizontalAlignment("center")
     setDate(studentsCount)
  }

  essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (essaySettings.hasComments == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (essaySettings.hasRemarks == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (essaySettings.hasFilter == true) {
    essay.getRange(1, 1, essay.getLastRow(), essay.getLastColumn()).createFilter()
  }

  importDocToSheet(essayName, essaySettings.studentsDocUrl)

  if (essaySettings.hasGroupSort == true) {
    essay.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (essaySettings.hasSurnameSort == true) {
    essay.getRange(2, 2, studentsCount, 2).sort(2)
  }

  essay.autoResizeColumns(1, essay.getLastColumn())
}
