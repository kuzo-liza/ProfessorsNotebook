const abstractSettingsKey = 'abstractSettingsKey'
const abstractName = '��������'

function showDialogAbstract() {
  var html = HtmlService.createTemplateFromFile('AbstractPage').evaluate().setWidth(400).setHeight(380)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ���������')
}

function saveAbstractSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(abstractSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ���������', err.message)
  }
}

function abstract() {
  createSheet(abstractName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var abstractSettings = JSON.parse(propertyService.getProperty(abstractSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ���������', err.message)
  }

  var studentsCount = addStudents(abstractSettings.studentsDocUrl)
  var abstract = SpreadsheetApp.getActiveSheet()
  
  abstract.getRange('A1').setValue("������").mergeVertically().setHorizontalAlignment("center")
  abstract.getRange('B1').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  abstract.getRange('C1').setValue("��������").mergeVertically().setHorizontalAlignment("center")

  if (abstractSettings.hasVariant == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("�������").mergeVertically().setHorizontalAlignment("center")
  }
  
  abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("���� �����").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (abstractSettings.hasDate == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("���� ������").mergeVertically().setHorizontalAlignment("center")
    setDate(studentsCount)
  }

  abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (abstractSettings.hasComments == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (abstractSettings.hasRemarks == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (abstractSettings.hasFilter == true) {
    abstract.getRange(1, 1, abstract.getLastRow(), abstract.getLastColumn()).createFilter()
  }

  importDocToSheet(abstractName, abstractSettings.studentsDocUrl)

  if (abstractSettings.hasGroupSort == true) {
    abstract.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (abstractSettings.hasSurnameSort == true) {
    abstract.getRange(2, 2, studentsCount, 2).sort(2)
  }

  abstract.autoResizeColumns(1, abstract.getLastColumn())
}
