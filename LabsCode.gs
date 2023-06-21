const labsSettingsKey = 'labsSettingsKey'
const labsName = '������������ ������'

function showDialogLabs() {
  var html = HtmlService.createTemplateFromFile('LabsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ������������ �����')
}

function saveLabsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(labsSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ������������ �����', err.message)
  }
}

function labs() {
  createSheet(labsName)
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var labsSettings = JSON.parse(propertyService.getProperty(labsSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ������������ �����', err.message)
  }

  var studentsCount = addStudents(labsSettings.studentsDocUrl)
  var lab = SpreadsheetApp.getActiveSheet()
  var countParams = 3;
  if (labsSettings.hasReport == true) {
    countParams = countParams + 1
  }

  if (labsSettings.number == "") {
    labsSettings.number = 1
  }

  lab.getRange('A1:A2').setValue("������").mergeVertically().setHorizontalAlignment("center")
  lab.getRange('B1:B2').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  lab.getRange(1, 3, 1, labsSettings.number * countParams).setValue("������������ ������").mergeAcross().setHorizontalAlignment("center")

  var numRow = 2
  var dateRow = 3  
  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  for (let i = 0; i < labsSettings.number; i++) {
    let currentColumn = 3 + i * countParams

    let currentLabRange = lab.getRange(numRow, currentColumn, 1, countParams)
    currentLabRange.setValue(i + 1).mergeAcross().setHorizontalAlignment("center")
    if (labsSettings.hasName == true) {
      currentLabRange.setNote("�������� ������������ ������ " + (i + 1) + ":" + '\n')
    }

    lab.getRange(dateRow, currentColumn).setValue("�������").setHorizontalAlignment("center")

    let givenColumn = currentColumn + 1
    lab.getRange(dateRow, givenColumn).setValue("������").setHorizontalAlignment("center")
    for (let k = 0; k < studentsCount; k++) {
      lab.getRange(4 + k, givenColumn).setValue(currentDate).setHorizontalAlignment("center")
    }

    let doneColumn = givenColumn + 1
    lab.getRange(dateRow, doneColumn).setValue("�����").setHorizontalAlignment("center")
    for (let k = 0; k < studentsCount; k++) {
      lab.getRange(4 + k, doneColumn).setValue(currentDate).setHorizontalAlignment("center")
    }

    if (labsSettings.hasReport == true) {
      let reportColumn = doneColumn + 1
      lab.getRange(dateRow, reportColumn).setValue("�����").setHorizontalAlignment("center")
      for (let k = 0; k < studentsCount; k++) {
        lab.getRange(4 + k, reportColumn).insertCheckboxes()
      }
    }
  }

  lab.getRange(1, lab.getLastColumn() + 1, 2, 1).setValue("������").mergeVertically().setHorizontalAlignment("center")
  
  var labRule = SpreadsheetApp.newDataValidation().requireValueInList(["�����", "�������", "�������", "������", "�����������������", "�������������������", "�������", "�� �������", "�� ������"], true).build()
  var range = lab.getRange(4, lab.getLastColumn(), studentsCount, 1)
  range.setDataValidation(labRule)
  range.setValue(null) 

  if (labsSettings.hasComments == true) {
    lab.getRange(1, lab.getLastColumn() + 1, 2, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (labsSettings.hasRemarks == true) {
    lab.getRange(1, lab.getLastColumn() + 1, 2, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (labsSettings.hasFilter == true) {
    lab.getRange(3, 1, lab.getLastRow() - 2, lab.getLastColumn()).createFilter()
  }

  importDocToSheet(labsName, labsSettings.studentsDocUrl)

  if (labsSettings.hasGroupSort == true) {
    lab.getRange(4, 1, studentsCount, 1).sort(1)
  }

  if (labsSettings.hasSurnameSort == true) {
    lab.getRange(4, 2, studentsCount, 2).sort(2)
  }

  lab.autoResizeColumns(1, lab.getLastColumn())

}  
