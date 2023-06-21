const practicesSettingsKey = 'practicesSettingsKey'
const practicesName = '������������ �������'

function showDialogPractices() {
  var html = HtmlService.createTemplateFromFile('PracticesPage').evaluate().setWidth(400).setHeight(480)
  SpreadsheetApp.getUi().showModalDialog(html, '��������� ������� ������������ �������')
}

function savePracticesSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(practicesSettingsKey, formSettings)
  } catch(err) {
    console.log('�� ���������� �������� ������ �������� ������������ �������', err.message)
  }
}

function practices() {
  createSheet(practicesName)
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var practicesSettings = JSON.parse(propertyService.getProperty(practicesSettingsKey))
  } catch (err) {
    console.log('�� ���������� ������� ������ �� �������� ������������ �������', err.message)
  }

  var studentsCount = addStudents(practicesSettings.studentsDocUrl)
  var practices = SpreadsheetApp.getActiveSheet()

  if (practicesSettings.number == "") {
    practicesSettings.number = 1
  }

  practices.getRange('A1:A2').setValue("������").mergeVertically().setHorizontalAlignment("center")
  practices.getRange('B1:B2').setValue("�������").mergeVertically().setHorizontalAlignment("center")
  practices.getRange(1, 3, 1, practicesSettings.number).setValue("���� � ����� ������������� �������").mergeAcross().setHorizontalAlignment("center")

  var numRow = 2
  var dateRow = 3  
  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  for (let i = 1; i <= practicesSettings.number; i++) {
    practices.getRange(numRow, 2 + i).setValue(i).setHorizontalAlignment("center")
    practices.getRange(dateRow, 2 + i).setValue(currentDate)
    if (practicesSettings.hasName == true) {
      practices.getRange(numRow, 2 + i).setNote("�������� ������������� ������� " + i + ":" + '\n')
    }
    if (practicesSettings.hasFormat == true) {
      var note = practices.getRange(numRow, 2 + i).getNote()
      practices.getRange(numRow, 2 + i).setNote(note + "������ �������: " + '\n')
    }
    if (practicesSettings.hasDuration == true) {
      var note = practices.getRange(numRow, 2 + i).getNote()
      practices.getRange(numRow, 2 + i).setNote(note + "������������ �������: ")
    }
  }  

  if (practicesSettings.hasStudentDuration == true) {
    practices.getRange(4, 3, studentsCount, practicesSettings.number).setNote("������������ ��������� ���������: ")
  }

  practices.getRange(4, 3, studentsCount, practicesSettings.number).insertCheckboxes()
  practices.getRange(1, practices.getLastColumn() + 1, 2, 1).setValue("����� ������������").mergeVertically().setHorizontalAlignment("center")

  var formula
  var lastColumn
  var address
  var cell 

  for (var k = 4; k < (Number(studentsCount) + 4); k++) {
    lastColumn = practices.getLastColumn() - 1
    address = "=ADDRESS("+ k + "; " + lastColumn + "; 4)"
    practices.getRange(100,2).setValue(address).setFontColor('white')
    cell = practices.getRange(100,2).getValue()
    formula = "=TRUNC((COUNTIF(C" + k + ":" + cell + "; \"TRUE\"))*100/" + practicesSettings.number + ")"
    practices.getRange(k, lastColumn + 1).setValue(formula)
  }

  practices.getRange(100,2).setValue("")
  practices.getRange(4, practices.getLastColumn(), studentsCount, 1).setBackground('#FF6666')
  practices.getRange(1, practices.getLastColumn() + 1, 2, 1).setValue("�����").mergeVertically().setHorizontalAlignment("center")
  
  var practicesRule = SpreadsheetApp.newDataValidation().requireValueInList(["�����", "�������", "�������", "������", "�����������������", "�������������������", "�������", "�� �������", "�� ������"], true).build()
  var range = practices.getRange(4, practices.getLastColumn(), studentsCount, 1)
  range.setDataValidation(practicesRule)
  range.setValue(null) 

  if (practicesSettings.hasComments == true) {
    practices.getRange(1, practices.getLastColumn() + 1, 2, 1).setValue("�����������").mergeVertically().setHorizontalAlignment("center")
  }

  if (practicesSettings.hasRemarks == true) {
    practices.getRange(1, practices.getLastColumn() + 1, 2, 1).setValue("���������").mergeVertically().setHorizontalAlignment("center")
  }

  if (practicesSettings.hasFilter == true) {
    practices.getRange(3, 1, practices.getLastRow() - 2, practices.getLastColumn()).createFilter()
  }

  importDocToSheet(practicesName, practicesSettings.studentsDocUrl)

  if (practicesSettings.hasGroupSort == true) {
    practices.getRange(4, 1, studentsCount, 1).sort(1)
  }

  if (practicesSettings.hasSurnameSort == true) {
    practices.getRange(4, 2, studentsCount, 2).sort(2)
  }

  practices.autoResizeColumns(1, practices.getLastColumn())
}
