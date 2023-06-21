const lectionsSettingsKey = 'lectionsSettingsKey'
const lectionsName = 'Лекции'

function showDialogLections() {
  var html = HtmlService.createTemplateFromFile('LectionsPage').evaluate().setWidth(400).setHeight(460)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона лекций')
}

function saveLectionsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(lectionsSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек лекций', err.message)
  }
}

function lections() {
  createSheet(lectionsName)
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var lectionsSettings = JSON.parse(propertyService.getProperty(lectionsSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек лекций', err.message)
  }

  var studentsCount = addStudents(lectionsSettings.studentsDocUrl)
  var lection = SpreadsheetApp.getActiveSheet()

  if (lectionsSettings.number == "") {
    lectionsSettings.number = 1
  }

  lection.getRange('A1:A2').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  lection.getRange('B1:B2').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  lection.getRange(1, 3, 1, lectionsSettings.number).setValue("Дата лекции и номер").mergeAcross().setHorizontalAlignment("center")

  var numRow = 2
  var dateRow = 3
  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  for (let i = 1; i <= lectionsSettings.number; i++) {
    lection.getRange(numRow, 2 + i).setValue(i).setHorizontalAlignment("center")
    lection.getRange(dateRow, 2 + i).setValue(currentDate)
    if (lectionsSettings.hasName == true) {
      lection.getRange(numRow, 2 + i).setNote("Название лекции " + i + ":" + '\n')
    }
    if (lectionsSettings.hasFormat == true) {
      var note = lection.getRange(numRow, 2 + i).getNote()
      lection.getRange(numRow, 2 + i).setNote(note + "Формат лекции: " + '\n')
    }
    if (lectionsSettings.hasDuration == true) {
      var note = lection.getRange(numRow, 2 + i).getNote()
      lection.getRange(numRow, 2 + i).setNote(note + "Длительность лекции: ")
    }
  }  

  if (lectionsSettings.hasStudentDuration == true) {
    lection.getRange(4, 3, studentsCount, lectionsSettings.number).setNote("Длительность посещения студентом: ")
  }

  lection.getRange(4, 3, studentsCount, lectionsSettings.number).insertCheckboxes()
  lection.getRange(1, lection.getLastColumn() + 1, 2, 1).setValue("Общая посещаемость").mergeVertically().setHorizontalAlignment("center")

  var formula
  var lastColumn
  var address
  var cell

  for (var k = 4; k < (Number(studentsCount) + 4); k++) {
    lastColumn = lection.getLastColumn() - 1
    address = "=ADDRESS("+ k + "; " + lastColumn + "; 4)"
    lection.getRange(100,2).setValue(address).setFontColor('white')
    cell = lection.getRange(100,2).getValue()
    formula = "=TRUNC((COUNTIF(C" + k + ":" + cell + "; \"TRUE\"))*100/" + lectionsSettings.number + ")"
    lection.getRange(k, lastColumn + 1).setValue(formula)
  }

  lection.getRange(100,2).setValue("")
  lection.getRange(4, lection.getLastColumn(), studentsCount, 1).setBackground('#FF6666')
  lection.getRange(1, lection.getLastColumn() + 1, 2, 1).setValue("Экзамен").mergeVertically().setHorizontalAlignment("center")
  
  var lectionsRule = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Допущен", "Не допущен", "Не явился"], true).build()
  var rangeExam = lection.getRange(4, lection.getLastColumn(), studentsCount, 1)
  rangeExam.setDataValidation(lectionsRule)
  rangeExam.setValue(null) // set initial value to "Неизвестно"

  if (lectionsSettings.hasComments == true) {
    lection.getRange(1, lection.getLastColumn() + 1, 2, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (lectionsSettings.hasRemarks == true) {
    lection.getRange(1, lection.getLastColumn() + 1, 2, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (lectionsSettings.hasFilter == true) {
    lection.getRange(3, 1, lection.getLastRow() - 2, lection.getLastColumn()).createFilter()
  }

  importDocToSheet(lectionsName, lectionsSettings.studentsDocUrl)

  if (lectionsSettings.hasGroupSort == true) {
    lection.getRange(4, 1, studentsCount, 1).sort(1)
  }

  if (lectionsSettings.hasSurnameSort == true) {
    lection.getRange(4, 2, studentsCount, 2).sort(2)
  }

  lection.autoResizeColumns(1, lection.getLastColumn())
}
