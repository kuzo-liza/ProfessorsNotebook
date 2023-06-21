const edPracticeSettingsKey = 'edPracticeSettingsKey'
const edPracticeName = 'Практика'

function showDialogEdPractice() {
  var html = HtmlService.createTemplateFromFile('EdPracticePage').evaluate().setWidth(400).setHeight(370)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона практики')
}

function saveEdPracticeSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(edPracticeSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек практики', err.message)
  }
}

function edPractice() {
  createSheet(edPracticeName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var edPracticeSettings = JSON.parse(propertyService.getProperty(edPracticeSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек практики', err.message)
  }

  var studentsCount = addStudents(edPracticeSettings.studentsDocUrl)
  var edPractice = SpreadsheetApp.getActiveSheet()

  edPractice.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  edPractice.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  edPractice.getRange('C1').setValue("Вид практики").mergeVertically().setHorizontalAlignment("center")
  
  var edPracticeRule = SpreadsheetApp.newDataValidation().requireValueInList(["Научно-исследовательская", "Производственная", "Преддипломная"], true).build()
  var range = edPractice.getRange(2, edPractice.getLastColumn(), studentsCount, 1)
  range.setDataValidation(edPracticeRule)
  range.setValue(null) 

  edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Дата начала").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Дата окончания").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Отчет").mergeVertically().setHorizontalAlignment("center")
  edPractice.getRange(2, edPractice.getLastColumn(), studentsCount, 1).insertCheckboxes()

  if (edPracticeSettings.hasCurator == true) {
    edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Куратор").mergeVertically().setHorizontalAlignment("center")
  }

  if (edPracticeSettings.hasTheme == true) {
    edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Тема").mergeVertically().setHorizontalAlignment("center")
  }

  edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (edPracticeSettings.hasComments == true) {
    edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (edPracticeSettings.hasRemarks == true) {
    edPractice.getRange(1, edPractice.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (edPracticeSettings.hasFilter == true) {
    edPractice.getRange(1, 1, edPractice.getLastRow(), edPractice.getLastColumn()).createFilter()
  }

  importDocToSheet(edPracticeName, edPracticeSettings.studentsDocUrl)

  if (edPracticeSettings.hasGroupSort == true) {
    edPractice.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (edPracticeSettings.hasSurnameSort == true) {
    edPractice.getRange(2, 2, studentsCount, 2).sort(2)
  }

  edPractice.autoResizeColumns(1, edPractice.getLastColumn())
}
