const reportsSettingsKey = 'reportsSettingsKey'
const reportsName = 'Доклады'

function showDialogReports() {
  var html = HtmlService.createTemplateFromFile('ReportsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона докладов')
}

function saveReportsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(reportsSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек докладов', err.message)
  }
}

function reports() {
  createSheet(reportsName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var reportsSettings = JSON.parse(propertyService.getProperty(reportsSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек докладов', err.message)
  }

  var studentsCount = addStudents(reportsSettings.studentsDocUrl)
  var reports = SpreadsheetApp.getActiveSheet()

  reports.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  reports.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  reports.getRange('C1').setValue("Название").mergeVertically().setHorizontalAlignment("center")
  reports.getRange('D1').setValue("Дата выступления").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  if (reportsSettings.hasNumber == true) {
    reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Номер выступления").mergeVertically().setHorizontalAlignment("center")
  }

  if (reportsSettings.hasDuration == true) {
    reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Длительность доклада").mergeVertically().setHorizontalAlignment("center")
  }
  
  reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (reportsSettings.hasPresentation == "true") {
    reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Оценка за презентацию").mergeVertically().setHorizontalAlignment("center")
    let range = reports.getRange(2, reports.getLastColumn(), studentsCount, 1)
    range.setDataValidation(reportsRule)
    range.setValue(null)
  }

  if (reportsSettings.hasComments == true) {
    reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (reportsSettings.hasRemarks == true) {
    reports.getRange(1, reports.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (reportsSettings.hasFilter == true) {
    reports.getRange(1, 1, reports.getLastRow(), reports.getLastColumn()).createFilter()
  }

  importDocToSheet(reportsName, reportsSettings.studentsDocUrl)

  if (reportsSettings.hasGroupSort == true) {
    reports.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (reportsSettings.hasSurnameSort == true) {
    reports.getRange(2, 2, studentsCount, 2).sort(2)
  }

  reports.autoResizeColumns(1, reports.getLastColumn())
}
