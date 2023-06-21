const essaySettingsKey = 'essaySettingsKey'
const essayName = 'Эссе'

function showDialogEssay() {
  var html = HtmlService.createTemplateFromFile('EssayPage').evaluate().setWidth(400).setHeight(380)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона эссе')
}

function saveEssaySettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(essaySettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек эссе', err.message)
  }
}

function essay() {
  createSheet(essayName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var essaySettings = JSON.parse(propertyService.getProperty(essaySettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек эссе', err.message)
  }

  var studentsCount = addStudents(essaySettings.studentsDocUrl)
  var essay = SpreadsheetApp.getActiveSheet()

  essay.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  essay.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  essay.getRange('C1').setValue("Название").mergeVertically().setHorizontalAlignment("center")

  if (essaySettings.hasVariant == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  }

  essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Дата сдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (essaySettings.hasDate == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Дата выдачи").mergeVertically().setHorizontalAlignment("center")
     setDate(studentsCount)
  }

  essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (essaySettings.hasComments == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (essaySettings.hasRemarks == true) {
    essay.getRange(1, essay.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
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
