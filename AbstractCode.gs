const abstractSettingsKey = 'abstractSettingsKey'
const abstractName = 'Рефераты'

function showDialogAbstract() {
  var html = HtmlService.createTemplateFromFile('AbstractPage').evaluate().setWidth(400).setHeight(380)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона рефератов')
}

function saveAbstractSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(abstractSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек рефератов', err.message)
  }
}

function abstract() {
  createSheet(abstractName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var abstractSettings = JSON.parse(propertyService.getProperty(abstractSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек рефератов', err.message)
  }

  var studentsCount = addStudents(abstractSettings.studentsDocUrl)
  var abstract = SpreadsheetApp.getActiveSheet()
  
  abstract.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  abstract.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  abstract.getRange('C1').setValue("Название").mergeVertically().setHorizontalAlignment("center")

  if (abstractSettings.hasVariant == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  }
  
  abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Дата сдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (abstractSettings.hasDate == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Дата выдачи").mergeVertically().setHorizontalAlignment("center")
    setDate(studentsCount)
  }

  abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (abstractSettings.hasComments == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (abstractSettings.hasRemarks == true) {
    abstract.getRange(1, abstract.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
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
