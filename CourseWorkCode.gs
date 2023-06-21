const courseWorkSettingsKey = 'courseWorkSettingsKey'
const courseWorkName = 'Курсовая работа'

function showDialogCourseWork() {
  var html = HtmlService.createTemplateFromFile('CourseWorkPage').evaluate().setWidth(400).setHeight(370)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона курсовых работ')
}

function saveCourseWorkSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(courseWorkSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек курсовых работ', err.message)
  }
}

function courseWork() {
  createSheet(courseWorkName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var courseWorkSettings = JSON.parse(propertyService.getProperty(courseWorkSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек курсовых работ', err.message)
  }

  var studentsCount = addStudents(courseWorkSettings.studentsDocUrl)
  var courseWork = SpreadsheetApp.getActiveSheet()

  courseWork.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  courseWork.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  courseWork.getRange('C1').setValue("Название").mergeVertically().setHorizontalAlignment("center")
  
  if (courseWorkSettings.hasVariant == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  }
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Дата выдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Дата сдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Готовность").mergeVertically().setHorizontalAlignment("center")

  if (courseWorkSettings.hasReport == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Отчет").mergeVertically().setHorizontalAlignment("center")
    courseWork.getRange(2, courseWork.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  var courseWorkRule = SpreadsheetApp.newDataValidation().requireValueInList(["Зачет", "Незачет", "Не явился"], true).build()
  var range = courseWork.getRange(2, courseWork.getLastColumn(), studentsCount, 1)
  range.setDataValidation(courseWorkRule)
  range.setValue(null) 

  if (courseWorkSettings.hasComments == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseWorkSettings.hasRemarks == true) {
    courseWork.getRange(1, courseWork.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseWorkSettings.hasFilter == true) {
    courseWork.getRange(1, 1, courseWork.getLastRow(), courseWork.getLastColumn()).createFilter()
  }

  importDocToSheet(courseWorkName, courseWorkSettings.studentsDocUrl)

  if (courseWorkSettings.hasGroupSort == true) {
    courseWork.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (courseWorkSettings.hasSurnameSort == true) {
    courseWork.getRange(2, 2, studentsCount, 2).sort(2)
  }

  courseWork.autoResizeColumns(1, courseWork.getLastColumn())
}
