const courseProjectSettingsKey = 'courseProjectSettingsKey'
const courseProjectName = 'Курсовой проект'

function showDialogCourseProject() {
  var html = HtmlService.createTemplateFromFile('CourseProjectPage').evaluate().setWidth(400).setHeight(370)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона курсовых проектов')
}

function saveCourseProjectSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(courseProjectSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек курсовых проектов', err.message)
  }
}

function courseProject() {
  createSheet(courseProjectName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var courseProjectSettings = JSON.parse(propertyService.getProperty(courseProjectSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек курсовых проектов', err.message)
  }

  var studentsCount = addStudents(courseProjectSettings.studentsDocUrl)
  var courseProject = SpreadsheetApp.getActiveSheet()

  courseProject.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  courseProject.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  courseProject.getRange('C1').setValue("Название").mergeVertically().setHorizontalAlignment("center")
  
  if (courseProjectSettings.hasVariant == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  }
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Дата выдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Дата сдачи").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Готовность").mergeVertically().setHorizontalAlignment("center")

  if (courseProjectSettings.hasReport == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Отчет").mergeVertically().setHorizontalAlignment("center")
    courseProject.getRange(2, courseProject.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  var courseProjectRule = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Не явился"], true).build()
  var range = courseProject.getRange(2, courseProject.getLastColumn(), studentsCount, 1)
  range.setDataValidation(courseProjectRule)
  range.setValue(null) 

  if (courseProjectSettings.hasComments == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseProjectSettings.hasRemarks == true) {
    courseProject.getRange(1, courseProject.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (courseProjectSettings.hasFilter == true) {
    courseProject.getRange(1, 1, courseProject.getLastRow(), courseProject.getLastColumn()).createFilter()
  }

  importDocToSheet(courseProjectName, courseProjectSettings.studentsDocUrl)

  if (courseProjectSettings.hasGroupSort == true) {
    courseProject.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (courseProjectSettings.hasSurnameSort == true) {
    courseProject.getRange(2, 2, studentsCount, 2).sort(2)
  }

  courseProject.autoResizeColumns(1, courseProject.getLastColumn())
}
