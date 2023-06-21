const controlsSettingsKey = 'controlsSettingsKey'
const controlsName = 'Контрольная работа'

function showDialogControls() {
  var html = HtmlService.createTemplateFromFile('ControlsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона контрольных работ')
}

function saveControlsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(controlsSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек контрольных работ', err.message)
  }
}

function controls() {
  createSheet(controlsName)
  let propertyService = PropertiesService.getScriptProperties()

  try {
    var controlsSettings = JSON.parse(propertyService.getProperty(controlsSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек контрольных работ', err.message)
  }

  var studentsCount = addStudents(controlsSettings.studentsDocUrl)
  var controls = SpreadsheetApp.getActiveSheet()

  controls.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('C1').setValue("Номер").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('D1').setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  controls.getRange('E1').setValue("Дата").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)
  
  if (controlsSettings.hasName == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Название").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasDuration == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Длительность работы").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasStudentDuration == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Длительность решения").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasPresence == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Присутствие").mergeVertically().setHorizontalAlignment("center")
    controls.getRange(2, controls.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (controlsSettings.hasComments == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasRemarks == true) {
    controls.getRange(1, controls.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (controlsSettings.hasFilter == true) {
    controls.getRange(1, 1, controls.getLastRow(), controls.getLastColumn()).createFilter()
  }

  importDocToSheet(controlsName, controlsSettings.studentsDocUrl)

  if (controlsSettings.hasGroupSort == true) {
    controls.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (controlsSettings.hasSurnameSort == true) {
    controls.getRange(2, 2, studentsCount, 2).sort(2)
  }

  controls.autoResizeColumns(1, controls.getLastColumn())
}
