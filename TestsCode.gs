const testsSettingsKey = 'testsSettingsKey'
const testsName = 'Тест'

function showDialogTests() {
  var html = HtmlService.createTemplateFromFile('TestsPage').evaluate().setWidth(400).setHeight(400)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона тестов')
}

function saveTestsSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(testsSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек тестов', err.message)
  }
}

function tests() {
  createSheet(testsName)
  
  let propertyService = PropertiesService.getScriptProperties()
  try { 
    var testsSettings = JSON.parse(propertyService.getProperty(testsSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек тестов', err.message)
  }

  var studentsCount = addStudents(testsSettings.studentsDocUrl)
  var tests = SpreadsheetApp.getActiveSheet()

  tests.getRange('A1').setValue("Группа").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('B1').setValue("Студент").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('C1').setValue("Номер").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('D1').setValue("Вариант").mergeVertically().setHorizontalAlignment("center")
  tests.getRange('E1').setValue("Дата").mergeVertically().setHorizontalAlignment("center")
  setDate(studentsCount)

  if (testsSettings.hasName == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Название").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasDuration == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Длительность теста").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasStudentDuration == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Длительность решения").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasPresence == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Присутствие").mergeVertically().setHorizontalAlignment("center")
    tests.getRange(2, tests.getLastColumn(), studentsCount, 1).insertCheckboxes()
  }
  
  tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Оценка").mergeVertically().setHorizontalAlignment("center")
  setGrade(studentsCount)

  if (testsSettings.hasComments == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasRemarks == true) {
    tests.getRange(1, tests.getLastColumn() + 1, 1, 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (testsSettings.hasFilter == true) {
    tests.getRange(1, 1, tests.getLastRow(), tests.getLastColumn()).createFilter()
  }

  importDocToSheet(testsName, testsSettings.studentsDocUrl)

  if (testsSettings.hasGroupSort == true) {
    tests.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (testsSettings.hasSurnameSort == true) {
    tests.getRange(2, 2, studentsCount, 2).sort(2)
  }

  tests.autoResizeColumns(1, tests.getLastColumn())
}
