const vkrSettingsKey = 'vkrSettingsKey'
const vkrFirstName = 'Аттестации ВКР'
const vkrSecondName = 'Защита ВКР'

function showDialogVKRCertification() {
  var html = HtmlService.createTemplateFromFile('VKRPage').evaluate().setWidth(400).setHeight(540)
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка шаблона аттестации и защиты ВКР')
}

function saveVKRSettings(formSettings) {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    propertyService.setProperty(vkrSettingsKey, formSettings)
  } catch(err) {
    console.log('Не получилось записать данные настроек ВКР', err.message)
  }
}

function vkrCertificationAndGraduation() {
  let propertyService = PropertiesService.getScriptProperties()
  try {
    var vkrSettings = JSON.parse(propertyService.getProperty(vkrSettingsKey))
  } catch (err) {
    console.log('Не получилось достать данные из настроек ВКР', err.message)
    return
  }
  vkrCertification(vkrSettings)
  vkrGraduation(vkrSettings)
}

function vkrCertification(vkrSettings) {
  createSheet(vkrFirstName)

  var studentsCount = addStudents(vkrSettings.studentsDocUrl)
  var certification = SpreadsheetApp.getActiveSheet()

  certification.getRange('A1').setValue("Группа")
  certification.getRange('B1').setValue("Студент")
  certification.getRange('C1').setValue("Номер аттестации")

  var vkrRule = SpreadsheetApp.newDataValidation().requireValueInList(["1", "2", "3", "4"], true).build()
  var first = certification.getRange(2, certification.getLastColumn(), studentsCount, 1)
  var second = certification.getRange(studentsCount + 2, certification.getLastColumn(), studentsCount, 1)
  var third = certification.getRange(studentsCount * 2 + 2, certification.getLastColumn(), studentsCount, 1)
  var fourth = certification.getRange(studentsCount * 3 + 2, certification.getLastColumn(), studentsCount, 1)

  first.setDataValidation(vkrRule)
  second.setDataValidation(vkrRule)
  third.setDataValidation(vkrRule)
  fourth.setDataValidation(vkrRule)

  first.setValue(1) 
  second.setValue(2) 
  third.setValue(3) 
  fourth.setValue(4)

  certification.getRange(1, certification.getLastColumn() + 1).setValue("Дата проведения")

  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  for (let i = 1; i <= studentsCount * 4; i++) {
    certification.getRange(1 + i, 4).setValue(currentDate)
  }  

  certification.getRange(1, certification.getLastColumn() + 1).setValue("Тема ВКР")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("Руководитель")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("Консультант")
  certification.getRange(1, certification.getLastColumn() + 1).setValue("Оценка за аттестацию")

  var ruleAttestationPoints = SpreadsheetApp.newDataValidation().requireValueInList(["Зачтено", "Не зачтено", "Не явился"], true).build()
  var rangeAttestation = certification.getRange(2, certification.getLastColumn(), studentsCount * 4, 1)
  rangeAttestation.setDataValidation(ruleAttestationPoints)
  rangeAttestation.setValue(null) 
  
  if (vkrSettings.hasAttendance == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("Присутствие").mergeVertically().setHorizontalAlignment("center")
    certification.getRange(2, certification.getLastColumn(), studentsCount * 4, 1).insertCheckboxes()
  }

  if (vkrSettings.hasCommentsAttestation == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasRemarksAttestation == true) {
    certification.getRange(1, certification.getLastColumn() + 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  importDocToSheet(vkrFirstName, vkrSettings.studentsDocUrl)

  if (vkrSettings.hasGroupSort == true) {
    certification.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (vkrSettings.hasSurnameSort == true) {
    certification.getRange(2, 2, studentsCount, 2).sort(2)
  }

 if (vkrSettings.hasAttendanceNumberSort == true) {
    certification.getRange(2, 3, studentsCount, 3).sort(3)
  }

  if (vkrSettings.hasFilter == true) {
    certification.getRange(1, 1, certification.getLastRow(), certification.getLastColumn()).createFilter()
  }
}

function vkrGraduation(vkrSettings) {
  createSheet(vkrSecondName)
  var graduation = SpreadsheetApp.getActiveSheet()

  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")

  var studentsCount = addStudents(vkrSettings.studentsDocUrl)

  graduation.getRange('A1').setValue("Группа")
  graduation.getRange('B1').setValue("Студент")
  graduation.getRange('C1').setValue("Тема ВКР")
  graduation.getRange('D1').setValue("Руководитель")
  graduation.getRange('E1').setValue("Консультант")
  graduation.getRange('F1').setValue("Рецензент")
  graduation.getRange('G1').setValue("Нормоконтроль")
  graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1).insertCheckboxes()
  graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Дата защиты")

  for (let i = 1; i <= studentsCount; i++) {
    graduation.getRange(1 + i, 8).setValue(currentDate)
  }  

  if (vkrSettings.hasPointsSupervisor == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Оценка руководителя").mergeVertically().setHorizontalAlignment("center")

    var ruleSupervisorPoints = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Не допущен", "Не явился"], true).build()
    var rangePoinsS = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangePoinsS.setDataValidation(ruleSupervisorPoints)
    rangePoinsS.setValue(null) 
  }

  if (vkrSettings.hasPointsReviewer == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Оценка рецензента").mergeVertically().setHorizontalAlignment("center")

    var ruleReviewerPoints = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Не допущен", "Не явился"], true).build()
    var rangePoinsR = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangePoinsR.setDataValidation(ruleReviewerPoints)
    rangePoinsR.setValue(null) 
  }

  if (vkrSettings.hasOriginality == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Оригинальность").mergeVertically().setHorizontalAlignment("center")
    graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1).setValue(0)
  }

  if (vkrSettings.hasResult == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Результат").mergeVertically().setHorizontalAlignment("center")

    var ruleResult = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Не допущен", "Не явился"], true).build()
    var rangeResult = graduation.getRange(2, graduation.getLastColumn(), studentsCount, 1)
    rangeResult.setDataValidation(ruleResult)
    rangeResult.setValue(null) 
  }

  if (vkrSettings.hasRemarks == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Замечания").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasComments == true) {
    graduation.getRange(1, graduation.getLastColumn() + 1).setValue("Комментарии").mergeVertically().setHorizontalAlignment("center")
  }

  if (vkrSettings.hasFilter == true) {
    graduation.getRange(1, 1, graduation.getLastRow(), graduation.getLastColumn()).createFilter()
  }

  importDocToSheet(vkrSecondName, vkrSettings.studentsDocUrl)

  if (vkrSettings.hasGroupSort == true) {
    graduation.getRange(2, 1, studentsCount, 1).sort(1)
  }

  if (vkrSettings.hasSurnameSort == true) {
    graduation.getRange(2, 2, studentsCount, 2).sort(2)
  }
  
  graduation.autoResizeColumns(1, graduation.getLastColumn())
}
