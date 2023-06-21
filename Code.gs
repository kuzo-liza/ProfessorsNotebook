const samplesQuantityKey = "samplesQuantityKey"
const prefixSampleKey = "prefixSampleKey"
const lectionsNaming = "Лекции"
const practicesNaming = "Практические занятия"
const labsNaming = "Лабораторные работы"

function onOpen() { 
  var html = HtmlService.createTemplateFromFile('Page')
      .evaluate()
      .setTitle('Создание и настройка таблиц аттестаций')
  
  SpreadsheetApp.getUi().showSidebar(html)
}

function onInstall() {
  onOpen()
}

function onEdit(e) {
  var range = e.range
  var sheet = range.getSheet()
  var cell = e.range
  var column = cell.getColumn()
  var row = cell.getRow()
    
  if (column == 1 && row > 3 && cell.getValue() !== "") {
    var floatValue = parseFloat(cell.getValue())
    if (isNaN(floatValue) || floatValue < 0) {
      sheet.getRange(row, column).setValue("")
      Browser.msgBox("Допустимы только числа от 0 до 100000")
    }
  }
}

function myIndexOf(s, text){
  return s.indexOf(text) >= 0;
}

function onSelectionChange(e) {
  const range = e.range
  const sheet = range.getSheet()
  const maxRows = sheet.getLastRow()
  const maxColumns = sheet.getLastColumn()

  var countRow = 1
  var activeCellRow
  let propertyService = PropertiesService.getScriptProperties()
  var lectionsSettings
  var practicesSettings

  try {
    let lectionSettingsJSON = propertyService.getProperty(lectionsSettingsKey)
    if (lectionSettingsJSON != null) {
      lectionsSettings = JSON.parse(lectionSettingsJSON)
    }
    let practicesSettingsJSON = propertyService.getProperty(practicesSettingsKey)
    if (practicesSettingsJSON != null) {
      practicesSettings = JSON.parse(practicesSettingsJSON)
    }
    activeCellRow = propertyService.getProperty('active cell row')
  } catch(err) {
      console.log('Не получилось считать данные', err.message)
  }

  var sheetName = sheet.getName()
  if (lectionsSettings != null && myIndexOf(sheetName, lectionsNaming)) {
    if (lectionsSettings.hasComments == true) {
      countRow = countRow + 1
    } 
  
    if (lectionsSettings.hasRemarks == true) {
      countRow = countRow + 1
    }
  }

  if (practicesSettings != null && myIndexOf(sheetName, practicesNaming)) {
    if (practicesSettings.hasComments == true) {
      countRow = countRow + 1
    } 
  
    if (practicesSettings.hasRemarks == true) {
      countRow = countRow + 1
    }
  }
 
  var columnIndex = maxColumns - countRow
  
  if (activeCellRow != null) {
    let activeCellRowAsInt = parseInt(activeCellRow)
    
    if (myIndexOf(sheetName, lectionsNaming) || myIndexOf(sheetName, practicesNaming)) {
      sheet.getRange(activeCellRowAsInt, 1, 1, columnIndex - 1).setBackground(null)
      sheet.getRange(activeCellRowAsInt, columnIndex + 1, 1, sheet.getLastColumn()).setBackground(null)
      if (activeCellRow > 3) {
        coloring(sheet.getRange(activeCellRowAsInt, columnIndex))
      } 
    } else {
      sheet.getRange(activeCellRowAsInt, 1, 1, sheet.getLastColumn()).setBackground(null)
    }
    
    sheet.getRange(activeCellRowAsInt, 1, 1, columnIndex - 1).setBackground(null)
    sheet.getRange(activeCellRowAsInt, columnIndex + 1, 1, sheet.getLastColumn()).setBackground(null)
    if (activeCellRow > 3 && (myIndexOf(sheetName, lectionsNaming) || myIndexOf(sheetName, practicesNaming))) {
      coloring(sheet.getRange(activeCellRowAsInt, columnIndex))
    } 
  }

  if ((myIndexOf(sheetName, lectionsNaming) || myIndexOf(sheetName, practicesNaming) || myIndexOf(sheetName, labsNaming)) &&range.getRow() > 3 && range.getRow() < maxRows + 1) {
    sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).setBackground('orange')
  } else if (range.getRow() > 1 && range.getRow() < maxRows + 1) {
    sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).setBackground('orange')
  }

  var cellRow = range.getRow()
  var cellColumn = range.getColumn()

  var checkRange = sheet.getRange(cellRow, columnIndex)
  if ((myIndexOf(sheetName, lectionsNaming) || myIndexOf(sheetName, practicesNaming)) && cellColumn > 2 && cellColumn < columnIndex) {
    coloring(checkRange)
  }
  propertyService.setProperty('active cell row', cellRow)
}

function coloring(checkRange) {
  if (checkRange.getValue() >= "90") {
      checkRange.setBackground('#006600')
    } 
    
    if (checkRange.getValue() < "90" && checkRange.getValue()>= "75") {
      checkRange.setBackground('#CCFF99')
    }

    if (checkRange.getValue() < "75" && checkRange.getValue() >= "60") {
      checkRange.setBackground('#FFFF99')
    }

    if (checkRange.getValue() < "60") {
      checkRange.setBackground('#FF6666')
    }
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent()
}

function importDocToSheet(name, url) {
  var document
  try {
    document = DocumentApp.openByUrl(url) 
  } catch(err) {
    console.log('Не получается открыть документ со студентами', err)
    return
  }
  var body = document.getBody()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getActiveSheet()
  var paragraphs = body.getParagraphs()

  if (name == lectionsNaming || name == practicesNaming || name == labsNaming) {
    for (var i = 0; i < paragraphs.length; i++) {
      var text = paragraphs[i].getText()
      sheet.getRange(i + 4, 2).setValue(text)
    }
  } else if (name == 'Аттестации ВКР') {
    for (var j = 0; j < 4; j++) {
      for (var i = 0; i < paragraphs.length; i++) {
        var text = paragraphs[i].getText()
        sheet.getRange(j * paragraphs.length + i + 2, 2).setValue(text)
      }
    }
  } else {
    for (var i = 0; i < paragraphs.length; i++) {
      var text = paragraphs[i].getText()
      sheet.getRange(i + 2, 2).setValue(text)
    }    
  }
}

function addStudents(url) {
  var document
  try {
    document = DocumentApp.openByUrl(url) 
  } catch(err) {
    console.log('Не получается открыть документ со студентами', err)
    return 10
  }
  var studentsCount = document.getBody().getParagraphs().length
  return studentsCount
}

function mySample() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  activeSpreadsheet.insertSheet('Шаблон', activeSpreadsheet.getNumSheets() + 1)
}

function getCustomSamples() {
  try {
    let propertyService = PropertiesService.getUserProperties()
    let samplesQuantityPropertyValue = propertyService.getProperty(samplesQuantityKey)
    if (samplesQuantityPropertyValue == null) {
      console.log('Не получилось достать шаблон. Шаблонов нет.')
      return 0
    } else {
      return parseInt(samplesQuantityPropertyValue)
    }
  } catch (err) {
    console.log('Не получилось достать данные из настроек', err.message)
  }
  return 0
}

function myDataForSampleAsJSON() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSheet()
  if (activeSpreadsheet.getSheetName().toString() == "Шаблон") {
    var columnNumber = activeSpreadsheet.getLastColumn()
    var rowNumber = activeSpreadsheet.getLastRow()

    var cellObjects = []

    let customSampleObject = new Object()
    customSampleObject.rows = rowNumber
    customSampleObject.columns = columnNumber

    for (let column = 1; column <= columnNumber; column++) {
      for (let row = 1; row <= rowNumber; row++) {
          var cell = activeSpreadsheet.getRange(row, column)

          let cellObject = new Object()
          cellObject.row = row
          cellObject.column = column
          cellObject.value = cell.getValue()
          cellObjects.push(cellObject)
      }
    }
    customSampleObject.cells = cellObjects

    let propertyService = PropertiesService.getUserProperties()
    try {
      var samplesQuantityPropertyValue = propertyService.getProperty(samplesQuantityKey)
      let sampleId
      if (samplesQuantityPropertyValue == null) {
        sampleId = 1
      } else {
        sampleId = parseInt(samplesQuantityPropertyValue) + 1
      }
      propertyService.setProperty(samplesQuantityKey, sampleId)
      propertyService.setProperty(prefixSampleKey + sampleId, JSON.stringify(customSampleObject))
    } catch(err) {
      console.log('Не получилось записать данные собственного шаблона', err.message)
    }
    return false
  } else {
    return true
  }
}

function createMySampleFromJSON(id) {
  SpreadsheetApp.getActiveSpreadsheet().insertSheet('Мой шаблон ' + id)
  var mySample = SpreadsheetApp.getActiveSheet()

  var sampleObject
  try {
    let propertyService = PropertiesService.getUserProperties()
    sampleObject = JSON.parse(propertyService.getProperty(prefixSampleKey + id))
    if (sampleObject == null) {
      console.log('Не получилось достать шаблон. Шаблон не найден.')
      return
    }
  } catch (err) {
    console.log('Не получилось достать данные из настроек', err.message)
  }

  for (let i = 0; i < sampleObject.cells.length; i++) {
    let cell = sampleObject.cells[i]
    mySample.getRange(cell.row, cell.column).setValue(cell.value)
  }
}

function createSheet(sheetName) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let id = activeSpreadsheet.getNumSheets()
  activeSpreadsheet.insertSheet(sheetName + " " + (id + 1), id + 1)
}

function setGrade(studentsCount) {
  let activeSheet = SpreadsheetApp.getActiveSheet()
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Отлично", "Хорошо", "Удовлетворительно", "Неудовлетворительно", "Зачет", "Не зачет", "Не явился", "Допущен", "Не допущен"], true).build()
  var range = activeSheet.getRange(2, activeSheet.getLastColumn(), studentsCount, 1)
  range.setDataValidation(rule)
  range.setValue(null)
}

function setDate(studentsCount) {
  let activeSheet = SpreadsheetApp.getActiveSheet()
  var date = new Date()
  var currentDate = Utilities.formatDate(date, "GMT", "dd.MM")
  activeSheet.getRange(2, activeSheet.getLastColumn(), studentsCount,1).setValue(currentDate).setHorizontalAlignment("center")
}

