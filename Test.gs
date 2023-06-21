function runTests() {
  if ((typeof GasTap)==='undefined') { // GasT Initialization. (only if not initialized yet.)
    eval(UrlFetchApp.fetch('https://raw.githubusercontent.com/huan/gast/master/src/gas-tap-lib.js').getContentText())
  }

  var test = new GasTap()
  let testStudentsDocUrl = "https://docs.google.com/document/d/10H0jGQS0pqyu8Ce1PYlHUUHfTn_VRSC0Rqn8S2nvmn0/edit"

  test('TestColoring', function(t) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var tmpSheet = activeSpreadsheet.insertSheet()

    var a1Range = tmpSheet.getRange('A1')
    a1Range.setValue(90)
    coloring(a1Range)
    t.equal('#006600', a1Range.getBackground(), 'Check dark green color')

    a1Range.setValue(75)
    coloring(a1Range)
    t.equal('#ccff99', a1Range.getBackground(), 'Check light green color')

    a1Range.setValue(60)
    coloring(a1Range)
    t.equal('#ffff99', a1Range.getBackground(), 'Check yellow color')

    a1Range.setValue(59)
    coloring(a1Range)
    t.equal('#ff6666', a1Range.getBackground(), 'Check red color')

    deleteSheetIfExist(tmpSheet)
  })

  test('TestAddStudents', function(t) {
    let actualStudentsCount = addStudents(testStudentsDocUrl)
    t.equal(10, actualStudentsCount, 'Check students count')
  })

  test('TestCreateSheet', function(t) {
    let sheetName = "mySample"
    createSheet(sheetName)
    
    let sheetsNumber = SpreadsheetApp.getActiveSpreadsheet().getNumSheets()
    let actualSheet = getSheet(sheetName + ' ' + sheetsNumber)
    
    t.ok(actualSheet != null, 'Sheet was created')

    deleteSheetIfExistByName(actualSheet)
  })

  test('TestMySample', function(t) {
    var actualSheet = getSheet('Шаблон')
    if (actualSheet != null) {
      t.fail('Sample already exist')
    } else {
      mySample()
      actualSheet = getSheet('Шаблон')
      t.ok(actualSheet != null, 'Sheet was created')
    }
    deleteSheetIfExist(actualSheet)
  })

  test('TestImportDocToSheet', function(t){
    let activeSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet()
    var document
    try {
      document = DocumentApp.openByUrl(testStudentsDocUrl) 
    } catch(err) {
      t.fail('Could not read students')
      return
    }

    var paragraphs = document.getBody().getParagraphs()
    importDocToSheet('Лекции', testStudentsDocUrl)
    for (let i = 0; i < paragraphs.length; i++) {
      if (activeSheet.getRange(i + 4, 2).getValue() != paragraphs[i].getText()) {
        t.fail('Students in lections are not the same')
      }
    }

    t.ok(true, 'Students in lections are same')
    activeSheet.clear()

    importDocToSheet('Аттестации ВКР', testStudentsDocUrl)
    for (let j = 0; j < 4; j++) {
      for (let i = 0; i < paragraphs.length; i++) {
        if (activeSheet.getRange(j * paragraphs.length + i + 2, 2).getValue() != paragraphs[i].getText()) {
          t.fail('Students in attestation are not the same')
        }
      }
    }
    
    t.ok(true, 'Students in attestation are same')
    activeSheet.clear()

    importDocToSheet('Эссе', testStudentsDocUrl)
    for (let i = 0; i < paragraphs.length; i++) {
      if (activeSheet.getRange(i + 2, 2).getValue() != paragraphs[i].getText()) {
        t.fail('Students in essay are not the same')
      }
    }

    t.ok(true, 'Students in essay are same')
    deleteSheetIfExist(activeSheet)
  })

  test('TestLectionWriteParams', function(t) {
    let lectionsSettings = new Object()
    lectionsSettings.number = 4
    lectionsSettings.hasName = true
    lectionsSettings.hasFormat = true
    lectionsSettings.hasDuration = true
    lectionsSettings.hasStudentDuration = true
    lectionsSettings.hasComments = false
    lectionsSettings.hasRemarks = true
    lectionsSettings.hasFilter = true
    lectionsSettings.hasGroupSort = false
    lectionsSettings.hasSurnameSort = true
    lectionsSettings.studentsDocUrl = testStudentsDocUrl

    saveLectionsSettings(JSON.stringify(lectionsSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualLectionsSettingsJSON = propertyService.getProperty(lectionsSettingsKey)
    let actualLectionsSettings
    if (actualLectionsSettingsJSON != null) {
      actualLectionsSettings = JSON.parse(actualLectionsSettingsJSON)
    } else {
      t.fail('Lections settings not found')
    }

    let actualLectionsNum = actualLectionsSettings.number
    let actualStudensDocUrl = actualLectionsSettings.studentsDocUrl

    t.equal(lectionsSettings.number, actualLectionsNum, 'Check lections number')
    t.equal(testStudentsDocUrl, actualStudensDocUrl, 'Check student document url')
    t.deepEqual(lectionsSettings, actualLectionsSettings, 'Check settings')
  })

  test('TestLectionSampleFilling', function(t) {
    let lectionsSettings = new Object()
    lectionsSettings.number = 4
    lectionsSettings.hasName = false
    lectionsSettings.hasFormat = false
    lectionsSettings.hasDuration = false
    lectionsSettings.hasStudentDuration = true
    lectionsSettings.hasComments = false
    lectionsSettings.hasRemarks = false
    lectionsSettings.hasFilter = false
    lectionsSettings.hasGroupSort = false
    lectionsSettings.hasSurnameSort = false
    lectionsSettings.studentsDocUrl = testStudentsDocUrl

    saveLectionsSettings(JSON.stringify(lectionsSettings))
    lections()

    let lection = SpreadsheetApp.getActiveSheet();
    let lectionsStudentDurationRange = lection.getRange(4, 3, 11, lectionsSettings.number)
    t.equal("Длительность посещения студентом: ", lectionsStudentDurationRange.getNote(), 'Check lections student duration')

    deleteSheetIfExist(lection)
  })

  test('TestPracticeWriteParams', function(t) {
    let practicesSettings = new Object()
    practicesSettings.number = 4
    practicesSettings.hasName = true
    practicesSettings.hasFormat = true
    practicesSettings.hasDuration = true
    practicesSettings.hasStudentDuration = true
    practicesSettings.hasComments = false
    practicesSettings.hasRemarks = true
    practicesSettings.hasFilter = true
    practicesSettings.hasGroupSort = false
    practicesSettings.hasSurnameSort = true
    practicesSettings.studentsDocUrl = testStudentsDocUrl

    savePracticesSettings(JSON.stringify(practicesSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualPracticesSettingsJSON = propertyService.getProperty(practicesSettingsKey)
    let actualPracticesSettings
    if (actualPracticesSettingsJSON != null) {
      actualPracticesSettings = JSON.parse(actualPracticesSettingsJSON)
    } else {
      t.fail('Practices settings not found')
    }

    let actualPracticesNum = actualPracticesSettings.number
    let actualStudensDocUrl = actualPracticesSettings.studentsDocUrl

    t.equal(practicesSettings.number, actualPracticesNum, 'Check practices number')
    t.equal(testStudentsDocUrl, actualStudensDocUrl, 'Check student document url')
    t.deepEqual(practicesSettings, actualPracticesSettings, 'Check settings')
  })

  test('TestPracticesSampleFilling', function(t) {
    let practicesSettings = new Object()
    practicesSettings.number = 4
    practicesSettings.hasName = false
    practicesSettings.hasFormat = true
    practicesSettings.hasDuration = false
    practicesSettings.hasStudentDuration = false
    practicesSettings.hasComments = false
    practicesSettings.hasRemarks = false
    practicesSettings.hasFilter = false
    practicesSettings.hasGroupSort = false
    practicesSettings.hasSurnameSort = false
    practicesSettings.studentsDocUrl = testStudentsDocUrl

    savePracticesSettings(JSON.stringify(practicesSettings))
    practices()

    let practicesSheet = SpreadsheetApp.getActiveSheet();
    let practicesFormatRange = practicesSheet.getRange(2, 3, 11, practicesSettings.number)
    t.equal("Формат занятия: " + '\n', practicesFormatRange.getNote(), 'Check practices format')

    deleteSheetIfExist(practicesSheet)
  })

  test('TestLabsWriteParams', function(t) {
    let labsSettings = new Object()
    labsSettings.number = 4
    labsSettings.hasName = true
    labsSettings.hasReport = true
    labsSettings.hasComments = false
    labsSettings.hasRemarks = true
    labsSettings.hasFilter = true
    labsSettings.hasGroupSort = false
    labsSettings.hasSurnameSort = true
    labsSettings.studentsDocUrl = testStudentsDocUrl

    saveLabsSettings(JSON.stringify(labsSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualLabsSettingsJSON = propertyService.getProperty(labsSettingsKey)
    let actualLabsSettings
    if (actualLabsSettingsJSON != null) {
      actualLabsSettings = JSON.parse(actualLabsSettingsJSON)
    } else {
      t.fail('Labs settings not found')
    }

    let actualLabsNum = actualLabsSettings.number
    let actualStudensDocUrl = actualLabsSettings.studentsDocUrl

    t.equal(labsSettings.number, actualLabsNum, 'Check labs number')
    t.equal(testStudentsDocUrl, actualStudensDocUrl, 'Check student document url')
    t.deepEqual(labsSettings, actualLabsSettings, 'Check settings')
  })

  test('TestLabsSampleFilling', function(t) {
    let labsSettings = new Object()
    labsSettings.number = 4
    labsSettings.hasName = true
    labsSettings.hasReport = true
    labsSettings.hasComments = false
    labsSettings.hasRemarks = false
    labsSettings.hasFilter = false
    labsSettings.hasGroupSort = false
    labsSettings.hasSurnameSort = false
    labsSettings.studentsDocUrl = testStudentsDocUrl

    saveLabsSettings(JSON.stringify(labsSettings))
    labs()

    let labSheet = SpreadsheetApp.getActiveSheet();
    let labsNameRange = labSheet.getRange(2, 3, 11, 4)
    t.equal("Название лабораторной работы " + 1 + ":" + '\n', labsNameRange.getNote(), 'Check labs name')

    deleteSheetIfExist(labSheet)
  })

  test('TestCourseWorkWriteParams', function(t) {
    let courseWorkSettings = new Object()
    courseWorkSettings.hasVariant = true
    courseWorkSettings.hasReport = true
    courseWorkSettings.hasComments = true
    courseWorkSettings.hasRemarks = true
    courseWorkSettings.hasFilter = true
    courseWorkSettings.hasGroupSort = false
    courseWorkSettings.hasSurnameSort = true
    courseWorkSettings.studentsDocUrl = testStudentsDocUrl

    saveCourseWorkSettings(JSON.stringify(courseWorkSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualCourseWorkSettingsJSON = propertyService.getProperty(courseWorkSettingsKey)
    let actualCourseWorkSettings
    if (actualCourseWorkSettingsJSON != null) {
      actualCourseWorkSettings = JSON.parse(actualCourseWorkSettingsJSON)
    } else {
      t.fail('Course work settings not found')
    }

    t.deepEqual(courseWorkSettings, actualCourseWorkSettings, 'Check settings')
  })

 test('TestCourseWorkFilling', function(t) {
   let courseWorkSettings = new Object()
   courseWorkSettings.hasVariant = true
   courseWorkSettings.hasReport = true
   courseWorkSettings.hasComments = true
   courseWorkSettings.hasRemarks = false
   courseWorkSettings.hasFilter = false
   courseWorkSettings.hasGroupSort = false
   courseWorkSettings.hasSurnameSort = false
   courseWorkSettings.studentsDocUrl = testStudentsDocUrl

   saveCourseWorkSettings(JSON.stringify(courseWorkSettings))
   courseWork()

   let courseWorkSheet = SpreadsheetApp.getActiveSheet();

   t.equal("Группа", courseWorkSheet.getRange('A1').getValue(), 'Check add group column')
   t.equal("Студент", courseWorkSheet.getRange('B1').getValue(), 'Check add student column')
   t.equal("Название", courseWorkSheet.getRange('C1').getValue(), 'Check add name column')
   t.equal("Вариант", courseWorkSheet.getRange('D1').getValue(), 'Check add variant column')
   t.equal("Дата выдачи", courseWorkSheet.getRange('E1').getValue(), 'Check add date of isuue column')
   t.equal("Дата сдачи", courseWorkSheet.getRange('F1').getValue(), 'Check add delivery date column')
   t.equal("Готовность", courseWorkSheet.getRange('G1').getValue(), 'Check add ready column')
   t.equal("Отчет", courseWorkSheet.getRange('H1').getValue(), 'Check add report column')
   t.equal("Оценка", courseWorkSheet.getRange('I1').getValue(), 'Check add points column')
   t.equal("Комментарии", courseWorkSheet.getRange('J1').getValue(), 'Check add comments column')

   deleteSheetIfExist(courseWorkSheet)
 })

  test('TestCourseProjectWriteParams', function(t) {
    let courseProjectSettings = new Object()
    courseProjectSettings.hasVariant = true
    courseProjectSettings.hasReport = true
    courseProjectSettings.hasComments = true
    courseProjectSettings.hasRemarks = true
    courseProjectSettings.hasFilter = true
    courseProjectSettings.hasGroupSort = false
    courseProjectSettings.hasSurnameSort = true
    courseProjectSettings.studentsDocUrl = testStudentsDocUrl

    saveCourseProjectSettings(JSON.stringify(courseProjectSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualCourseProjectSettingsJSON = propertyService.getProperty(courseProjectSettingsKey)
    let actualCourseProjectSettings
    if (actualCourseProjectSettingsJSON != null) {
      actualCourseProjectSettings = JSON.parse(actualCourseProjectSettingsJSON)
    } else {
      t.fail('Course project settings not found')
    }

    t.deepEqual(courseProjectSettings, actualCourseProjectSettings, 'Check settings')
  })

 test('TestCourseProjectFilling', function(t) {
   let courseProjectSettings = new Object()
   courseProjectSettings.hasVariant = true
   courseProjectSettings.hasReport = true
   courseProjectSettings.hasComments = false
   courseProjectSettings.hasRemarks = true
   courseProjectSettings.hasFilter = false
   courseProjectSettings.hasGroupSort = false
   courseProjectSettings.hasSurnameSort = false
   courseProjectSettings.studentsDocUrl = testStudentsDocUrl

   saveCourseProjectSettings(JSON.stringify(courseProjectSettings))
   courseProject()

   let courseProjectSheet = SpreadsheetApp.getActiveSheet();

   t.equal("Группа", courseProjectSheet.getRange('A1').getValue(), 'Check add group column')
   t.equal("Студент", courseProjectSheet.getRange('B1').getValue(), 'Check add student column')
   t.equal("Название", courseProjectSheet.getRange('C1').getValue(), 'Check add name column')
   t.equal("Вариант", courseProjectSheet.getRange('D1').getValue(), 'Check add variant column')
   t.equal("Дата выдачи", courseProjectSheet.getRange('E1').getValue(), 'Check add date of isuue column')
   t.equal("Дата сдачи", courseProjectSheet.getRange('F1').getValue(), 'Check add delivery date column')
   t.equal("Готовность", courseProjectSheet.getRange('G1').getValue(), 'Check add ready column')
   t.equal("Отчет", courseProjectSheet.getRange('H1').getValue(), 'Check add report column')
   t.equal("Оценка", courseProjectSheet.getRange('I1').getValue(), 'Check add points column')
   t.equal("Замечания", courseProjectSheet.getRange('J1').getValue(), 'Check add remarks column')

   deleteSheetIfExist(courseProjectSheet)
  })

  test('TestTestsWriteParams', function(t) {
    let testsSettings = new Object()
    testsSettings.hasName = true
    testsSettings.hasDuration = true
    testsSettings.hasStudentDuration = true
    testsSettings.hasPresence = true
    testsSettings.hasComments = true
    testsSettings.hasRemarks = true
    testsSettings.hasFilter = true
    testsSettings.hasGroupSort = false
    testsSettings.hasSurnameSort = true
    testsSettings.studentsDocUrl = testStudentsDocUrl

    saveTestsSettings(JSON.stringify(testsSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualTestsSettingsJSON = propertyService.getProperty(testsSettingsKey)
    let actualTestsSettings
    if (actualTestsSettingsJSON != null) {
      actualTestsSettings = JSON.parse(actualTestsSettingsJSON)
    } else {
      t.fail('Test settings not found')
    }

    t.deepEqual(testsSettings, actualTestsSettings, 'Check settings')
  })

   test('TestTestsFilling', function(t) {
    let testsSettings = new Object()
    testsSettings.hasName = true
    testsSettings.hasDuration = true
    testsSettings.hasStudentDuration = true
    testsSettings.hasPresence = true
    testsSettings.hasComments = false
    testsSettings.hasRemarks = true
    testsSettings.hasFilter = false
    testsSettings.hasGroupSort = false
    testsSettings.hasSurnameSort = false
    testsSettings.studentsDocUrl = testStudentsDocUrl

    saveTestsSettings(JSON.stringify(testsSettings))
    tests()

    let testsSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", testsSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", testsSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Номер", testsSheet.getRange('C1').getValue(), 'Check add number column')
    t.equal("Вариант", testsSheet.getRange('D1').getValue(), 'Check add variant column')
    t.equal("Дата", testsSheet.getRange('E1').getValue(), 'Check add date column')
    t.equal("Название", testsSheet.getRange('F1').getValue(), 'Check add name column')
    t.equal("Длительность теста", testsSheet.getRange('G1').getValue(), 'Check add duration column')
    t.equal("Длительность решения", testsSheet.getRange('H1').getValue(), 'Check add students duration column')
    t.equal("Присутствие", testsSheet.getRange('I1').getValue(), 'Check add presence column')
    t.equal("Оценка", testsSheet.getRange('J1').getValue(), 'Check add points column')
    t.equal("Замечания", testsSheet.getRange('K1').getValue(), 'Check add remarks column')

    deleteSheetIfExist(testsSheet)
  })

  test('TestControlsWriteParams', function(t) {
    let controlsSettings = new Object()
    controlsSettings.hasName = true
    controlsSettings.hasDuration = true
    controlsSettings.hasStudentDuration = true
    controlsSettings.hasPresence = true
    controlsSettings.hasComments = true
    controlsSettings.hasRemarks = true
    controlsSettings.hasFilter = true
    controlsSettings.hasGroupSort = false
    controlsSettings.hasSurnameSort = true
    controlsSettings.studentsDocUrl = testStudentsDocUrl

    saveControlsSettings(JSON.stringify(controlsSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualControlsSettingsJSON = propertyService.getProperty(controlsSettingsKey)
    let actualControlsSettings
    if (actualControlsSettingsJSON != null) {
      actualControlsSettings = JSON.parse(actualControlsSettingsJSON)
    } else {
      t.fail('Controls settings not found')
    }

    t.deepEqual(controlsSettings, actualControlsSettings, 'Check settings')
  })

  test('TestControlsFilling', function(t) {
    let controlsSettings = new Object()
    controlsSettings.hasName = true
    controlsSettings.hasDuration = true
    controlsSettings.hasStudentDuration = true
    controlsSettings.hasPresence = true
    controlsSettings.hasComments = true
    controlsSettings.hasRemarks = false
    controlsSettings.hasFilter = false
    controlsSettings.hasGroupSort = false
    controlsSettings.hasSurnameSort = false
    controlsSettings.studentsDocUrl = testStudentsDocUrl

    saveControlsSettings(JSON.stringify(controlsSettings))
    controls()

    let controlsSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", controlsSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", controlsSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Номер", controlsSheet.getRange('C1').getValue(), 'Check add number column')
    t.equal("Вариант", controlsSheet.getRange('D1').getValue(), 'Check add variant column')
    t.equal("Дата", controlsSheet.getRange('E1').getValue(), 'Check add date column')
    t.equal("Название", controlsSheet.getRange('F1').getValue(), 'Check add name column')
    t.equal("Длительность работы", controlsSheet.getRange('G1').getValue(), 'Check add duration column')
    t.equal("Длительность решения", controlsSheet.getRange('H1').getValue(), 'Check add students duration column')
    t.equal("Присутствие", controlsSheet.getRange('I1').getValue(), 'Check add presence column')
    t.equal("Оценка", controlsSheet.getRange('J1').getValue(), 'Check add points column')
    t.equal("Комментарии", controlsSheet.getRange('K1').getValue(), 'Check add comments column')

    deleteSheetIfExist(controlsSheet)
  })

  test('TestReportsWriteParams', function(t) {
    let reportsSettings = new Object()
    reportsSettings.hasNumber = true
    reportsSettings.hasDuration = false
    reportsSettings.hasPresentation = false
    reportsSettings.hasComments = false
    reportsSettings.hasRemarks = false
    reportsSettings.hasFilter = false
    reportsSettings.hasGroupSort = false
    reportsSettings.hasSurnameSort = false
    reportsSettings.studentsDocUrl = testStudentsDocUrl

    saveReportsSettings(JSON.stringify(reportsSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualReportsSettingsJSON = propertyService.getProperty(reportsSettingsKey)
    let actualReportsSettings
    if (actualReportsSettingsJSON != null) {
      actualReportsSettings = JSON.parse(actualReportsSettingsJSON)
    } else {
      t.fail('Reports settings not found')
    }

    t.deepEqual(reportsSettings, actualReportsSettings, 'Check settings')
  })

  test('TestReportsFilling', function(t) {
    let reportsSettings = new Object()
    reportsSettings.hasNumber = true
    reportsSettings.hasDuration = false
    reportsSettings.hasPresentation = false
    reportsSettings.hasComments = false
    reportsSettings.hasRemarks = false
    reportsSettings.hasFilter = false
    reportsSettings.hasGroupSort = false
    reportsSettings.hasSurnameSort = false
    reportsSettings.studentsDocUrl = testStudentsDocUrl

    saveReportsSettings(JSON.stringify(reportsSettings))
    reports()

    let reportsSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", reportsSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", reportsSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Название", reportsSheet.getRange('C1').getValue(), 'Check add report name column')
    t.equal("Дата выступления", reportsSheet.getRange('D1').getValue(), 'Check add date column')
    t.equal("Номер выступления", reportsSheet.getRange('E1').getValue(), 'Check add report number column')

    deleteSheetIfExist(reportsSheet)
  })

  test('TestAbstractWriteParams', function(t) {
    let abstractSettings = new Object()
    abstractSettings.hasVariant = true
    abstractSettings.hasDate = true
    abstractSettings.hasComments = true
    abstractSettings.hasRemarks = false
    abstractSettings.hasFilter = false
    abstractSettings.hasGroupSort = false
    abstractSettings.hasSurnameSort = false
    abstractSettings.studentsDocUrl = testStudentsDocUrl

    saveAbstractSettings(JSON.stringify(abstractSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualAbstractSettingsJSON = propertyService.getProperty(abstractSettingsKey)
    let actualAbstractSettings
    if (actualAbstractSettingsJSON != null) {
      actualAbstractSettings = JSON.parse(actualAbstractSettingsJSON)
    } else {
      t.fail('Abstract settings not found')
    }

    t.deepEqual(abstractSettings, actualAbstractSettings, 'Check settings')
  })

  test('TestAbstractFilling', function(t) {
    let abstractSettings = new Object()
    abstractSettings.hasVariant = true
    abstractSettings.hasDate = true
    abstractSettings.hasComments = true
    abstractSettings.hasRemarks = false
    abstractSettings.hasFilter = false
    abstractSettings.hasGroupSort = false
    abstractSettings.hasSurnameSort = false
    abstractSettings.studentsDocUrl = testStudentsDocUrl

    saveAbstractSettings(JSON.stringify(abstractSettings))
    abstract()

    let abstractSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", abstractSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", abstractSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Название", abstractSheet.getRange('C1').getValue(), 'Check add abstract name column')
    t.equal("Вариант", abstractSheet.getRange('D1').getValue(), 'Check add variant column')
    t.equal("Дата сдачи", abstractSheet.getRange('E1').getValue(), 'Check add date of issue column')
    t.equal("Дата выдачи", abstractSheet.getRange('F1').getValue(), 'Check add delivery date column')
    t.equal("Оценка", abstractSheet.getRange('G1').getValue(), 'Check add points column')
    t.equal("Комментарии", abstractSheet.getRange('H1').getValue(), 'Check add comments column')

    deleteSheetIfExist(abstractSheet)
  })

  test('TestEssayWriteParams', function(t) {
    let essaySettings = new Object()
    essaySettings.hasVariant = true
    essaySettings.hasDate = true
    essaySettings.hasComments = false
    essaySettings.hasRemarks = true
    essaySettings.hasFilter = false
    essaySettings.hasGroupSort = false
    essaySettings.hasSurnameSort = false
    essaySettings.studentsDocUrl = testStudentsDocUrl

    saveEssaySettings(JSON.stringify(essaySettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualEssaySettingsJSON = propertyService.getProperty(essaySettingsKey)
    let actualEssaySettings
    if (actualEssaySettingsJSON != null) {
      actualEssaySettings = JSON.parse(actualEssaySettingsJSON)
    } else {
      t.fail('Essay settings not found')
    }

    t.deepEqual(essaySettings, actualEssaySettings, 'Check settings')
  })

   test('TestEssayFilling', function(t) {
    let essaySettings = new Object()
    essaySettings.hasVariant = true
    essaySettings.hasDate = true
    essaySettings.hasComments = false
    essaySettings.hasRemarks = true
    essaySettings.hasFilter = false
    essaySettings.hasGroupSort = false
    essaySettings.hasSurnameSort = false
    essaySettings.studentsDocUrl = testStudentsDocUrl

    saveEssaySettings(JSON.stringify(essaySettings))
    essay()

    let essaySheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", essaySheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", essaySheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Название", essaySheet.getRange('C1').getValue(), 'Check add name column')
    t.equal("Вариант", essaySheet.getRange('D1').getValue(), 'Check add variant column')
    t.equal("Дата сдачи", essaySheet.getRange('E1').getValue(), 'Check add date of issue column')
    t.equal("Дата выдачи", essaySheet.getRange('F1').getValue(), 'Check add delivery date column')
    t.equal("Оценка", essaySheet.getRange('G1').getValue(), 'Check add points column')
    t.equal("Замечания", essaySheet.getRange('H1').getValue(), 'Check add remarks column')

    deleteSheetIfExist(essaySheet)
  })

    test('TestHomeworkWriteParams', function(t) {
    let homeworkSettings = new Object()
    homeworkSettings.hasName = true
    homeworkSettings.hasBlackboardAnswer = true
    homeworkSettings.hasComments = true
    homeworkSettings.hasRemarks = true
    homeworkSettings.hasFilter = true
    homeworkSettings.hasGroupSort = false
    homeworkSettings.hasSurnameSort = true
    homeworkSettings.studentsDocUrl = testStudentsDocUrl

    saveHomeworksSettings(JSON.stringify(homeworkSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualHomeworkSettingsJSON = propertyService.getProperty(homeworksSettingsKey)
    let actualHomeworkSettings
    if (actualHomeworkSettingsJSON != null) {
      actualHomeworkSettings = JSON.parse(actualHomeworkSettingsJSON)
    } else {
      t.fail('Homework settings not found')
    }

    t.deepEqual(homeworkSettings, actualHomeworkSettings, 'Check settings')
  })

   test('TestHomeworkFilling', function(t) {
    let homeworkSettings = new Object()
    homeworkSettings.hasName = true
    homeworkSettings.hasBlackboardAnswer = true
    homeworkSettings.hasComments = false
    homeworkSettings.hasRemarks = true
    homeworkSettings.hasFilter = false
    homeworkSettings.hasGroupSort = false
    homeworkSettings.hasSurnameSort = false
    homeworkSettings.studentsDocUrl = testStudentsDocUrl

    saveHomeworksSettings(JSON.stringify(homeworkSettings))
    homeworks()

    let homeworkSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", homeworkSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", homeworkSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Номер", homeworkSheet.getRange('C1').getValue(), 'Check add number column')
    t.equal("Вариант", homeworkSheet.getRange('D1').getValue(), 'Check add variant column')
    t.equal("Дата сдачи", homeworkSheet.getRange('E1').getValue(), 'Check add delivery date column')
    t.equal("Название", homeworkSheet.getRange('F1').getValue(), 'Check add name column')
    t.equal("Выход к доске", homeworkSheet.getRange('G1').getValue(), 'Check add blackboard answer column')
    t.equal("Оценка", homeworkSheet.getRange('H1').getValue(), 'Check add points column')
    t.equal("Замечания", homeworkSheet.getRange('I1').getValue(), 'Check add remarks column')

    deleteSheetIfExist(homeworkSheet)
  })

    test('TestEdPracticeWriteParams', function(t) {
    let edPracticeSettings = new Object()
    edPracticeSettings.hasCurator = true
    edPracticeSettings.hasTheme = true
    edPracticeSettings.hasComments = true
    edPracticeSettings.hasRemarks = false
    edPracticeSettings.hasFilter = false
    edPracticeSettings.hasGroupSort = false
    edPracticeSettings.hasSurnameSort = false
    edPracticeSettings.studentsDocUrl = testStudentsDocUrl

    saveEdPracticeSettings(JSON.stringify(edPracticeSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualEdPracticeSettingsJSON = propertyService.getProperty(edPracticeSettingsKey)
    let actualEdPracticeSettings
    if (actualEdPracticeSettingsJSON != null) {
      actualEdPracticeSettings = JSON.parse(actualEdPracticeSettingsJSON)
    } else {
      t.fail('Education practice settings not found')
    }

    t.deepEqual(edPracticeSettings, actualEdPracticeSettings, 'Check settings')
  })

 test('TestEdPracticeFilling', function(t) {
    let edPracticeSettings = new Object()
    edPracticeSettings.hasCurator = true
    edPracticeSettings.hasTheme = true
    edPracticeSettings.hasComments = true
    edPracticeSettings.hasRemarks = false
    edPracticeSettings.hasFilter = false
    edPracticeSettings.hasGroupSort = false
    edPracticeSettings.hasSurnameSort = false
    edPracticeSettings.studentsDocUrl = testStudentsDocUrl

    saveEdPracticeSettings(JSON.stringify(edPracticeSettings))
    edPractice()

    let edPracticeSheet = SpreadsheetApp.getActiveSheet();

    t.equal("Группа", edPracticeSheet.getRange('A1').getValue(), 'Check add group column')
    t.equal("Студент", edPracticeSheet.getRange('B1').getValue(), 'Check add student column')
    t.equal("Вид практики", edPracticeSheet.getRange('C1').getValue(), 'Check add type column')
    t.equal("Дата начала", edPracticeSheet.getRange('D1').getValue(), 'Check add start date column')
    t.equal("Дата окончания", edPracticeSheet.getRange('E1').getValue(), 'Check add end date column')
    t.equal("Отчет", edPracticeSheet.getRange('F1').getValue(), 'Check add report column')
    t.equal("Куратор", edPracticeSheet.getRange('G1').getValue(), 'Check add curator column')
    t.equal("Тема", edPracticeSheet.getRange('H1').getValue(), 'Check add theme column')
    t.equal("Оценка", edPracticeSheet.getRange('I1').getValue(), 'Check add points column')
    t.equal("Комментарии", edPracticeSheet.getRange('J1').getValue(), 'Check add comments column')

    deleteSheetIfExist(edPracticeSheet)
  })

    test('TestVKRWriteParams', function(t) {
    let vkrSettings = new Object()
    vkrSettings.hasAttendance = false
    vkrSettings.hasCommentsAttestation = false
    vkrSettings.hasRemarksAttestation = false
    vkrSettings.hasPointsSupervisor = false
    vkrSettings.hasPointsReviewer = false
    vkrSettings.hasOriginality = false
    vkrSettings.hasResult = false
    vkrSettings.hasComments = false
    vkrSettings.hasRemarks = false
    vkrSettings.hasFilter = false
    vkrSettings.hasGroupSort = false
    vkrSettings.hasSurnameSort = false
    vkrSettings.hasAttendanceNumberSort = false
    vkrSettings.studentsDocUrl = testStudentsDocUrl

    saveVKRSettings(JSON.stringify(vkrSettings))

    let propertyService = PropertiesService.getScriptProperties();
    let actualVKRSettingsJSON = propertyService.getProperty(vkrSettingsKey)
    let actualVKRSettings
    if (actualVKRSettingsJSON != null) {
      actualVKRSettings = JSON.parse(actualVKRSettingsJSON)
    } else {
      t.fail('VKR settings not found')
    }

    t.deepEqual(vkrSettings, actualVKRSettings, 'Check settings')
  })

  test('TestVKRFilling', function(t) {
    let vkrSettings = new Object()
    vkrSettings.hasAttendance = false
    vkrSettings.hasCommentsAttestation = false
    vkrSettings.hasRemarksAttestation = false
    vkrSettings.hasPointsSupervisor = false
    vkrSettings.hasPointsReviewer = false
    vkrSettings.hasOriginality = false
    vkrSettings.hasResult = false
    vkrSettings.hasComments = false
    vkrSettings.hasRemarks = false
    vkrSettings.hasFilter = false
    vkrSettings.hasGroupSort = false
    vkrSettings.hasSurnameSort = false
    vkrSettings.hasAttendanceNumberSort = false
    vkrSettings.studentsDocUrl = testStudentsDocUrl

    saveVKRSettings(JSON.stringify(vkrSettings))

    vkrCertification(vkrSettings)

    let certificationSheet = SpreadsheetApp.getActiveSheet()
    if (certificationSheet == null) {
      t.fail('Certification sheet not found')
    } else {
      t.equal("Группа", certificationSheet.getRange('A1').getValue(), 'Check add group column')
      t.equal("Студент", certificationSheet.getRange('B1').getValue(), 'Check add student column')
      t.equal("Номер аттестации", certificationSheet.getRange('C1').getValue(), 'Check add certification number column')
    }

    vkrGraduation(vkrSettings)

    let vkrSheet = SpreadsheetApp.getActiveSheet()
    if (vkrSheet == null) {
      t.fail('VKR sheet not found')
    } else {
      t.equal("Группа", vkrSheet.getRange('A1').getValue(), 'Check add group column')
      t.equal("Студент", vkrSheet.getRange('B1').getValue(), 'Check add student column')
      t.equal("Тема ВКР", vkrSheet.getRange('C1').getValue(), 'Check add theme column')
      t.equal("Руководитель", vkrSheet.getRange('D1').getValue(), 'Check add main person column')
      t.equal("Консультант", vkrSheet.getRange('E1').getValue(), 'Check add second person column')
      t.equal("Рецензент", vkrSheet.getRange('F1').getValue(), 'Check add feedback column')
      t.equal("Нормоконтроль", vkrSheet.getRange('G1').getValue(), 'Check add control column')
    }

    deleteSheetIfExist(certificationSheet)
    deleteSheetIfExist(vkrSheet)
  })

  test('TestCreateCustomSample', function(t) {
    deleteSheetIfExistByName('Шаблон')
    mySample()

    let customSampleSheet = getSheet('Шаблон')
    for(let i = 1; i < 5; i++) {
      for (let j = 1; j < 3; j++) {
        customSampleSheet.getRange(j, i).setValue('cell' + j + i)
      }
    }
    myDataForSampleAsJSON()

    let propertyService = PropertiesService.getUserProperties()
    let customSampleId = propertyService.getProperty(samplesQuantityKey)
    if (customSampleId == null) {
      t.fail('Custom samples does not exist.')
      return
    }
    customSampleId = parseInt(customSampleId)

    createMySampleFromJSON(customSampleId)
    let readSheetName = 'Мой шаблон ' + customSampleId
    let readSampleSheet = getSheet(readSheetName)

    var done = true
    for(let i = 1; i < 5; i++) {
      for (let j = 1; j < 3; j++) {
        if (readSampleSheet.getRange(j, i).getValue() != 'cell' + j + i) {
          done = false
        }
      }
    }
    t.equal(true, done, 'Cells is correct')

    propertyService.deleteProperty(prefixSampleKey + customSampleId)
    if (propertyService.getProperty(prefixSampleKey + customSampleId) == null) {
      propertyService.setProperty(samplesQuantityKey, parseInt(customSampleId) - 1)
    } else {
      t.fail('Could not delete test custom sample property')
    }

    deleteSheetIfExistByName(readSheetName)
  })

   test.finish()
}

function getSheet(sheetName) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  return activeSpreadsheet.getSheetByName(sheetName)
}

function deleteSheetIfExistByName(sheetName) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = getSheet(sheetName)
  if (sheet != null) {
    activeSpreadsheet.deleteSheet(sheet)
  }
}

function deleteSheetIfExist(sheet) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  activeSpreadsheet.deleteSheet(sheet)
}
