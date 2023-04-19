// import convert-excel-to-json module
// NOTE: Default null to check if module installed or not
let excelToJson = null
let jsonToExcel = null

try {
  excelToJson = require('convert-excel-to-json')
  jsonToExcel = require('json-as-xlsx')
} catch (e) {
  console.error(e)
}

// Function ---------------------------------------------------------------------------------------
function correctDate (arr) {
  const result = []

  for (let obj of arr) {
    if (!obj['Start Date'] || !obj['End Date']) {
      result.push(obj)
      continue
    }

    let newObj = {...obj}

    const startDate = new Date(obj['Start Date'])
    const endDate = new Date(obj['End Date'])

    startDate.setUTCFullYear(startDate.getFullYear())
    startDate.setUTCMonth(startDate.getMonth())
    startDate.setUTCDate(startDate.getDate())
    startDate.setUTCHours(0)
    startDate.setUTCMinutes(0)
    startDate.setUTCSeconds(0)

    endDate.setUTCFullYear(endDate.getFullYear())
    endDate.setUTCMonth(endDate.getMonth())
    endDate.setUTCDate(endDate.getDate())
    endDate.setUTCHours(0)
    endDate.setUTCMinutes(0)
    endDate.setUTCSeconds(0)

    newObj['Start Date'] = startDate
    newObj['End Date'] = endDate

    result.push(newObj)
  }

  console.log('Correct Date Finished !!!!!!!!')
  return result
}

function parseNames (arr) {
  // Matches any word character
  const regex = /\b[\w\s]+\b/g
  const result = []

  for (let obj of arr) {
    // If 'Assigned To' is empty add it and skip parse process
    if (!obj['Assigned To']) {
      result.push(obj)
      continue
    }

    const names = obj['Assigned To'].match(regex)
    for (let name of names) {
      let newObj = {...obj}
      newObj['Assigned To'] = name
      result.push(newObj)
    }
  }

  console.log('Parse Name Finished !!!!!!!!')
  return result
}

function parseDateRanges (arr) {
  let result = []

  for (let obj of arr) {
    let startDate = new Date(obj['Start Date'])
    let endDate = new Date(obj['End Date'])
    let currentYear = startDate.getFullYear()
    let currentMonth = startDate.getMonth()
    let targetYear = endDate.getFullYear()
    let targetMonth = endDate.getMonth()

    if (!obj['Start Date'] || !obj['End Date']) {
      result.push(obj)
      continue
    } else if (currentYear !== targetYear) {
      result.push(obj)
      continue
    }

    while ((currentYear === targetYear) &&
           (currentMonth <= targetMonth)) {
      let rangeEnd = null

      if (currentMonth < targetMonth) {
        const lastDay = new Date(startDate)

        // Set the date to the last day of month
        lastDay.setFullYear(lastDay.getUTCFullYear())
        lastDay.setUTCMonth(lastDay.getUTCMonth() + 1)
        lastDay.setUTCDate(0)

        rangeEnd = lastDay
      } else {
        rangeEnd = endDate
      }

      let newObj = {...obj}
      newObj['Start Date'] = startDate.toISOString().slice(0, 10)
      newObj['End Date'] = rangeEnd.toISOString().slice(0, 10)

      result.push(newObj)

      startDate = setNextMonthFirstDay(startDate)
      currentMonth = startDate.getMonth()
      currentYear = startDate.getFullYear()
    }
  }

  console.log('Parse Date Finished !!!!!!!!')
  return result
}

function setNextMonthFirstDay (time) {
  const date = new Date(time)

  // Get next month's index(0 based)
  const nextMonth = date.getMonth() + 1
  // Get year
  const year = date.getFullYear() + (nextMonth === 12 ? 1: 0)

  date.setUTCFullYear(year)
  date.setUTCMonth(nextMonth%12)
  date.setUTCDate(1)
  date.setUTCHours(0)
  date.setUTCMinutes(0)
  date.setUTCSeconds(0)

  return date
}

// ------------------------------------------------------------------------------------------------

// ---Main-----------------------------------------------------------------------------------------

// Check input data and SDK is installed or not
if (process.argv.length <= 2) {
  console.error('argument not enough')
  process.exit(1)
} else if (process.argv.length > 2 && (excelToJson === null || jsonToExcel === null)) {
  console.error('Please use \'yarn install\' to install module')
  process.exit(1)
}

if (process.argv.length > 2 && excelToJson) {
  // Parse xlsx to JSON
  const result = excelToJson({
    columnToKey: {
      '*': '{{columnHeader}}'
    },
    header:{
        rows: 1
    },
    sourceFile: process.argv[2].toString()
  })

  // Separate every sheets' names in column 'Assigned To'
  const parsedData = {}

  Object.keys(result).forEach(sheet => {
    let tempArray = correctDate(result[sheet])
    tempArray = parseNames(tempArray)
    tempArray = parseDateRanges(tempArray)

    parsedData[sheet] = tempArray
  })

  // Setting xlsx data including sheet, column and content
  let xlsxData = []
  Object.keys(parsedData).forEach(sheet => {
    let sheetData = {}

    const columns = []
    Object.keys(parsedData[sheet][0]).forEach(key => {
      let tempData = {}

      tempData.label = key
      tempData.value = key

      columns.push(tempData)
    })

    sheetData.sheet = sheet
    sheetData.columns = columns
    sheetData.content = parsedData[sheet]

    xlsxData.push(sheetData)
  })

  // Setting xlsx config
  // ref: https://www.npmjs.com/package/json-as-xlsx
  let xlsxConfig = {
    fileName: "parsedData", // Name of the resulting spreadsheet
    extraLength: 3, // A bigger number means that columns will be wider
  }

  // Set JSON to xlsx and download
  // NOTE: xlsx is in the file where is same file as main.js
  jsonToExcel(xlsxData, xlsxConfig)
}

// ------------------------------------------------------------------------------------------------