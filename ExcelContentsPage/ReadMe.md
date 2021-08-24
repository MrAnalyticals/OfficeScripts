**Create an Excel Online Contents Page **

Using Office Scripts macro language you can create a Contents page that lists the worksheets, tables and charts in your work book. Simply past the following code into the code pane in Excel, save the script and run it.

![image](https://user-images.githubusercontent.com/47678539/130545223-3e33961f-8a0b-4ab8-9d51-a854c9eb281e.png)

**Excel Office Script Code**

function main(workbook: ExcelScript.Workbook) {
  let allSheets = workbook.getWorksheets()
  let ToCExists: boolean
  let arrayOfSheetNames: Array<string>
  arrayOfSheetNames = [''] //have to assign something to the array.
  let arrayOfTableAddresses: Array<string>
  arrayOfTableAddresses = [''] //have to assign something to the array.
  let arrayOfTableAddressesVal : string
  let k: number = 0

  allSheets.forEach((sheetobj) => {
    arrayOfSheetNames[k] = sheetobj.getName()
    if (sheetobj.getName().toString() == 'ToC') {
      ToCExists = true
      let ToCSheet = workbook.getWorksheet('ToC')
      ToCSheet.setPosition(allSheets.length-1)
    }
  })
  if (ToCExists != true) {workbook.addWorksheet("ToC")} 
  let tocWorkSheet = workbook.getWorksheet('ToC')
  tocWorkSheet.getRange('A1').setValue('Contents Page')
  //delete existing ToC tables
  tocWorkSheet.getRange("B:I").delete(ExcelScript.DeleteShiftDirection.left)
  //Tables
  let allTables = workbook.getTables()
  let sheetname:string
  let cellref:string
  let formula:string
  let j:number
  if (allTables.length !== 0) {
    let tableCell = tocWorkSheet.getRange("E2")
    tableCell.setValue("Table")
    tableCell.getOffsetRange(0, 1).setValue("Link")
    for (j = 0; j < allTables.length; j++) {
      tableCell.getOffsetRange(j + 1, 0).setValue(allTables[j].getName())
      //adding a table range address to an array
      arrayOfTableAddressesVal = allTables[j].getRange().getAddress().toString()
      arrayOfTableAddressesVal = arrayOfTableAddressesVal.split(':').splice(0, 1).toString() 
      sheetname = arrayOfTableAddressesVal.split('!').splice(0, 1).toString()
      cellref = arrayOfTableAddressesVal.split('!').splice(1, 1).toString()
      arrayOfTableAddresses[j] = '=HYPERLINK("#' + "'" + sheetname + "'!" + cellref + '","' + allTables[j].getName()+ '")'
      tableCell.getOffsetRange(j + 1, 1).setFormulaLocal(arrayOfTableAddresses[j])
    }
  }
  let jStr:string = String(j+2)
  let tableoftablesRange:string
  tableoftablesRange = `${'E2:F'+jStr}`
  workbook.addTable(tableoftablesRange, true)

//Worksheets
  let reportCell = tocWorkSheet.getRange("B2")
  reportCell.setValue("Worksheet")
  let newTable = workbook.addTable(reportCell, true)
  reportCell.getOffsetRange(0, 1).setValue("Link")
  for (let i = 0; i < allSheets.length; i++) {
    reportCell.getOffsetRange(i + 1, 0).setValue(allSheets[i].getName())
  }
  reportCell.getOffsetRange(1, 1).setFormulaR1C1("=HYPERLINK(\"#'\"&RC[-1]&\"'!A1\",RC[-1])")
//Charts
//Test if Charts exist
  let countOfCharts:number=0
  allSheets.forEach((sheetobj) => {
    let sheetsCollection = sheetobj.getCharts()
    sheetsCollection.forEach((sheetChart) => {
      countOfCharts++
    })
  })
  if(countOfCharts!=0){
    let M:number =0
    let chartCell = tocWorkSheet.getRange("H2")
    chartCell.setValue("Chart")
    chartCell.getOffsetRange(0, 1).setValue("Link")
    let chartNameVal:string
    let chartFormula:string
    let SheetNameVal: string

    //Loop through all worksheets to access each chart on each sheet
    allSheets.forEach((sheetobj1) => {
      SheetNameVal = sheetobj1.getName().toString()
        console.log('SheetNameVal: ' +SheetNameVal)
      chartFormula = '=HYPERLINK("#' + "'" + sheetobj1.getName() + "'!A1" + '","' + sheetobj1.getName() + '")'
      if(sheetobj1.getName()!='ToC'){
      chartCell.getOffsetRange(M + 1, 1).setFormulaLocal(chartFormula)}
      let sheetChartCollection = sheetobj1.getCharts()
      sheetChartCollection.forEach((sheetChart) => {
        chartNameVal = sheetChart.getName().toString()
        console.log('M: ' + M)
        console.log('chartNameVal: '+chartNameVal)
        chartCell.getOffsetRange(M + 1, 0).setValue(chartNameVal)
            M++
        })
      })
    let MStr: string = String(M + 2)
    let chartTableRange: string
    chartTableRange = `${'H2:I' + jStr}`
    workbook.addTable(chartTableRange, true)
  }
  
  // Auto fit the columns of range B:J on selectedSheet
  tocWorkSheet.getRange("B:J").getFormat().autofitColumns()
  //Remove applied filter form all tables on ToC sheet.
  let ToCsheetTableCollection = tocWorkSheet.getTables()
  ToCsheetTableCollection.forEach((sheetTable) => {
    let tableNameVal = sheetTable.getName().toString()
    sheetTable.getAutoFilter().remove();
  })
console.log('Contents Page Created')
  
}


