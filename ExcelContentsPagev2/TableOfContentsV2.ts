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
  // Set font bold to true for range A1 on selectedSheet
  tocWorkSheet.getRange("A1").getFormat().getFont().setBold(true);
  // Set font color to 4472C4 for range A1 on selectedSheet
  tocWorkSheet.getRange("A1").getFormat().getFont().setColor("4472C4");
  // Set font underline to true for range A1 on selectedSheet
  tocWorkSheet.getRange("A1").getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.single);


  //delete existing ToC tables
  tocWorkSheet.getRange("B:N").delete(ExcelScript.DeleteShiftDirection.left)
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
  tableoftablesRange = `${"'ToC'!" + 'E2:F' + jStr}`
  //tableoftablesRange = `${'E2:F'+jStr}`
  workbook.addTable(tableoftablesRange, true)

//Worksheets
  let reportCell = tocWorkSheet.getRange("B2")
  reportCell.setValue("Worksheet")
  workbook.addTable(reportCell, true)
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
    let chartAddressesVal:string
    //Loop through all worksheets to access each chart on each sheet
    allSheets.forEach((sheetobj1) => {
    SheetNameVal = sheetobj1.getName().toString()
    //console.log('SheetNameVal: ' +SheetNameVal)
    let sheetChartCollection = sheetobj1.getCharts()
    sheetChartCollection.forEach((sheetChart) => {
      chartNameVal = sheetChart.getName().toString()
       //console.log('M: ' + M)
       //console.log('chartNameVal: '+chartNameVal)
  chartCell.getOffsetRange(M + 1, 0).setValue(chartNameVal)
  let topLeftCell = getCellUnderChart(sheetChart, sheetobj1)
  chartAddressesVal = topLeftCell.getAddress().toString()
  chartAddressesVal = chartAddressesVal.split('!').splice(1, 1).toString()
        chartFormula = '=HYPERLINK("#' + "'" + sheetobj1.getName() + "'!" + chartAddressesVal + '","' + sheetobj1.getName() + '")'
  if (sheetobj1.getName() != 'ToC') {
     chartCell.getOffsetRange(M + 1, 1).setFormulaLocal(chartFormula)
  }
  M++
  })})
    let MStr: string = String(M + 2)
    let chartTableRange: string
    chartTableRange = `${"'ToC'!" + 'H2:I' + MStr}`
    workbook.addTable(chartTableRange, true)
  }
  
  //Ranges
    var RowCt = 1;
    var sheetName = "";
    let LinkFormula : string
    let tempFormula : string
    let bangPosition: number
    let sheeetNameStr: string
    let cellRefStr: string
    let tempFormulaArray: Array<string>
    tempFormulaArray = [''] //have to assign something to the array.
    let MyNames = workbook.getNames();
    let worksheets = workbook.getWorksheets()
    tocWorkSheet.getRange("K2:N2").setValues([["Scope", "Range Name", "RefersTo", "Visible"]]);
    for (let i = 0; i < MyNames.length; i++) {
      RowCt++
      tempFormula = MyNames[i].getFormula() //'=Sheet1!$B$26:$F$29
      tempFormulaArray = tempFormula.split(':')//=Sheet1!$B$26,$F$29
      tempFormula = tempFormulaArray[0].toString()//=Sheet1!$B$26
      tempFormula = tempFormula.substring(1)//Sheet1!$B$26
      bangPosition = tempFormula.search('!')//6
      sheeetNameStr = tempFormula.substring(0,bangPosition)//Sheet1
      cellRefStr = tempFormula.substring(bangPosition)
      LinkFormula = '=HYPERLINK("#' + "'" + sheeetNameStr + "'" + cellRefStr + '","' + MyNames[i].getName() +'")' 
      //console.log(LinkFormula)
      //=HYPERLINK("#'Sheet2'!A15","Sheet2")
  //
      tocWorkSheet.getCell(RowCt, 10).getResizedRange(0, 3).setValues([
        ["Workbook", MyNames[i].getName(), '', "'" + MyNames[i].getVisible()]]);
      tocWorkSheet.getCell(RowCt, 12).setFormula(LinkFormula)  
    }

    let newTable = workbook.addTable(tocWorkSheet.getRange("K2").getSurroundingRegion(), true);
    newTable.getRange().getFormat().autofitColumns();

  // Auto fit the columns of range B:N on selectedSheet
  tocWorkSheet.getRange("B:N").getFormat().autofitColumns()
  //Remove applied filter form all tables on ToC sheet.
  let ToCsheetTableCollection = tocWorkSheet.getTables()
  ToCsheetTableCollection.forEach((sheetTable) => {
    let tableNameVal = sheetTable.getName().toString()
    sheetTable.getAutoFilter().remove();
  })
console.log('Contents Page Created')
  
}

function getCellUnderChart(cht: ExcelScript.Chart, ws: ExcelScript.Worksheet): ExcelScript.Range {
  let topLeftCell = ws.getRange("A1");
  let i = 0;
  do {
    i++;
    topLeftCell = topLeftCell.getOffsetRange(0, i)
  }
  while (topLeftCell.getLeft() < cht.getLeft());
  i = 0;
  do {
    i++;
    topLeftCell = topLeftCell.getOffsetRange(i, 0)
  }
  while (topLeftCell.getTop() < cht.getTop());

  return topLeftCell.getOffsetRange(-1, -1);
}

