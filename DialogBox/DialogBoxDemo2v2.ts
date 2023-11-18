function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getWorksheet('DialogBoxDemo2');
  let chartSheet = workbook.getWorksheet('SalesInvoicesChart');  
  //Insert this at the top of your main function
  let dialogMsgStr:string //declare the dialog message
  let cellReference:string //use a single cell reference
  let invoiceSentTableRowCount = selectedSheet.getTable("Table1").getColumnByName("Country").getRangeBetweenHeaderAndTotal().getRowCount()
  if (invoiceSentTableRowCount > 0){
      //create the chart on sheet DialogBoxDemo2
    let ChartObj1 = chartSheet.getCharts()[0]
    let chartImage = ChartObj1.getImage()
    let chartImageShape = selectedSheet.addImage(chartImage)
    chartImageShape.setLeft(0);
    chartImageShape.setTop(198.75);
    chartImageShape.setWidth(369);
    chartImageShape.setHeight(225);

    dialogMsgStr = "Invoice Table Chart Created successfully. There was more than zero invoices this month!"
    cellReference = 'C2'
    DisplayDialog(workbook, dialogMsgStr, cellReference)
}

else{
    dialogMsgStr = "Invoice Table Chart Was not Created. There were no invoices this month!"
    cellReference = 'C2'
    DisplayDialog(workbook, dialogMsgStr, cellReference)
}}  

function CleardataVal(workbook: ExcelScript.Workbook,cellRefs:string) {
  // Get the data validation object for C2:C2 in the current worksheet.
  let selectedSheet1 = workbook.getWorksheet('DialogBoxDemo2');
  let dataValidation = selectedSheet1.getRange(cellRefs + ":"+ cellRefs).getDataValidation();
  // Clear any previous validation to avoid conflicts.
  dataValidation.clear();
  selectedSheet1.getCell(1, 2).setValue(null)
  selectedSheet1.getCell(2, 2).select()
}

function DisplayDialog(workbook: ExcelScript.Workbook,dialogMsg:string,cellRef:string) {
  // Get the data validation object for C2:C2 in the current worksheet.
  let selectedSheet1 = workbook.getWorksheet('DialogBoxDemo2');
  let dataValidation = selectedSheet1.getRange(cellRef + ":" + cellRef).getDataValidation();
  // Clear any previous validation to avoid conflicts.
  dataValidation.clear()
  const prompt: ExcelScript.DataValidationPrompt = {
    showPrompt: true,
    title: "DialogBoxDemo2 Message",
    message: dialogMsg
  }
  dataValidation.setPrompt(prompt)
  selectedSheet1.getCell(1, 2).setValue(1)
  selectedSheet1.getCell(1, 2).select()// clear()//.showCard()
  CleardataVal(workbook,"c2")
  sleepy(5)
}


function sleepy(seconds: number): void {
  const waitUntil: number = new Date().getTime() + seconds * 1000;
  while (new Date().getTime() < waitUntil) { }
}