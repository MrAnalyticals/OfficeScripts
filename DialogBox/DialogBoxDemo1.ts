function main(workbook: ExcelScript.Workbook) {
let selectedSheet = workbook.getWorksheet('DialogBoxDemo');
//Insert this at the top of your main function
 Function.call(dataVal(workbook))
//use this code to display the dialog box in your code
selectedSheet.getCell(1,2).setValue(1)
selectedSheet.getCell(1, 2).select()// clear()//.showCard()
//Insert this immediately after your code
  Function.call(CleardataVal(workbook))
  sleepy(5)}

 function CleardataVal(workbook: ExcelScript.Workbook) {
  // Get the data validation object for C2:C2 in the current worksheet.
  let selectedSheet1 = workbook.getWorksheet('DialogBoxDemo');
  let dataValidation = selectedSheet1.getRange("C2:C2").getDataValidation();
  // Clear any previous validation to avoid conflicts.
  dataValidation.clear();
  selectedSheet1.getCell(1, 2).setValue(null)
  selectedSheet1.getCell(2, 2).select()
}

function dataVal(workbook: ExcelScript.Workbook){
  // Get the data validation object for C2:C2 in the current worksheet.
  let selectedSheet1 = workbook.getWorksheet('DialogBoxDemo');
  let dataValidation = selectedSheet1.getRange("C2:C2").getDataValidation();
  // Clear any previous validation to avoid conflicts.
  dataValidation.clear()
  const prompt: ExcelScript.DataValidationPrompt = {
    showPrompt: true,
    title: "Script dialog box",
    message: "The Dialog Demo Script completed successfully."
  }
  dataValidation.setPrompt(prompt)}

function sleepy(seconds: number): void {
  const waitUntil: number = new Date().getTime() + seconds * 1000;
  while (new Date().getTime() < waitUntil) {}}