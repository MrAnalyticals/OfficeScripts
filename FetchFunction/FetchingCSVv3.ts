async function main(workbook: ExcelScript.Workbook): Promise<string>  {
//*** Fetch example for csv return types ***//
//creating new worksheet and naming it
  try{
  let date: Date = new Date();
  let hoursStr: string
  let minutesStr: string
  let secondsStr: string
  if(date.getHours().toString().length == 1){
    hoursStr = '0'+date.getHours().toString()
  } else { hoursStr = date.getHours().toString()}
  if (date.getMinutes().toString().length == 1) {
    minutesStr = '0' + date.getMinutes().toString()
  } else { minutesStr = date.getMinutes().toString()}
  if (date.getSeconds().toString().length == 1) {
    secondsStr = '0' + date.getSeconds().toString()
  } else { secondsStr = date.getSeconds().toString() }
  let resultDate = `${'result' + hoursStr + minutesStr + secondsStr}`
  if(resultDate.length==5){
    resultDate="0"+resultDate}
  let newSheet = workbook.addWorksheet(resultDate);
    const target = 'https://raw.githubusercontent.com/treselle-systems/customer_churn_analysis/master/WA_Fn-UseC_-Telco-Customer-Churn.csv'; //file
  const res = await fetch(target)
  const data = await res.text();
  let responseArray = data.split('\n')  //string
  //Create a table
  let firstRow = responseArray[0].split(',')
  let lengthRow = firstRow.length
  //build rangeAddressStr
  let columnNo:number=0
  for (let headerItem of firstRow) {
    newSheet.getCell(0, columnNo).setValue(headerItem)
    columnNo++ }
  let endAddress: string = newSheet.getCell(0,columnNo-1).getAddress()
  let endAddressStr: string
  //Extract string after ':'
   endAddressStr = endAddress.substring(13); 
   endAddressStr = 'A1:' + endAddressStr
  let resultTable = workbook.addTable(newSheet.getRange(endAddressStr), true);
  resultTable.setPredefinedTableStyle("TableStyleLight9");
  //building rows array
  let rows: (string | boolean | number)[][] = [];
  let leninputArray: number = responseArray.length
    console.log('Length of result array: '+leninputArray+' rows.')
    console.log('cellCount: ' + lengthRow * leninputArray)
  let rowArray: (string | boolean | number)[] = [];
   for (let repo of responseArray) {
     rowArray = repo.split(',')
     resultTable.addRow(-1, rowArray)}
  console.log('The fetch return dataset was: ' + data.length + ' bytes in size.')
} 
catch (Error) {
    let errorMsg:string = Error.toString().substring(10)
    if (errorMsg.trim()=='Failed to fetch'){
    console.log('Error: The data set is too large to fetch.')
    return 'Routine not completed'}
    else{ return 'Routine completed'}
}}
