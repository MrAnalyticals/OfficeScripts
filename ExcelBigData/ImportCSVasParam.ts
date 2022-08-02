function main(workbook: ExcelScript.Workbook, csvParam: string) {
//data: string[][]
//console.log('csvParam: ' + csvParam)
//*** Outputs an input CSV parameter to a new worksheet ***//
//creating new worksheet and naming it //Note: 500,000 cells seems to be limit.
try{
  //let csvParam: string[]= ["location_name, location_id, document_count, percentage","Poland, pl, 2020, 29.27","Germany, de, 1279, 18.53","Greece, gr, 526, 7.62","Spain, es, 480, 6.95","United Kingdom, gb, 381, 5.52","Italy, it, 292, 4.23","Romania, ro, 214, 3.1","Turkey, tr, 145, 2.1","South Africa, za, 125, 1.81","Switzerland, ch, 119, 1.72","Russia, ru, 106, 1.54","France, fr, 102, 1.48","Netherlands, nl, 92, 1.33","Croatia, hr, 88, 1.27","Nigeria, ng, 78, 1.13","Portugal, pt, 69, 1","Austria, at, 67, 0.97","Belgium, be, 63, 0.91","Bulgaria, bg, 52, 0.75","Latvia, lv, 47, 0.68","Ghana, gh, 44, 0.64","Ukraine, ua, 43, 0.62","Israel, il, 36, 0.52","Denmark, dk, 35, 0.51","Cyprus, cy, 33, 0.48","Czechia, cz, 28, 0.41","Ireland, ie, 27, 0.39","Slovakia, sk, 27, 0.39","Hungary, hu, 25, 0.36","Luxembourg, lu, 22, 0.32","Moldova, md, 22, 0.32","Finland, fi, 20, 0.29","Pakistan, pk, 19, 0.28","Albania, al, 16, 0.23","Lithuania, lt, 13, 0.19","Sweden, se, 12, 0.17","Uganda, ug, 12, 0.17","Egypt, eg, 9, 0.13","Slovenia, si, 9, 0.13","Bosnia & Herzegovina, ba, 8, 0.12","Jordan, jo, 8, 0.12","Kazakhstan, kz, 8, 0.12","Norway, no, 8, 0.12","Morocco, ma, 7, 0.1","Saudi Arabia, sa, 6, 0.09","Uruguay, uy, 6, 0.09","Azerbaijan, az, 4, 0.06","Iceland, is, 4, 0.06","Angola, ao, 3, 0.04","Georgia, ge, 3, 0.04","Iran, ir, 3, 0.04","Cambodia, kh, 3, 0.04","Tunisia, tn, 3, 0.04","Tanzania, tz, 3, 0.04","Zambia, zm, 3, 0.04","Zimbabwe, zw, 3, 0.04","Bahrain, bh, 2, 0.03","Central African Republic, cf, 2, 0.03","Estonia, ee, 2, 0.03","Montenegro, me, 2, 0.03","Macedonia, mk, 2, 0.03","Namibia, na, 2, 0.03","French Polynesia, pf, 2, 0.03","Rwanda, rw, 2, 0.03","Armenia, am, 1, 0.01","Botswana, bw, 1, 0.01","Congo - Brazzaville, cg, 1, 0.01","Cameroon, cm, 1, 0.01","Algeria, dz, 1, 0.01","Gambia, gm, 1, 0.01"]
//to do: remove last ',' if it is present
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
  //  const target = 'https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/FetchFunction/MXSales60000.csv'; //file
  //const res = await fetch(target)
  //const data = await res.text();
      //let csvParamStr = csvParam.toString()
      //if (csvParamStr.length == 0) {return}
  //let responseArray = csvParam.split('","')
  let responseArray = csvParam.split('\r\n')
  //let responseArray = data.split('\n')  //string
  //Create a table
  let firstRow = responseArray[0].split(',')
  console.log('responseArray[0]: ' + responseArray[0])
  //console.log('firstRow: ' + firstRow) 
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
     rows.push(rowArray)}
  //console.log('responseArray: '+ responseArray.toString())
  const resultRange = newSheet.getRange('A1').getResizedRange(leninputArray-1, lengthRow-1);
  resultRange.setValues(rows);
  newSheet.getRange(endAddressStr).getFormat().autofitColumns()
  resultTable.getAutoFilter().remove()
}
catch (Error) {
    let errorMsg:string = Error.toString().substring(10)
    console.log(Error)
    if (errorMsg.trim()=='Failed to fetch'){
    console.log('Error: The data set is too large to fetch.')
    return 'Routine not completed'}
    else{ return 'Routine completed'}
}}