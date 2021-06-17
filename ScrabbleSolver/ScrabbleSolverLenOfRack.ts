function main(workbook: ExcelScript.Workbook):string
{
 let Solver = workbook.getWorksheet('Solver')
 let RackLen = Solver.getCell(3,4).getValue().toString().length
 if(RackLen>7){
 return 'Err'
 }
 else{return 'Good'}
}
