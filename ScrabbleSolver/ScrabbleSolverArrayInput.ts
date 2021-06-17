function main(workbook: ExcelScript.Workbook, defns: definitionsArray[]) {
  let Solver = workbook.getWorksheet('Solver')
  if(defns.length==0){
    return
  }
let j: number = 0
let k: number = 0
for (let defnsItem of defns) {
  k = j + 8
  Solver.getCell(k, 2).setValue(defns[j])
  j++
}
  // Toggle auto filter on Solver
  Solver.getAutoFilter().apply("B8:C11100");
  // Create a new temporary sheet view
  //Solver.enterTemporaryNamedSheetView() //need this
  // Apply custom filter on Solver
  Solver.getAutoFilter().apply("C8", 1, { filterOn: ExcelScript.FilterOn.custom, criterion1: "<>Word does not exist" });
  console.log('ScrabbleSolverArrayInput Routine finished')
}
interface definitionsArray {
  definition: string
}
