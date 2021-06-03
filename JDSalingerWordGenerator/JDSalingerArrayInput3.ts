function main(workbook: ExcelScript.Workbook, defns: definitionsArray[]) {
  //console.log('defns : ' + defns)
  let JDSalinger = workbook.getWorksheet('JDSalinger')
  let defns1 = ["Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Royal Garrison Artillery.", "Word does not exist", "Word does not exist", "Word does not exist", "A sweet, especially an inexpensive, one of a type intended mainly for children.", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist", "Word does not exist"]//used for testing only   //let defnsStr = defns1.toString()  //use this line for testing
  //the static array above
  let defnsStr = defns.toString()
  if(defnsStr.length==0){
    return
  }
  console.log('defnsStr: ' + defnsStr)
  let defnsValR = defnsStr.replace(', ', ' ')
  //Same as replaceAll below
  let defnsValR1 = defnsValR.split(', ').join(' ')
  console.log('defnsValR1 joined: ' + defnsValR1)
  let defnsArray = defnsValR1.split(',')
  let j: number = 0
  j = 4
  for (let defnsVal of defnsArray) {
    JDSalinger.getCell(j, 2).setValue(defnsVal)
    j++
    if(j==24){
      break
    }
  }
let k: number
  console.log('JDSalingerArrayInput3 Routine finished')
}

interface definitionsArray {
  definition: string
}
