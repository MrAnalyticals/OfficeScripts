function main(workbook: ExcelScript.Workbook) {
  let Password = workbook.getWorksheet('Password')
  let x: number
  let mapStr: string = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
  let D2Str: string = Password.getRange("D2").getValues().toString()
  let B3Str: string = Password.getRange("B3").getValues().toString()
  let D2:number = parseInt(D2Str)
  let GeneratedPosition: number
  let StrVal: string
  let StrValStr: string =''
  for (let x = 0; x < D2; x++) {
  GeneratedPosition = getRandomInt(62)
  StrVal = mapStr.substring(GeneratedPosition,GeneratedPosition+1)
  StrValStr = StrValStr + StrVal}
  if(B3Str==''){
    Password.getRange("B3").setValue(StrValStr) 
  }
  else{
    Password.getRange("B2").getRangeEdge(ExcelScript.KeyboardDirection.down).getOffsetRange(1, 0).setValue(StrValStr)
  }
  return
}

  function getRandomInt(max:number) {
    return Math.floor(Math.random() * max);
  }