function main(workbook: ExcelScript.Workbook): wordsArray[] {
  let JDSheet = workbook.getWorksheet('JDSalinger')
  let wordsArrayObjects: wordsArray[] = []
  let randomSalingerWord: String = ''
  let randomLetterNumber: number
  let randomLetter: string
  let test: string = ''
  let testy: string = ''
  let g1 = JDSheet.getRange('g1').getValue().toString()
  console.log('g1: ' + g1)
  if (g1 == '1'){
    JDSheet.getRange("g1").clear(ExcelScript.ClearApplyTo.contents)
  JDSheet.getRange("b5:b24").clear(ExcelScript.ClearApplyTo.contents)

for(let j = 4; j<24;j++){
  for (let i = 1; i < 4; i++) {
    randomLetterNumber = Math.floor(Math.random() * 26) + 1
    while (randomLetterNumber=26) {
      randomLetterNumber = Math.floor(Math.random() * 26) + 1
      if(randomLetterNumber!=26)
      break
    }
    //testy = randomLetterNumber.toString() + testy
    JDSheet.getCell(1, 5).setFormula("=CHAR(" + (randomLetterNumber + 65).toString() + ")")
    randomLetter = JDSheet.getCell(1, 5).getValue().toString()
    randomSalingerWord = lowercase(workbook, randomLetter) + randomSalingerWord
  }
  JDSheet.getCell(j, 1).setValue(randomSalingerWord)
  //let word: string
  let word = randomSalingerWord
  wordsArrayObjects.push({word: word as string})
  randomLetter = ''
  randomSalingerWord = ''
  }
  JDSheet.getRange("F2").clear(ExcelScript.ClearApplyTo.contents)
  } 
  //console.log('g1 = ""')
  return wordsArrayObjects
}

function lowercase(workbook: ExcelScript.Workbook, letter: string): string {
  switch (letter) {
    case 'A':
      return 'a'
    case 'B':
      return 'b'
    case 'C':
      return 'c'
    case 'D':
      return 'd'
    case 'E':
      return 'e'
    case 'F':
      return 'f'
    case 'G':
      return 'g'
    case 'G':
      return 'g'
    case 'H':
      return 'h'
    case 'I':
      return 'i'
    case 'J':
      return 'j'
    case 'K':
      return 'k'
    case 'L':
      return 'l'
    case 'M':
      return 'm'
    case 'N':
      return 'n'
    case 'O':
      return 'o'
    case 'P':
      return 'p'
    case 'Q':
      return 'i'
    case 'R':
      return 'r'
    case 'S':
      return 's'
    case 'T':
      return 't'
    case 'U':
      return 'u'
    case 'V':
      return 'v'
    case 'W':
      return 'w'
    case 'X':
      return 'e'
    case 'Y':
      return 'y'
    case 'Z':
      return 'a'
  }
}


interface wordsArray {
  word: string
}
