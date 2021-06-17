function main(workbook: ExcelScript.Workbook): wordsArray[] {
  let Solver = workbook.getWorksheet('Solver')
  let returnArray: wordsArray[] = []
  let randomSalingerWord: String = ''
  let randomLetterNumber: number
  let randomLetter: string
  let test: string = ''
  let testy: string = ''
  let d4 = Solver.getRange('d4').getValue().toString()
  let d5 = Solver.getRange('d5').getValue().toString()
  console.log('d4: ' + d4)
  if (d5 == '1'){
    Solver.getAutoFilter().clearCriteria()
    Solver.getRange("d5").clear(ExcelScript.ClearApplyTo.contents)
  console.log('D5 cleared')
  Solver.getRange("B9:C11000").clear(ExcelScript.ClearApplyTo.contents)
let D4Str: string = d4.toString()
//let returnArray = []
//Runs anagram for all characters. Level 1
if(D4Str.length <3){
  return 
}
returnArray = allAnagrams(workbook,D4Str)
let D4StrLen = D4Str.length
// D4StrLen - 1 characters
let D4StrLenMinus1 = D4StrLen - 1
let shortWordsArray = []
let allAnagramsArray = []
//Iterate through the number of characters to find anagrams for smaller words
if (D4StrLenMinus1 < 2) {
      return
}
if (D4StrLenMinus1 > 2){
let k: number
shortWordsArray = subWords(workbook, D4Str)
let shortWordsArrayLen:number = shortWordsArray.length
 for (let k = 0; k < shortWordsArrayLen; k++) { //Level 2
  allAnagramsArray = allAnagrams(workbook, shortWordsArray[k])
  //console.log('k allAnagramsArray: ' + `${k + allAnagramsArray[k]}`)
  //console.log(k +'k allAnagramsArray: ' + allAnagramsArray)
  returnArray = returnArray.concat(allAnagramsArray)
  }
}
if (D4StrLenMinus1 > 2) {
let l:number
let shortWordsArray1 = []
let shortWordsArray2 = []
shortWordsArray1 = subWords(workbook, D4Str)
let shortWordsArrayLen: number = shortWordsArray1.length
if (shortWordsArrayLen > 2) {
for (let l = 0; l < shortWordsArrayLen-1; l++) { //Level 3
  shortWordsArray2 = subWords(workbook, shortWordsArray1[l])
  if (shortWordsArray2[0].length > 2) {  
  allAnagramsArray = allAnagrams(workbook, shortWordsArray2[l])
  //console.log(l + 'l allAnagramsArray Level 3: ' + `${l + allAnagramsArray[l]}`)
  //console.log(l + 'l allAnagramsArray Level 3: ' + allAnagramsArray)
  returnArray = returnArray.concat(allAnagramsArray)
  }
}}

let m: number
let shortWordsArray3 = []
if (shortWordsArrayLen > 2) {
  for (let m = 0; m < shortWordsArray2.length - 1; m++) {//Level 4
    shortWordsArray3 = subWords(workbook, shortWordsArray2[m])
    if (shortWordsArray3[0].length > 2) {
      allAnagramsArray = allAnagrams(workbook, shortWordsArray3[m])
      //console.log(m + 'm allAnagramsArray Level 4: ' + `${m + allAnagramsArray[m]}`)
      //console.log(m + 'm allAnagramsArray Level 4: ' + allAnagramsArray)
      returnArray = returnArray.concat(allAnagramsArray)
    }
  }
}
let n: number
let shortWordsArray4 = []
if (shortWordsArrayLen > 2) {
  for (let n = 0; n < shortWordsArray3.length - 1; n++) {//Level 5
    shortWordsArray4 = subWords(workbook, shortWordsArray3[n])
    if (shortWordsArray4[0].length > 2) {
      allAnagramsArray = allAnagrams(workbook, shortWordsArray4[n])
      //console.log(n + 'n allAnagramsArray Level 5: ' + `${n + allAnagramsArray[n]}`)
      //console.log(n + 'n allAnagramsArray Level 5: ' + allAnagramsArray)
      returnArray = returnArray.concat(allAnagramsArray)
    }
  }
}
}
//Count number of letters in supplied word
//console.log('returnArray: ' + returnArray)
let returnArrayItem: number
let countOfThousands: number = returnArray.length/1000 //1
let countOfThousandsRem: number=  returnArray.length % 1000 //584
//console.log('countOfThousands: ' + countOfThousands)
//output returnArray to worksheet in chunks of 1000 to avoid the overload error.
//let rowVal: number = 8
for (let thousandsCounter = 1; thousandsCounter < countOfThousands + 1; thousandsCounter++){
  for (let returnArrayItem = ((thousandsCounter-1)*1000);returnArrayItem < thousandsCounter * 1000; returnArrayItem++){
    Solver.getCell(returnArrayItem+8, 1).setValue(returnArray[returnArrayItem])      
  }
  console.log('thousandsCounter: ' + thousandsCounter)
}
//returnArray.forEach((cellobj) => {
//  Solver.getCell(rowVal,1).setValue(cellobj)      
//  returnArrayItem++
//})
//remove any duplicates //10952
Solver.getRange("B8:C11000").removeDuplicates([0], true)
//build returnArray
console.log('Started to build returnArray directly from worksheet')
let builtArray= []
let outputRange = Solver.getRange("B9:B11000").getValues()
let RowCounter: number = 0
for (let Cellval of outputRange) {
  if (Cellval.toString() == '') {
    break
  } else {
    RowCounter++
  }
  }
//let rowNumber = RowCounter.toString
RowCounter = RowCounter+8
console.log('rowNumber: ' + RowCounter)
let builtArrayRange = 'B9:B' + RowCounter
builtArray = Solver.getRange(builtArrayRange).getValues()
// remove [" "] from each value?
let builtArrayStr = builtArray.toString() //asdfert
builtArrayStr = builtArrayStr.replace('["','')
builtArrayStr = builtArrayStr.replace('"]', '')
let builtArrayNew = []
builtArrayNew = builtArrayStr.split(',')
//
returnArray =  builtArrayNew
console.log('Length of returnArray: ' + returnArray.length)
console.log('returnArray: ' + returnArray)
// Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range B:B on selectedSheet
Solver.getRange("B:B").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
Solver.getRange("B:B").getFormat().setIndentLevel(0);
// Set horizontal alignment to ExcelScript.HorizontalAlignment.left for range B1:B7 on selectedSheet
Solver.getRange("B1:B7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
return returnArray
}}

interface wordsArray {
  word: string
}

function allAnagrams(workbook: ExcelScript.Workbook,inputWord:string) {
  let shortwordArray = []
  if (inputWord.length < 2) {
    return [inputWord]
  } else {
    let allAnswers = []
    for (let i = 0; i < inputWord.length; i++) {
      let letter = inputWord[i]
      let shorterWord = inputWord.substr(0, i) + inputWord.substr(i + 1, inputWord.length - 1)
      shortwordArray = allAnagrams(workbook, shorterWord)
      for (let j = 0; j < shortwordArray.length; j++) {
        allAnswers.push(letter + shortwordArray[j])
      }
    }
    return allAnswers
  }
}

function subWords(workbook: ExcelScript.Workbook, inputString: string) {
  //number param. currently not being used.
  let shorterwordsArray = []
  let i: number
  let j: number
  let loopString: string
  let inputStringLen = inputString.length
  //inputString[i]  //obtains(reads)each letter in the string
  //Start from character 1 dropping each nth character, from the left, through the entire input string.
  let newShortWord = []
  let newWordSplitA = []
  let newShortWordSplit = inputString.split('') //asdfert
  let newInputString:string = inputString
  for (let i = 0; i < inputStringLen; i++) {
  //Use substring methd to extract ith elmnt from inputstring
  newInputString = inputString
    //console.log('newShortWordSplit[i]: ' + newShortWordSplit[i])
  let nSWS:string = newShortWordSplit[i].toString()
  newInputString = newInputString.replace(nSWS,'')    
    //loopString = loopString + newInputString
    //console.log('loopString: ' + loopString)
    newShortWord.push(newInputString)
    loopString = ''
  }
  //console.log('newShortWord: ' + newShortWord)  
  //newShortWord: undefinedsdfert,undefinedsdfertadfert,undefinedsdfertadfertasfert,undefinedsdfertadfertasfertasdert,undefinedsdfertadfertasfertasdertasdfrt,undefinedsdfertadfertasfertasdertasdfrtasdfet,undefinedsdfertadfertasfertasdertasdfrtasdfetasdfer
 //newShortWord: undefinedsdfert, undefinedsdfertdfert, undefinedsdfertdfertfert, undefinedsdfertdfertfertert, undefinedsdfertdfertfertertrt, undefinedsdfertdfertfertertrtt, undefinedsdfertdfertfertertrtt
return newShortWord
}