async function main(workbook: ExcelScript.Workbook): Promise<string | StringConstructor> {
  // Load the "HardWords" table from the workbook
  let hardWordsSheet = workbook.getWorksheet("Answers")
  let spellingsSheet = workbook.getWorksheet("Spellings")
  let hardWordsTable = hardWordsSheet.getTable("HardWords")
  let table1 = hardWordsSheet.getTable("table1")
  
  // Select 5 random words from the "HardWords" table
  let selectedWords: string[] = []
  let y: number = 0
  //populate the selectedWords array with words from the Hardwords table
  while (selectedWords.length < 5) {
    let rowIndex: number = Math.floor(Math.random() * hardWordsTable.getRangeBetweenHeaderAndTotal().getRowCount())
    let hardWord: string = hardWordsSheet.getRange('B' + rowIndex).getValue().toString()
    if (hardWord !== undefined) {
      selectedWords[y] = hardWord
      y++
    }
  }
  let difficulty: number = 2
  //retrieve the difficulty level
  let diffCellVal: string = spellingsSheet.getRange('h3').getValue().toString()
  console.log('diffCellVal: ' + diffCellVal)
  let diffCellVal1: number
  if (diffCellVal === '1' || diffCellVal === '2' || diffCellVal === '3') {
    diffCellVal1 = parseInt(diffCellVal)
  }
  else {
    diffCellVal1 = 2
  }
  // Replace two random characters in each selected word with asterisks
  let transformedWords = selectedWords.map((word) => {
    let indexes: number[] = []
    while (indexes.length < diffCellVal1) {
      let index = Math.floor(Math.random() * word.length)
      if (indexes.indexOf(index) === -1) {
        indexes.push(index)
      }
    }
    let transformedWord: string = ""
    for (let i = 0; i < word.length; i++) {
      if (indexes.indexOf(i) !== -1) {
        transformedWord += "*";
      } else {
        transformedWord += word.charAt(i);
      }
    }
    return transformedWord;
  });

  // Write the transformed words to the "Answers" worksheet
  hardWordsSheet.getRange("D1184").setValue(selectedWords[0])
  hardWordsSheet.getRange("D1185").setValue(selectedWords[1])
  hardWordsSheet.getRange("D1186").setValue(selectedWords[2])
  hardWordsSheet.getRange("D1187").setValue(selectedWords[3])
  hardWordsSheet.getRange("D1188").setValue(selectedWords[4])
  // sort the Table1
  table1.getSort().apply([{ key: 0, ascending: true }]);
  // Write the transformed words to the "Spellings" worksheet
  spellingsSheet.getRange("B3").setValue(transformedWords[0])
  spellingsSheet.getRange("B4").setValue(transformedWords[1])
  spellingsSheet.getRange("B5").setValue(transformedWords[2])
  spellingsSheet.getRange("B6").setValue(transformedWords[3])
  spellingsSheet.getRange("B7").setValue(transformedWords[4])


  // Find definition of the 5 words and enter them into an array and then output
  // then into column D of sheet Spellings.
  let defnCounter: number = 0
  //looping through the hard words
  for (let wordItem of selectedWords) {
    const url = `https://api.dictionaryapi.dev/api/v2/entries/en/${selectedWords[defnCounter]}`;
    try {
      const Response = await fetch(url);
      const data = await Response.json() as DictionaryApiResponse[];
      const firstDefinition = data[0]?.meanings[0]?.definitions[0]?.definition || String;
      spellingsSheet.getRange('E' + (3 + defnCounter)).setValue(firstDefinition)
      defnCounter++
      //return firstDefinition     
    }
    catch (error) {
      console.log('There was an error with retrieving the definition for: ' + selectedWords[wordItem])
      spellingsSheet.getRange('E' + (3 + defnCounter)).setValue('Error')
      defnCounter++
      //return 'Error'
    }
  }
  //this string is not returned to anywhere
  return 'Routine Finished'
}
//interface following the structure of the returned JSON
interface DictionaryApiResponse {
  word: string;
  meanings: {
    partOfSpeech: string;
    definitions: {
      definition: string;
      synonyms: string[];
      antonyms: string[];
    }[];
  }[];
}