function main(workbook: ExcelScript.Workbook) {
    const answersSheet = workbook.getWorksheet("Answers");
    const answersTable = answersSheet.getTable("AnswersTable");
    const sheet = workbook.getWorksheet("Sheet1");
    const wordListRange = sheet.getRange("A1:A20");
    sheet.getRange("B1").getExtendedRange(ExcelScript.KeyboardDirection.down).getExtendedRange(ExcelScript.KeyboardDirection.right).clear(ExcelScript.ClearApplyTo.contents);
    answersTable.getRangeBetweenHeaderAndTotal().clear()
  	let rowcount = answersTable.getRowCount()
   	answersTable.deleteRowsAt(0, rowcount);    
    
    /*
    Instructions:
    1. Word List Range: Adjust the range A1:A20 to match the range where your words are listed.
    2. New Game Cell: Set the cell A36 to "New Game" to generate a new word search grid.
    */
    let wordList: Array<string>;
    wordList = wordListRange.getValues().toString().split(',').map(word => word.toUpperCase());
    const gridSizeCell = sheet.getRange("A35");
    const gridSize = gridSizeCell.getValue() as number;
    const gridRange = sheet.getRangeByIndexes(0, 1, gridSize, gridSize);
    const newGameCell = sheet.getRange("A36");

    if (newGameCell.getValue() === "New Game") {
        generateWordSearchGrid(wordList, gridRange, workbook);
        //
    }
}

function generateWordSearchGrid(wordList: string[], gridRange: ExcelScript.Range, workbook: ExcelScript.Workbook) {
    const gridSize = gridRange.getRowCount();
    const grid = Array.from({ length: gridSize }, () => Array(gridSize).fill(''));
    const directions = [
        [0, 1], [1, 0], [1, 1], [1, -1]
    ];
    const answers: Array<Array<string>> = [];

    wordList.forEach(word => {
        if (word) {
            let placed = false;
            while (!placed) {
                const row = Math.floor(Math.random() * gridSize);
                const col = Math.floor(Math.random() * gridSize);
                const direction = directions[Math.floor(Math.random() * directions.length)];

                if (canPlaceWord(word, grid, row, col, direction)) {
                    placeWord(workbook, word, grid, row, col, direction, answers, gridRange);
                    placed = true;
                }
            }
        }
    });

    fillEmptyCells(grid);
    gridRange.setValues(grid);
    //writeAnswersToSheet(answers, workbook);
}

function canPlaceWord(word: string, grid: string[][], row: number, col: number, direction: number[]): boolean {
    for (let i = 0; i < word.length; i++) {
        const newRow = row + i * direction[0];
        const newCol = col + i * direction[1];
        if (newRow < 0 || newRow >= grid.length || newCol < 0 || newCol >= grid[0].length || (grid[newRow][newCol] && grid[newRow][newCol] !== word[i])) {
            return false;
        }
    }
    return true;
}

function placeWord(workbook: ExcelScript.Workbook, word: string, grid: string[][], row: number, col: number, direction: number[], answers: Array<Array<string>>, gridRange: ExcelScript.Range): void {
    const answersSheet = workbook.getWorksheet("Answers");
    const answersTable = answersSheet.getTable("AnswersTable");
    const wordAddresses: Array<string> = [word];
    for (let i = 0; i < word.length; i++) {
        const newRow = row + i * direction[0];
        const newCol = col + i * direction[1];
        grid[newRow][newCol] = word[i];
        console.log("word[i]:" + word[i])
        wordAddresses.push(gridRange.getCell(newRow, newCol).getAddress().split('!')[1]);
    }
    answers.push(wordAddresses);
    console.log("wordAddresses:"+ wordAddresses)
    answersTable.addRow(-1, [wordAddresses[0].toString(), wordAddresses.slice(1).toString()]);
}

function fillEmptyCells(grid: string[][]): void {
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    grid.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
            if (!cell) {
                grid[rowIndex][colIndex] = alphabet[Math.floor(Math.random() * alphabet.length)];
            }
        });
    });
}
