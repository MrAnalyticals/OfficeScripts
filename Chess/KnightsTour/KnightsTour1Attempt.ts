function main(workbook: ExcelScript.Workbook) {
  let SheetKnightsTour = workbook.getWorksheet('KnightsTour')
  let chartCollection = SheetKnightsTour.getCharts()
  let j: number
  let TheBlackKnight = SheetKnightsTour.getShape("BlackKnight");
  SheetKnightsTour.getRange("a1:H8").clear(ExcelScript.ClearApplyTo.contents);
  SheetKnightsTour.getRange("l3").clear(ExcelScript.ClearApplyTo.contents);
  SheetKnightsTour.getRange("k4").clear(ExcelScript.ClearApplyTo.contents);
  let startPos: string = SheetKnightsTour.getRange('L1').getValue().toString()
  let posNew: string = startPos
  for (j = 1; j < 65; j++) {
    console.log('loop j: ' + j)
    if (posNew == 'Err') { break }
    else {
      console.log('posNew: ' + posNew)
      moveKnight(posNew, workbook)
      SheetKnightsTour.getRange(posNew).setValue(j)
      SheetKnightsTour.getRange('L3').setValue(j)
      posNew = newPos(posNew, workbook)

    }
  }
  SheetKnightsTour.getCell(3, 10).setFormula('=IF(L3 = "", "", IF(L3 = 64, "Congratulations! You found a Full Knights Tour.", "Unfortunately, you did not find a full Knights Tour."))')
}

function newPos(pos: string, workbook: ExcelScript.Workbook): string {
  console.log('pos: ' + pos)
  let SheetKnightsTour = workbook.getWorksheet('KnightsTour')
  //startPos A:H 1:8
  let col: string = pos.substring(0, 1)
  let row: string = pos.substring(1, 2)
  let rownumber: number = parseInt(row)
  let colnumber: number
  //find all possible new positions. Build a newpositions array.
  //convert col to numbercol
  let newcol1: number
  if (convertColStrNo(col) != 0) { newcol1 = convertColStrNo(col) + 1 } else { newcol1 = 0 }
  let newrow1: number = rownumber + 1
  if (newrow1 > 8) { newrow1 = 0 }
  let newcol2: number
  if (convertColStrNo(col) != 0) { newcol2 = convertColStrNo(col) + 2 } else { newcol2 = 0 }
  let newrow2: number = rownumber + 2
  if (newrow2 > 8) { newrow2 = 0 }
  let newcolMinus1: number = newcol1 - 2
  if (newcolMinus1 < 1) { newcolMinus1 = 0 }
  let newcolMinus2: number = newcol1 - 3
  if (newcolMinus2 < 1) { newcolMinus2 = 0 }
  let newrowMinus1: number = rownumber - 1
  if (newrowMinus1 < 1) { newrowMinus1 = 0 }
  let newrowMinus2: number = rownumber - 2
  if (newrowMinus2 < 1) { newrowMinus2 = 0 }
  let newcol1Str: string = convertColNoStr(newcol1)
  let newcol2Str: string = convertColNoStr(newcol2)
  let newcolMinus1Str: string = convertColNoStr(newcolMinus1)
  let newcolMinus2Str: string = convertColNoStr(newcolMinus2)

  let new1: string = newcol1Str + newrow2  //Err+0
  let new2: string = newcol2Str + newrow1
  let new3: string = newcol1Str + newrowMinus2
  let new4: string = newcol2Str + newrowMinus1

  let new5: string = newcolMinus1Str + newrow2
  let new6: string = newcolMinus2Str + newrow1
  let new7: string = newcolMinus1Str + newrowMinus2
  let new8: string = newcolMinus2Str + newrowMinus1

  var newPosyArray = [new1, new2, new3, new4, new5, new6, new7, new8]
  console.log('Pre error removal newPosyArray: ' + newPosyArray)
  //loop through array finding the 'Err0' items.
  let x: number = 0
  //filter the array
  newPosyArray = newPosyArray.filter((obj) => {
    return obj.length < 4;
  });
  newPosyArray = newPosyArray.filter((obj) => {
    return obj.search('0') == -1;
  });

  console.log('Post error removal newPosyArray: ' + newPosyArray)
  let sizeArray: number = newPosyArray.length
  let randomItemPos: number
  randomItemPos = Math.floor(Math.random() * sizeArray)
  if (sizeArray == 0) { return 'Err' }
  else { //Randomly select one item from newpositions array. 
    while (SheetKnightsTour.getRange(newPosyArray[randomItemPos]).getValue().toString().length != 0) {
      newPosyArray = newPosyArray.filter((obj) => {
        return obj.search(newPosyArray[randomItemPos]) == -1
      })
      console.log('Post pos removal: newPosyArray: ' + newPosyArray);
      sizeArray = newPosyArray.length
      if (sizeArray == 0) { return 'Err' }
      randomItemPos = Math.floor(Math.random() * sizeArray)
    }
    return newPosyArray[randomItemPos]
  }
}

function convertColNoStr(col: number): string {
  switch (col) {
    case 1: { return 'A' }
    case 2: { return 'B' }
    case 3: { return 'C' }
    case 4: { return 'D' }
    case 5: { return 'E' }
    case 6: { return 'F' }
    case 7: { return 'G' }
    case 8: { return 'H' }
    default: { return 'Err' }
  }
}

function convertColStrNo(col: string): number {
  switch (col) {
    case 'A': { return 1 }
    case 'B': { return 2 }
    case 'C': { return 3 }
    case 'D': { return 4 }
    case 'E': { return 5 }
    case 'F': { return 6 }
    case 'G': { return 7 }
    case 'H': { return 8 }
    default: { return 0 }
  }
}

function moveKnight(movepos: string, workbook: ExcelScript.Workbook) {
  let SheetKnightsTour = workbook.getWorksheet('KnightsTour')
  let TheBlackKnight = SheetKnightsTour.getShape("BlackKnight");
  switch (movepos) {
    case 'A1' || 'a1': { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(0); break }
    case 'B1' || 'b1': { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(0); break }
    case ('C1' || 'c1'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(0); break }
    case ('D1' || 'd1'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(0); break }
    case ('E1' || 'e1'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(0); break }
    case ('F1' || 'f1'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(0); break }
    case ('G1' || 'g1'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(0); break }
    case ('H1' || 'h1'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(0); break }
    case ('A2' || 'a2'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(62.25); break }
    case ('B2' || 'b2'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(62.25); break }
    case ('C2' || 'c2'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(62.25); break }
    case ('D2' || 'd2'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(62.25); break }
    case ('E2' || 'e2'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(62.25); break }
    case ('F2' || 'f2'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(62.25); break }
    case ('G2' || 'g2'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(62.25); break }
    case ('H2' || 'h2'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(62.25); break }
    case ('A3' || 'a3'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(124.5); break }
    case ('B3' || 'b3'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(124.5); break }
    case ('C3' || 'c3'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(124.5); break }
    case ('D3' || 'd3'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(124.5); break }
    case ('E3' || 'e3'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(124.5); break }
    case ('F3' || 'f3'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(124.5); break }
    case ('G3' || 'g3'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(124.5); break }
    case ('H3' || 'h3'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(124.5); break }
    case ('A4' || 'a4'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(186.75); break }
    case ('B4' || 'b4'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(186.75); break }
    case ('C4' || 'c4'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(186.75); break }
    case ('D4' || 'd4'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(186.75); break }
    case ('E4' || 'e4'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(186.75); break }
    case ('F4' || 'f4'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(186.75); break }
    case ('G4' || 'g4'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(186.75); break }
    case ('H4' || 'h4'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(186.75); break }
    case ('A5' || 'a5'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(249); break }
    case ('B5' || 'b5'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(249); break }
    case ('C5' || 'c5'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(249); break }
    case ('D5' || 'd5'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(249); break }
    case ('E5' || 'e5'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(249); break }
    case ('F5' || 'f5'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(249); break }
    case ('G5' || 'g5'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(249); break }
    case ('H5' || 'h5'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(249); break }
    case ('A6' || 'a6'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(310.5); break }
    case ('B6' || 'b6'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(310.5); break }
    case ('C6' || 'c6'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(310.5); break }
    case ('D6' || 'd6'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(310.5); break }
    case ('E6' || 'e6'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(310.5); break }
    case ('F6' || 'f6'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(310.5); break }
    case ('G6' || 'g6'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(310.5); break }
    case ('H6' || 'h6'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(310.5); break }
    case ('A7' || 'a7'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(374.25); break }
    case ('B7' || 'b7'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(374.25); break }
    case ('C7' || 'c7'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(374.25); break }
    case ('D7' || 'd7'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(374.25); break }
    case ('E7' || 'e7'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(374.25); break }
    case ('F7' || 'f7'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(374.25); break }
    case ('G7' || 'g7'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(374.25); break }
    case ('H7' || 'h7'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(374.25); break }
    case ('A8' || 'a8'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(436.5); break }
    case ('B8' || 'b8'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(436.5); break }
    case ('C8' || 'c8'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(436.5); break }
    case ('D8' || 'd8'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(436.5); break }
    case ('E8' || 'e8'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(436.5); break }
    case ('F8' || 'f8'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(436.5); break }
    case ('G8' || 'g8'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(436.5); break }
    case ('H8' || 'h8'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(436.5); break }
  }
}
