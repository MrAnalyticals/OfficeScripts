function main(workbook: ExcelScript.Workbook) {
    let SheetKnightsTour = workbook.getWorksheet('KnightsTour')
    let j: number
  let TheBlackKnight = SheetKnightsTour.getShape("BlackKnight");
    // Move chart TheTheBlackKnight
    SheetKnightsTour.getRange("a1:H8").clear(ExcelScript.ClearApplyTo.contents)
let startpos:string = SheetKnightsTour.getRange('L1').getValue().toString()
  console.log(startpos)

  switch (startpos.toString()) {
    case 'A1' || 'a1': { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(0); break}
    case 'B1'||'b1':{console.log('Case B1 ran');
      TheBlackKnight.setLeft(61.5); 
      TheBlackKnight.setTop(0); break}
    case ('C1' || 'c1'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(0); break}
    case ('D1' || 'd1'):{ 
      TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(0); break}
    case ('E1' || 'e1'):{ 
      TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(0); break}
    case ('F1' || 'f1'):{ 
      TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(0); break}
    case ('G1' || 'g1'):{ 
      TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(0); break}
    case ('H1' || 'h1'):{ 
      TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(0); break}
    case ('A2' || 'a2'):{ 
      TheBlackKnight.setLeft(0); TheBlackKnight.setTop(62.25); break}
    case ('B2' || 'b2'):{ 
      TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(62.25); break}
    case ('C2' || 'c2'):{ 
      TheBlackKnight.setLeft(120); TheBlackKnight.setTop(62.25); break}
    case ('D2' || 'd2'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(62.25); break}
    case ('E2' || 'e2'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(62.25); break}
    case ('F2' || 'f2'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(62.25); break}
    case ('G2' || 'g2'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(62.25); break}
    case ('H2' || 'h2'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(62.25); break}
    case ('A3' || 'a3'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(124.5); break}
    case ('B3' || 'b3'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(124.5); break}
    case ('C3' || 'c3'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(124.5); break}
    case ('D3' || 'd3'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(124.5); break}
    case ('E3' || 'e3'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(124.5); break}
    case ('F3' || 'f3'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(124.5); break}
    case ('G3' || 'g3'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(124.5); break}
    case ('H3' || 'h3'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(124.5); break}
    case ('A4' || 'a4'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(186.75); break}
    case ('B4' || 'b4'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(186.75); break}
    case ('C4' || 'c4'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(186.75); break}
    case ('D4' || 'd4'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(186.75); break}
    case ('E4' || 'e4'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(186.75); break}
    case ('F4' || 'f4'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(186.75); break}
    case ('G4' || 'g4'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(186.75); break}
    case ('H4' || 'h4'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(186.75); break}
    case ('A5' || 'a5'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(249); break}
    case ('B5' || 'b5'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(249); break}
    case ('C5' || 'c5'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(249); break}
    case ('D5' || 'd5'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(249); break}
    case ('E5' || 'e5'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(249); break}
    case ('F5' || 'f5'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(249); break}
    case ('G5' || 'g5'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(249); break}
    case ('H5' || 'h5'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(249); break}
    case ('A6' || 'a6'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(310.5); break}
    case ('B6' || 'b6'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(310.5); break}
    case ('C6' || 'c6'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(310.5); break}
    case ('D6' || 'd6'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(310.5); break}
    case ('E6' || 'e6'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(310.5); break}
    case ('F6' || 'f6'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(310.5); break}
    case ('G6' || 'g6'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(310.5); break}
    case ('H6' || 'h6'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(310.5); break}
    case ('A7' || 'a7'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(374.25); break}
    case ('B7' || 'b7'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(374.25); break}
    case ('C7' || 'c7'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(374.25); break}
    case ('D7' || 'd7'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(374.25); break}
    case ('E7' || 'e7'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(374.25); break}
    case ('F7' || 'f7'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(374.25); break}
    case ('G7' || 'g7'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(374.25); break}
    case ('H7' || 'h7'): { TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(374.25); break}
    case ('A8' || 'a8'): { TheBlackKnight.setLeft(0); TheBlackKnight.setTop(436.5); break}
    case ('B8' || 'b8'): { TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(436.5); break}
    case ('C8' || 'c8'): { TheBlackKnight.setLeft(120); TheBlackKnight.setTop(436.5); break}
    case ('D8' || 'd8'): { TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(436.5); break}
    case ('E8' || 'e8'): { TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(436.5); break}
    case ('F8' || 'f8'): { TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(436.5); break}
    case ('G8' || 'g8'): { TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(436.5); break}
    case ('H8' || 'h8'):{TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(436.5);break}

    }}
