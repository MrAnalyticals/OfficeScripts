function main(workbook: ExcelScript.Workbook) {
  let SheetKnightsTour = workbook.getWorksheet('KnightsTour')
  let chartCollection = SheetKnightsTour.getCharts()
  let TheBlackKnight = SheetKnightsTour.getShape("BlackKnight")
  let j:number 
  SheetKnightsTour.getRange("a1:H8").clear(ExcelScript.ClearApplyTo.contents);
  SheetKnightsTour.getRange("l3").clear(ExcelScript.ClearApplyTo.contents);
  SheetKnightsTour.getRange("k4").clear(ExcelScript.ClearApplyTo.contents);
  //E1   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(0);
  SheetKnightsTour.getRange('e1').setValue(1);
  //G2   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(62.25);
  let m:number
  for(m=0;m<10;m++){console.log('Inserting some delay')}
  SheetKnightsTour.getRange('g2').setValue(2)
  //H4   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Inserting some delay') }
  SheetKnightsTour.getRange('h4').setValue(3)
  //F3   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f3').setValue(4)
  //E5   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e5').setValue(5)
  //G6   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g6').setValue(6)
  //H8   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h8').setValue(7)
  //F7   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f7').setValue(8)
  //D8   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d8').setValue(9)
  //B7   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b7').setValue(10)
  //A5   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a5').setValue(11)
  //C6   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c6').setValue(12)
  //D4   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d4').setValue(13)
  //B3   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b3').setValue(14)
  //A1   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a1').setValue(15)
  //C2   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c2').setValue(16)
  //B4   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b4').setValue(17)
  //A2   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a2').setValue(18)
  //C1   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c1').setValue(19)
  //D3   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d3').setValue(20)
  //C5   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c5').setValue(21)
  //A6   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a6').setValue(22)
  //B8   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b8').setValue(23)
  //D7   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d7').setValue(24)
  //F8   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f8').setValue(25)
  //H7   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h7').setValue(26)
  //G5   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g5').setValue(27)
  //e6
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e6').setValue(28)
  //F4   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f4').setValue(29)
  //H3   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h3').setValue(30)
  //G1   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g1').setValue(31)
  //E2   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e2').setValue(32)
  //C3   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c3').setValue(33)
  //D1   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d1').setValue(34)
  //B2   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b2').setValue(35)
  //A4   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a4').setValue(36)
  //B6   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b6').setValue(37)
  //A8   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a8').setValue(38)
  //C7   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c7').setValue(39)
  //D5   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d5').setValue(40)
  //F6   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f6').setValue(41)
  //E8   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e8').setValue(42)
  //G7   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g7').setValue(43)
  //H5   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h5').setValue(44)
  //G3   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g3').setValue(45)
  //H1   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h1').setValue(46)
  //F2   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f2').setValue(47)
  //G4   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g4').setValue(48)
  //H2   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h2').setValue(49)
  //F1   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f1').setValue(50)
  //E3   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e3').setValue(51)
  //F5   
  TheBlackKnight.setLeft(298.5); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('f5').setValue(52)
  //H6   
  TheBlackKnight.setLeft(415.5); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('h6').setValue(53)
  //G8   
  TheBlackKnight.setLeft(357.75); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('g8').setValue(54)
  //E7   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e7').setValue(55)
  //C8   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(436.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c8').setValue(56)
  //A7   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(374.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a7').setValue(57)
  //B5   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(249);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b5').setValue(58)
  //D6   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(310.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d6').setValue(59)
  //C4   
  TheBlackKnight.setLeft(120); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('c4').setValue(60)
  //A3   
  TheBlackKnight.setLeft(0); TheBlackKnight.setTop(124.5);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('a3').setValue(61)
  //B1   
  TheBlackKnight.setLeft(61.5); TheBlackKnight.setTop(0);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('b1').setValue(62)
  //D2   
  TheBlackKnight.setLeft(180.75); TheBlackKnight.setTop(62.25);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('d2').setValue(63)
  //E4   
  TheBlackKnight.setLeft(238.5); TheBlackKnight.setTop(186.75);
  for (m = 0; m < 10; m++) { console.log('Insetting some delay') }
  SheetKnightsTour.getRange('e4').setValue(64)
  }
