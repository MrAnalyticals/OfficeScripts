class CustomShape {
  private shape: ExcelScript.Shape; // The shape object
  private owner: string; // The owner of the shape

  constructor(shape: ExcelScript.Shape, owner: string) {
    this.shape = shape;
    this.owner = owner;
    this.setVisible(false); // Initially set the shape to be invisible
  }

  // Get the owner of the shape
  getOwner(): string {
    return this.owner;
  }

  // Set a new owner for the shape
  setOwner(owner: string) {
    this.owner = owner;
  }

  // Set the font size of the text within the shape
  setTextSize(size: number) {
    this.shape.getTextFrame().getTextRange().getFont().setSize(size);
  }

  // Set the text content of the shape
  setText(text: string) {
    this.shape.getTextFrame().getTextRange().setText(text);
  }

  // Set the fill color of the shape
  setFillColor(color: string) {
    this.shape.getFill().setSolidColor(color);
  }

  // Set the outline color of the shape
  setOutlineColor(color: string) {
    this.shape.getLineFormat().setColor(color);
  }

  // Set the position of the shape and log the position
  setPosition(left: number, top: number) {
    //console.log(`Setting position: left=${left}, top=${top}`);
    this.shape.setLeft(left);
    this.shape.setTop(top);
  }

  // Set the size of the shape
  setSize(width: number, height: number) {
    this.shape.setWidth(width);
    this.shape.setHeight(height);
  }

  // Set the visibility of the shape
  setVisible(visible: boolean) {
    this.shape.setVisible(visible);
  }

  // Find the rightmost shape among a list of shapes and return its rightmost position and top position
  static async findRightmostShape(shapes: ExcelScript.Shape[]): Promise<{ right: number, top: number }> {
    let rightmostShape: ExcelScript.Shape | null = null;
    let maxRight = -1;
    let topPosition: number = 0;
    for (const shape of shapes) {
      const shapeRight = shape.getLeft() + shape.getWidth();
      if (shapeRight > maxRight) {
        maxRight = shapeRight;
        rightmostShape = shape;
      }
    }
    if (rightmostShape) {
      topPosition = await rightmostShape.getTop();
    }
    //console.log(`Rightmost shape position: right=${maxRight}, top=${topPosition}`);
    return { right: maxRight, top: topPosition };
  }

  // Position the current shape next to the rightmost shape in the list and make it visible
  async positionNextToRightmostShape(shapes: ExcelScript.Shape[]) {
    const { right, top } = await CustomShape.findRightmostShape(shapes);
    this.setPosition(right + 20, top);
    this.setVisible(true); // Make the shape visible after repositioning
  }
}

// Main function to create and manipulate shapes in the workbook
async function main(workbook: ExcelScript.Workbook, setText: string, setshape: string, ownerName: string, alignShapes: boolean) {
  let sheet = workbook.getActiveWorksheet();
  sheet.getRange("A1").setValue(""); // Clear cell A1
  let shapetype: ExcelScript.GeometricShapeType;

  try {
    shapetype = getShapeType(setshape); // Determine the shape type
  } catch {
    console.log("Error! There is no shape called: " + setshape);
    sheet.getRange("A1").setValue("Error! There is no shape called: " + setshape);
    return "Error! There is no shape called: " + setshape;
  }

  let shaped = sheet.addGeometricShape(shapetype); // Add a new geometric shape to the worksheet
  let customShape = new CustomShape(shaped, ownerName); // Initialize a CustomShape instance

  // Set various properties of the shape
  customShape.setText(setText + "\n" + "New Property(Owner): " + ownerName);
  customShape.setFillColor("blue");
  customShape.setOutlineColor("red");
  customShape.setSize(100, 100);
  customShape.setTextSize(12);

  if (alignShapes) {
    let allShapes = sheet.getShapes(); // Get all shapes in the worksheet
    await customShape.positionNextToRightmostShape(allShapes); // Position the new shape next to the rightmost shape
  } else {
    // Uncomment and set random positions if not aligning shapes
    // let setLeft: number = 0//Math.floor(Math.random() * 200) + 1;
    // let setTop: number = 0//Math.floor(Math.random() * 200) + 1;
    // customShape.setPosition(setLeft, setTop);
    customShape.setVisible(true); // Make the shape visible if not aligning
  }
}

function getShapeType(shapeName:string) {
  switch (shapeName) {
    case "accentBorderCallout1":
      return ExcelScript.GeometricShapeType.accentBorderCallout1;
    case "accentBorderCallout2":
      return ExcelScript.GeometricShapeType.accentBorderCallout2;
    case "accentBorderCallout3":
      return ExcelScript.GeometricShapeType.accentBorderCallout3;
    case "accentCallout1":
      return ExcelScript.GeometricShapeType.accentCallout1;
    case "accentCallout2":
      return ExcelScript.GeometricShapeType.accentCallout2;
    case "accentCallout3":
      return ExcelScript.GeometricShapeType.accentCallout3;
    case "actionButtonBackPrevious":
      return ExcelScript.GeometricShapeType.actionButtonBackPrevious;
    case "actionButtonBeginning":
      return ExcelScript.GeometricShapeType.actionButtonBeginning;
    case "actionButtonBlank":
      return ExcelScript.GeometricShapeType.actionButtonBlank;
    case "actionButtonDocument":
      return ExcelScript.GeometricShapeType.actionButtonDocument;
    case "actionButtonEnd":
      return ExcelScript.GeometricShapeType.actionButtonEnd;
    case "actionButtonForwardNext":
      return ExcelScript.GeometricShapeType.actionButtonForwardNext;
    case "actionButtonHelp":
      return ExcelScript.GeometricShapeType.actionButtonHelp;
    case "actionButtonHome":
      return ExcelScript.GeometricShapeType.actionButtonHome;
    case "actionButtonInformation":
      return ExcelScript.GeometricShapeType.actionButtonInformation;
    case "actionButtonMovie":
      return ExcelScript.GeometricShapeType.actionButtonMovie;
    case "actionButtonReturn":
      return ExcelScript.GeometricShapeType.actionButtonReturn;
    case "actionButtonSound":
      return ExcelScript.GeometricShapeType.actionButtonSound;
    case "arc":
      return ExcelScript.GeometricShapeType.arc;
    case "bevel":
      return ExcelScript.GeometricShapeType.bevel;
    case "blockArc":
      return ExcelScript.GeometricShapeType.blockArc;
    case "borderCallout1":
      return ExcelScript.GeometricShapeType.borderCallout1;
    case "borderCallout2":
      return ExcelScript.GeometricShapeType.borderCallout2;
    case "borderCallout3":
      return ExcelScript.GeometricShapeType.borderCallout3;
    case "bracePair":
      return ExcelScript.GeometricShapeType.bracePair;
    case "bracketPair":
      return ExcelScript.GeometricShapeType.bracketPair;
    case "callout1":
      return ExcelScript.GeometricShapeType.callout1;
    case "callout2":
      return ExcelScript.GeometricShapeType.callout2;
    case "callout3":
      return ExcelScript.GeometricShapeType.callout3;
    case "can":
      return ExcelScript.GeometricShapeType.can;
    case "chartPlus":
      return ExcelScript.GeometricShapeType.chartPlus;
    case "chartStar":
      return ExcelScript.GeometricShapeType.chartStar;
    case "chartX":
      return ExcelScript.GeometricShapeType.chartX;
    case "chevron":
      return ExcelScript.GeometricShapeType.chevron;
    case "chord":
      return ExcelScript.GeometricShapeType.chord;
    case "circularArrow":
      return ExcelScript.GeometricShapeType.circularArrow;
    case "cloud":
      return ExcelScript.GeometricShapeType.cloud;
    case "cloudCallout":
      return ExcelScript.GeometricShapeType.cloudCallout;
    case "corner":
      return ExcelScript.GeometricShapeType.corner;
    case "cornerTabs":
      return ExcelScript.GeometricShapeType.cornerTabs;
    case "cross":
      return ExcelScript.GeometricShapeType.plus;
    case "cube":
      return ExcelScript.GeometricShapeType.cube;
    case "curvedDownArrow":
      return ExcelScript.GeometricShapeType.curvedDownArrow;
    case "curvedLeftArrow":
      return ExcelScript.GeometricShapeType.curvedLeftArrow;
    case "curvedRightArrow":
      return ExcelScript.GeometricShapeType.curvedRightArrow;
    case "curvedUpArrow":
      return ExcelScript.GeometricShapeType.curvedUpArrow;
    case "decagon":
      return ExcelScript.GeometricShapeType.decagon;
    case "diagonalStripe":
      return ExcelScript.GeometricShapeType.diagonalStripe;
    case "diamond":
      return ExcelScript.GeometricShapeType.diamond;
    case "dodecagon":
      return ExcelScript.GeometricShapeType.dodecagon;
    case "donut":
      return ExcelScript.GeometricShapeType.donut;
    case "doubleWave":
      return ExcelScript.GeometricShapeType.doubleWave;
    case "downArrow":
      return ExcelScript.GeometricShapeType.downArrow;
    case "downArrowCallout":
      return ExcelScript.GeometricShapeType.downArrowCallout;
    case "ellipse":
      return ExcelScript.GeometricShapeType.ellipse;
    case "ellipseRibbon":
      return ExcelScript.GeometricShapeType.ellipseRibbon;
    case "ellipseRibbon2":
      return ExcelScript.GeometricShapeType.ellipseRibbon2;
    case "flowChartAlternateProcess":
      return ExcelScript.GeometricShapeType.flowChartAlternateProcess;
    case "flowChartCollate":
      return ExcelScript.GeometricShapeType.flowChartCollate;
    case "flowChartConnector":
      return ExcelScript.GeometricShapeType.flowChartConnector;
    case "flowChartDecision":
      return ExcelScript.GeometricShapeType.flowChartDecision;
    case "flowChartDelay":
      return ExcelScript.GeometricShapeType.flowChartDelay;
    case "flowChartDisplay":
      return ExcelScript.GeometricShapeType.flowChartDisplay;
    case "flowChartDocument":
      return ExcelScript.GeometricShapeType.flowChartDocument;
    case "flowChartExtract":
      return ExcelScript.GeometricShapeType.flowChartExtract;
    case "flowChartInputOutput":
      return ExcelScript.GeometricShapeType.flowChartInputOutput;
    case "flowChartInternalStorage":
      return ExcelScript.GeometricShapeType.flowChartInternalStorage;
    case "flowChartMagneticDisk":
      return ExcelScript.GeometricShapeType.flowChartMagneticDisk;
    case "flowChartMagneticDrum":
      return ExcelScript.GeometricShapeType.flowChartMagneticDrum;
    case "flowChartMagneticTape":
      return ExcelScript.GeometricShapeType.flowChartMagneticTape;
    case "flowChartManualInput":
      return ExcelScript.GeometricShapeType.flowChartManualInput;
    case "flowChartManualOperation":
      return ExcelScript.GeometricShapeType.flowChartManualOperation;
    case "flowChartMerge":
      return ExcelScript.GeometricShapeType.flowChartMerge;
    case "flowChartMultidocument":
      return ExcelScript.GeometricShapeType.flowChartMultidocument;
    case "flowChartOfflineStorage":
      return ExcelScript.GeometricShapeType.flowChartOfflineStorage;
    case "flowChartOffpageConnector":
      return ExcelScript.GeometricShapeType.flowChartOffpageConnector;
    case "flowChartOnlineStorage":
      return ExcelScript.GeometricShapeType.flowChartOnlineStorage;
    case "flowChartOr":
      return ExcelScript.GeometricShapeType.flowChartOr;
    case "flowChartPredefinedProcess":
      return ExcelScript.GeometricShapeType.flowChartPredefinedProcess;
    case "flowChartPreparation":
      return ExcelScript.GeometricShapeType.flowChartPreparation;
    case "flowChartProcess":
      return ExcelScript.GeometricShapeType.flowChartProcess;
    case "flowChartPunchedCard":
      return ExcelScript.GeometricShapeType.flowChartPunchedCard;
    case "flowChartPunchedTape":
      return ExcelScript.GeometricShapeType.flowChartPunchedTape;
    case "flowChartSort":
      return ExcelScript.GeometricShapeType.flowChartSort;
    case "flowChartSummingJunction":
      return ExcelScript.GeometricShapeType.flowChartSummingJunction;
    case "flowChartTerminator":
      return ExcelScript.GeometricShapeType.flowChartTerminator;
    case "foldedCorner":
      return ExcelScript.GeometricShapeType.foldedCorner;
    case "frame":
      return ExcelScript.GeometricShapeType.frame;
    case "funnel":
      return ExcelScript.GeometricShapeType.funnel;
    case "gear6":
      return ExcelScript.GeometricShapeType.gear6;
    case "gear9":
      return ExcelScript.GeometricShapeType.gear9;
    case "halfFrame":
      return ExcelScript.GeometricShapeType.halfFrame;
    case "heart":
      return ExcelScript.GeometricShapeType.heart;
    case "heptagon":
      return ExcelScript.GeometricShapeType.heptagon;
    case "hexagon":
      return ExcelScript.GeometricShapeType.hexagon;
    case "homePlate":
      return ExcelScript.GeometricShapeType.homePlate;
    case "horizontalScroll":
      return ExcelScript.GeometricShapeType.horizontalScroll;
    case "irregularSeal1":
      return ExcelScript.GeometricShapeType.irregularSeal1;
    case "irregularSeal2":
      return ExcelScript.GeometricShapeType.irregularSeal2;
    case "leftArrow":
      return ExcelScript.GeometricShapeType.leftArrow;
    case "leftArrowCallout":
      return ExcelScript.GeometricShapeType.leftArrowCallout;
    case "leftBrace":
      return ExcelScript.GeometricShapeType.leftBrace;
    case "leftBracket":
      return ExcelScript.GeometricShapeType.leftBracket;
    case "lightningBolt":
      return ExcelScript.GeometricShapeType.lightningBolt;
    case "lineInv":
      return ExcelScript.GeometricShapeType.lineInverse;
//    case "line":
//      return ExcelScript.GeometricShapeType.line;
    case "mathDivide":
      return ExcelScript.GeometricShapeType.mathDivide;
    case "mathEqual":
      return ExcelScript.GeometricShapeType.mathEqual;
    case "mathMinus":
      return ExcelScript.GeometricShapeType.mathMinus;
    case "mathMultiply":
      return ExcelScript.GeometricShapeType.mathMultiply;
    case "mathNotEqual":
      return ExcelScript.GeometricShapeType.mathNotEqual;
    case "mathPlus":
      return ExcelScript.GeometricShapeType.mathPlus;
    case "moon":
      return ExcelScript.GeometricShapeType.moon;
    case "nonIsoscelesTrapezoid":
      return ExcelScript.GeometricShapeType.nonIsoscelesTrapezoid;
    case "noSmoking":
      return ExcelScript.GeometricShapeType.noSmoking;
    case "notchedRightArrow":
      return ExcelScript.GeometricShapeType.notchedRightArrow;
    case "octagon":
      return ExcelScript.GeometricShapeType.octagon;
    case "parallelogram":
      return ExcelScript.GeometricShapeType.parallelogram;
    case "pentagon":
      return ExcelScript.GeometricShapeType.pentagon;
    case "pie":
      return ExcelScript.GeometricShapeType.pie;
    case "pieWedge":
      return ExcelScript.GeometricShapeType.pieWedge;
    case "plaque":
      return ExcelScript.GeometricShapeType.plaque;
    case "plaqueTabs":
      return ExcelScript.GeometricShapeType.plaqueTabs;
    case "plus":
      return ExcelScript.GeometricShapeType.plus;
    case "quadArrow":
      return ExcelScript.GeometricShapeType.quadArrow;
    case "quadArrowCallout":
      return ExcelScript.GeometricShapeType.quadArrowCallout;
    case "rectangle":
      return ExcelScript.GeometricShapeType.rectangle;
    case "ribbon":
      return ExcelScript.GeometricShapeType.ribbon;
    case "ribbon2":
      return ExcelScript.GeometricShapeType.ribbon2;
    case "rightArrow":
      return ExcelScript.GeometricShapeType.rightArrow;
    case "rightArrowCallout":
      return ExcelScript.GeometricShapeType.rightArrowCallout;
    case "rightBrace":
      return ExcelScript.GeometricShapeType.rightBrace;
    case "rightBracket":
      return ExcelScript.GeometricShapeType.rightBracket;
    case "round1Rectangle":
      return ExcelScript.GeometricShapeType.round1Rectangle;
    case "round2DiagonalRectangle":
      return ExcelScript.GeometricShapeType.round2DiagonalRectangle;
    case "round2SameRectangle":
      return ExcelScript.GeometricShapeType.round2SameRectangle;
    case "roundRectangle":
      return ExcelScript.GeometricShapeType.roundRectangle;
    case "smileyFace":
      return ExcelScript.GeometricShapeType.smileyFace;
    case "snip1Rectangle":
      return ExcelScript.GeometricShapeType.snip1Rectangle;
    case "snip2DiagonalRectangle":
      return ExcelScript.GeometricShapeType.snip2DiagonalRectangle;
    case "snip2SameRectangle":
      return ExcelScript.GeometricShapeType.snip2SameRectangle;
    case "snipRoundRectangle":
      return ExcelScript.GeometricShapeType.snipRoundRectangle;
    case "squareTabs":
      return ExcelScript.GeometricShapeType.squareTabs;
    case "star10":
      return ExcelScript.GeometricShapeType.star10;
    case "star12":
      return ExcelScript.GeometricShapeType.star12;
    case "star16":
      return ExcelScript.GeometricShapeType.star16;
    case "star24":
      return ExcelScript.GeometricShapeType.star24;
    case "star32":
      return ExcelScript.GeometricShapeType.star32;
    case "star4":
      return ExcelScript.GeometricShapeType.star4;
    case "star5":
      return ExcelScript.GeometricShapeType.star5;
    case "star6":
      return ExcelScript.GeometricShapeType.star6;
    case "star7":
      return ExcelScript.GeometricShapeType.star7;
    case "star8":
      return ExcelScript.GeometricShapeType.star8;
    case "stripedRightArrow":
      return ExcelScript.GeometricShapeType.stripedRightArrow;
    case "sun":
      return ExcelScript.GeometricShapeType.sun;
    case "swooshArrow":
      return ExcelScript.GeometricShapeType.swooshArrow;
    case "teardrop":
      return ExcelScript.GeometricShapeType.teardrop;
    case "trapezoid":
      return ExcelScript.GeometricShapeType.trapezoid;
    case "triangle":
      return ExcelScript.GeometricShapeType.triangle;
    case "upArrow":
      return ExcelScript.GeometricShapeType.upArrow;
    case "upArrowCallout":
      return ExcelScript.GeometricShapeType.upArrowCallout;
    case "upDownArrow":
      return ExcelScript.GeometricShapeType.upDownArrow;
    case "upDownArrowCallout":
      return ExcelScript.GeometricShapeType.upDownArrowCallout;
    case "uturnArrow":
      return ExcelScript.GeometricShapeType.uturnArrow;
    case "verticalScroll":
      return ExcelScript.GeometricShapeType.verticalScroll;
    case "wave":
      return ExcelScript.GeometricShapeType.wave;
    case "wedgeEllipseCallout":
      return ExcelScript.GeometricShapeType.wedgeEllipseCallout;
    case "wedgeRectCallout":
      return ExcelScript.GeometricShapeType.wedgeRectCallout;
    case "wedgeRRectCallout":
      return ExcelScript.GeometricShapeType.wedgeRRectCallout;
    default:
      throw new Error("Invalid shape type");
  }
}