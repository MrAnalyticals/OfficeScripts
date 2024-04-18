
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell and worksheet.
    let selectedCell = workbook.getActiveCell();
    let selectedSheet = workbook.getWorksheet("Sheet1");

    // Set fill colour to yellow for the selected cell.
    // Define the pattern you want to apply
    const pattern = ExcelScript.FillPattern.solid
    //.SolidFillPattern.solid;

    // Define the fill color you want to use
    const fillColor = '#FFFF00'; // Yellow color

    // Get the range of cells you want to format
    const sheet = workbook.getWorksheet('Sheet1')
    const range = sheet.getRange('A1:B5')

    // Apply the pattern and fill color to the range
    range.getFormat().getFill().setPattern(pattern);
    range.getFormat().getFill().setColor(fillColor);
////
let currStatus = workbook.getActiveWorksheet().getRange("C1").getValue()
let conditionalFormatting: ExcelScript.ConditionalFormat;
  conditionalFormatting = selectedSheet.getRange("A1:B5").addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
 let positiveChange = conditionalFormatting.getCustom()
let condRule = "=($A$1=2)"
if(currStatus === "TODAY") {
    conditionalFormatting.getRange().getFormat().getFill().setPattern(ExcelScript.FillPattern.horizontal);
      positiveChange.getFormat().getFont().setColor("#9C0006");
      positiveChange.getFormat().getFill().setColor("#FFC7CE");
      positiveChange.getRule().setFormula("=($A$1=1)");
    }
 else {
  let selectedSheet = workbook.getActiveWorksheet();
  // Set fill color to FF0000 for range B8 on selectedSheet
  conditionalFormatting.getRange().getFormat().getFill().setPattern(ExcelScript.FillPattern.vertical);
       positiveChange.getFormat().getFill().setColor("FF0000");
    }
      positiveChange.getRule().setFormula(condRule);

}
