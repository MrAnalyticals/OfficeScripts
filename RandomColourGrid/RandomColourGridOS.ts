function main(workbook: ExcelScript.Workbook) {
  const worksheetFlash = workbook.getWorksheet("RandomColours");
  const startRow = 2; // Start from row 2
  const endRow = 100; // End at row 100
  const column = 2; // Column B
  let range100 = worksheetFlash.getRange("B2:O23").getValues()
  for (let i = 0; i < 12; i++) {
    console.log('hello')
    range100.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      // Calculate the row number based on rowIndex and startRow
      const currentRow = rowIndex + startRow;
      // Set the color of the cell
      worksheetFlash.getCell(currentRow,colIndex).getFormat().getFill().setColor(getRandomColor());
    });
  });
  }
}
// Function to get a random color from the array
function getRandomColor() {
  var colors = [
    "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", "#FF4500", "#9400D3", "#8B008B", "#FF1493",
    "#800080", "#FF8C00", "#32CD32", "#4682B4", "#00FFFF",
    "#000080", "#8A2BE2", "#556B2F", "#DAA520", "#2F4F4F"
  ]
  const randomIndex = Math.floor(Math.random() * colors.length);
  return colors[randomIndex];
}