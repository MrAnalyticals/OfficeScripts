function main(workbook: ExcelScript.Workbook) {
//throw new Error("Forced exit");
const Sheet1 = workbook.getWorksheet("Sheet1");
const InputSheet = workbook.getWorksheet("Input")
let WeatherTable = workbook.getTable("EastbourWeatherTable")
let WeatherTableHdFooter = workbook.getTable("EastbourWeatherTable").getRangeBetweenHeaderAndTotal()//.
 
//find last row, k, of Input table. 
let InvoiceIDInputTableVals = InputSheet.getRange("A3:a100000").getValues()
let k = 2
for (let cellval of InvoiceIDInputTableVals) {
            //console.log("cellval.toString(): "+cellval.toString())
            if (cellval.toString() == "") { break }	// let l = k
            k++
        }
let visibleRows = WeatherTableHdFooter.getVisibleView().getRows()
let rangeAddresses = visibleRows.map(row => row.getRange().getAddress());
            //loop through rangeAddresses array e.g.
for (let rowRangeStr of rangeAddresses) {
                //console.log("rowRangeStr.toString(): " + rowRangeStr.toString())
//limit the number of columns to copy if required by changing "g" to a letter before it.
let rowRange = Sheet1.getRange(rowRangeStr.replace("G", "g"))
                //console.log("rowRange.toString(): " + rowRange.toString())
InputSheet.getRange("A" + k).copyFrom(rowRange, ExcelScript.RangeCopyType.values, false, false);
        k++
}}
