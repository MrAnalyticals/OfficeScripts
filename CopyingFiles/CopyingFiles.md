Welcome to another demo. I show you how to ,copy data from one workbook to another using power automate and excel office scripts code. There are no formatted tables in either of the workbooks. The business use case for this is copying data from a protected workbook, once a day into a shared workbook accessible by others in the organisation or team. The source workbook contains the data. the destination workbook is where the new data is copied to. The trigger is a schedule action and the next two actions are the copy from and paste to scripts. 
CopyText script:

function main(workbook: ExcelScript.Workbook): string[][] {
    const sheet = workbook.getWorksheet("Source");
    const usedRange = sheet.getUsedRange();
    let data: (string | number | boolean)[][] = usedRange.getValues();
    return data as string[][];
}

PasteToDestination script:
function main(workbook: ExcelScript.Workbook, data: string[][]): void {
    const sheet = workbook.getWorksheet("Destination"); // Change if needed
    const usedRange = sheet.getUsedRange();
    if (usedRange) usedRange.clear();

    const targetRange = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
    targetRange.setValues(data);
}
