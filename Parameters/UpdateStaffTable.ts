function main(workbook: ExcelScript.Workbook, Name: string, Dept: string, BusRole: string, Gender: string, ContactMethod: string) {
  const StaffSheet = workbook.getWorksheet("StaffSheet");
  let StaffTable = workbook.getTable('StaffListTable')
  let inputRow = [[Name, Dept, BusRole, Gender, ContactMethod]]
  StaffTable.addRows(-1, inputRow)
  }
