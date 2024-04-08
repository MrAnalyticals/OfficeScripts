function main(workbook: ExcelScript.Workbook, Name:string,	Age: number, Grade_Average_Enter_grades:number[], Gender:string,	Year:number, Extra_Curicular_Activity: string) {
  const StudentSheet = workbook.getWorksheet("StaffSheet");
  let StudentProfileTable = workbook.getTable('StudentProfileTable')
  let sum = 0;
  Grade_Average_Enter_grades.forEach((grade) => {
    sum += grade;
  });
  const average = sum / Grade_Average_Enter_grades.length;

  let inputRow = [[Name, Age, average, Gender, Year, Extra_Curicular_Activity]]
  StudentProfileTable.addRows(-1, inputRow)
  }
