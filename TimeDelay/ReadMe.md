**Adding a time delay to your Office Scripts Code** 

![image](https://github.com/MrAnalyticals/OfficeScripts/assets/47678539/a1b5fb7d-7aec-4977-8a0a-94af9cf1e961)


Add this code to your routines where ever you need the code to wait. 

**YouTube video demo**: 


**Video Demo Code**
function main(workbook: ExcelScript.Workbook) {
let selectedSheet = workbook.getWorksheet('TimeDelay')
let timeDurationstart = new Date().getTime()
sleepy(5); // Wait for 5 seconds
let timeDuration = ((new Date().getTime()) - timeDurationstart)
selectedSheet.getCell(3, 2).setValue(timeDuration)
selectedSheet.getCell(3, 3).setValue("Miliseconds")}
function sleepy(seconds: number): void {
  const waitUntil: number = new Date().getTime() + seconds * 1000;
  while (new Date().getTime() < waitUntil) {}}

![image](https://github.com/MrAnalyticals/OfficeScripts/assets/47678539/c22be5f3-f6b8-47fd-8fcf-2523842cdc77)

  
