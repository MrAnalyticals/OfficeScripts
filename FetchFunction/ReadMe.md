


Scrape CSV web data with Excel Office Scripts without using Interfaces

In this repo I use the fetch function to retrieve csv data from a URL. The code did not need to use any interfaces to handle the returned data. 
Also included is a routine of VBA that uses the QueryTables.Add method to retrieve data from a URL.
The URL in the demo is :  https://raw.githubusercontent.com/treselle-systems/customer_churn_analysis/master/WA_Fn-UseC_-Telco-Customer-Churn.csv
It was sourced from webpage: Differences between the M Language and DAX in Power BI (sqlshack.com)
There is a 5million cells and 5MB size limit when using the fetch function in Excel Office Scripts. 
This URL tests this. It returns the “too big” notification: 
https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/FetchFunction/FR%20Sales.csv
This URL has 100,000 rows and is able to be fetched: 
https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/FetchFunction/MX%20Sales100000.csv
Ref:
Read or write to large ranges using the Excel JavaScript API - Office Add-ins | Microsoft Docs

See this reference for an example of using the JSON return type with the fetch function. 
Use external fetch calls in Office Scripts - Office Scripts | Microsoft Docs
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}

Here is the example VBA code script: 
 

Size Limits documentation: 
Resource limits and performance optimization for Office Add-ins - Office Add-ins | Microsoft Docs

 

The Office Script Example
FetchingCSVv3
