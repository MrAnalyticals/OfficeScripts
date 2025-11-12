This function is used in conjunction with Power Automate which inputs the JSON. 
Other input tools can be used of course. 

<img width="959" height="599" alt="image" src="https://github.com/user-attachments/assets/d1e0c59b-aa7c-4828-95a0-92543bf3c6d9" />


function main(workbook: ExcelScript.Workbook, jsonInput: string): string {
  // Define types
  type TextBlock = {
    type: string;
    text: string;
    weight: string;
    wrap: boolean;
    size: string;
  };

  type TableCell = {
    type: string;
    items: TextBlock[];
  };

  type TableRow = {
    type: string;
    cells: TableCell[];
  };

  // Parse JSON input
  const jsonData: TableRow[] = JSON.parse(jsonInput);

  // Extract headers from the first row
  const headerRow = jsonData[0];
  const headers: string[] = headerRow.cells.map(cell => {
    const textBlock = cell.items.find(item => item.type === "TextBlock");
    return textBlock?.text || "";
  });

  // Markdown header and divider
  const markdownHeader = `| ${headers.join(" | ")} |`;
  const markdownDivider = `| ${headers.map(() => "---").join(" | ")} |`;

  // Extract data rows
  const dataRows = jsonData.slice(1); // skip header
  const markdownData = dataRows.map(row => {
    const values = row.cells.map(cell => {
      const textBlock = cell.items.find(item => item.type === "TextBlock");
      return textBlock?.text || "";
    });
    return `| ${values.join(" | ")} |`;
  });

  // Combine all parts
  const markdownTable = [markdownHeader, markdownDivider, ...markdownData].join("\n");

  // Return to Power Automate
  return markdownTable;
}
