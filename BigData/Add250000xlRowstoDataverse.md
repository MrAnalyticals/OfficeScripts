
# ✨ Processing 250,000 Excel Rows into Dataverse (with Clean Chunking)



---

## **Goal**
Efficiently process **250,000 Excel rows**, **15,000 rows at a time**, chunked into **100-row batches**, and **add to Dataverse** — using clean `chunk()` handling, high concurrency, and Office Scripts.

---

## **Full Step-by-Step Instructions**

---

### 1. **Get File Properties**
- **Action**: `Get file properties` (SharePoint connector).
- **Purpose**: Find the Excel file to work with.

---

### 2. **Initialize Variables**
Create the following variables:
| Variable | Type | Value |
|:---|:---|:---|
| `StartRow` | Integer | `2` (Assuming row 1 = headers) |
| `BatchSize` | Integer | `15000` |
| `TotalRows` | Integer | Empty (to store total Excel rows) |
| `CurrentRows` | Array | Empty (to temporarily hold 15,000 rows) |

---

### 3. **Get Total Row Count**
- **Run an Office Script** to determine total rows.
- **Office Script Code**:

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  let usedRange = sheet.getUsedRange();
  return usedRange.getRowCount();
}
```

- Output: Save to `TotalRows` variable.

---

### 4. **Until Loop**: `StartRow <= TotalRows`
Create an **Until** loop that keeps running **until** you have processed all rows.

---

### Inside the `Until` Loop:

---
#### 4.1 **Run Office Script** — Get 15,000 Rows Starting from `StartRow`
- **Office Script Code** (with parameters):

```typescript
function main(workbook: ExcelScript.Workbook, startRow: number, batchSize: number) {
  let sheet = workbook.getActiveWorksheet();
  let usedRange = sheet.getUsedRange();
  let values = usedRange.getValues();
  let headers = values[0];
  let output = [];

  for (let i = startRow - 1; i < Math.min(values.length, startRow - 1 + batchSize); i++) {
    let row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    output.push(row);
  }
  return output;
}
```

✅ Outputs a **JSON array** of **15,000 records**.

- Save output to `CurrentRows`.

---

#### 4.2 **Chunk 15,000 Rows into Batches of 100**
- Add a **Compose** action.
- Use this **expression**:

```plaintext
chunk(variables('CurrentRows'), 100)
```
_(or if directly from output without variable)_:

```plaintext
chunk(outputs('Run_Office_Script')?['body'], 100)
```

✅ Now you have an array where:
- Each item = array of 100 records

---

#### 4.3 **Apply to Each (100-Row Batches)**

- Create an **Apply to Each**:
  - **Input**: The **Compose output** (the 100-record arrays).

- **Concurrency**:
  - Turn **Concurrency Control = On**
  - Set **Degree of Parallelism = 50**

---
#### 4.4 **Inside Apply to Each Batch**

Create **another nested Apply to Each** inside:
- **Input**: Each item in the 100 records batch.

Inside this inner loop:
- **Action**: `Add a new row` (Dataverse connector).
- Map the fields accordingly.

✅ This handles **each individual record** cleanly from within its batch.

---

#### 4.5 **Increment StartRow**
- After the batches are done, update `StartRow`:
  - `StartRow = StartRow + BatchSize`
  - (`StartRow = StartRow + 15000`)

✅ Prepares the flow to fetch the **next 15,000 rows**.

---

### 5. **End of Loop**

✅ When `StartRow > TotalRows`, **flow completes**!

---

## 🧠 Visual Layout (Simple Summary)

```plaintext
Get file properties
↓
Initialize variables
↓
Run Office Script → Get Total Row Count
↓
UNTIL StartRow > TotalRows
    ├ Run Office Script (fetch 15,000 rows starting at StartRow)
    ├ Compose (chunk(15k rows, 100))
    ├ Apply to Each (each batch of 100 rows)
        ├ Nested Apply to Each (each record in 100 rows)
            ├ Add new row to Dataverse
    ├ Increment StartRow by 15,000
↓
Done
```

---

# 📊 Performance Recap

| | |
|:---|:---|
| Rows | 250,000 |
| Batches per 15k | 150 |
| Total API calls | 2,500 |
| Concurrency | 50 threads |
| Time per 15k rows | ~1–2 minutes |
| Total time | ~30–40 minutes |

---

# 🔥 Advantages
- **Very clean batch handling** (`chunk()`).
- **Concurrency enabled** = extremely fast uploading.
- **No premium connectors** required.
- **High reliability** even for **huge datasets**.

---

# 📢 **Notes**
- The **outer Apply to Each** processes 100-row batches.
- The **inner Apply to Each** processes **individual rows** inside the 100 batch.
- If you want to optimize even further, you can **bulk insert** instead of "row-by-row" later (advanced optimization).

---

# ✅ Short Version of Key Expressions:
- Chunking:
  ```plaintext
  chunk(outputs('Run_Office_Script')?['body'], 100)
  ```

- StartRow increment:
  ```plaintext
  add(variables('StartRow'), 15000)
  ```

---

# 🎯 Ready!

---

Would you also like me to:
- 🖼 Create a **diagram** of the full flow for easy building?  
- 🏗 Provide the **Power Automate actions JSON** you can copy-paste directly?  
- 📦 Help create a **ready-to-import** `.zip` file template?

Tell me what you want next! 🚀  
(I'm happy to help you go even faster!)
