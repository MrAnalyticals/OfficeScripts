
## ✨ Automagic Insight: Fetch 15,000+ Excel Rows in 11 Seconds with Power Automate

---

### **Goal**  
Get over **15,000 rows** from an Excel file **without premium connectors**, **loops**, **pagination**, or **converting to a table** — and do it **blazingly fast** using **Office Scripts**.

---

### **Step-by-Step Flow Setup**

---

### 1. **Get File Properties**
- **Action**: `Get file properties` (SharePoint connector).
- **Purpose**: Identify and reference the Excel file stored in your SharePoint document library.
- **Time**: Very fast (~0.3 seconds).

---

### 2. **Run Office Script from SharePoint Library**
- **Action**: `Run script from SharePoint library`.
- **What the Script Does**:
  - **Reads a range** of Excel cells (no need to format as a table).
  - **Uses the first row as headers**.
  - **Outputs a clean JSON array** — **no additional parsing needed**.
- **Performance**: Fetches 15,000+ rows in **~11 seconds**!
- **How to Write the Script**:
  - You can ask ChatGPT:
    > *"Write an Office Script that reads an Excel range and returns an array of JSON objects using the first row as headers."*
  - This ensures your output is already perfectly structured for Power Automate (no messy post-processing).

---

### 3. **Select - Create Array (Optional)**
- **Action**: `Select` (Data Operation in Power Automate).
- **Purpose**: Further format the output if needed.
- **Tip**: Often, the JSON from the script is already clean enough that this step might be optional.

---

### **Result**
✅ 15,000+ rows handled smoothly  
✅ 30x faster than traditional methods  
✅ No premium licenses needed  
✅ Easy scaling for massive Excel files  

---

### **Bonus Tip**  
If you want ChatGPT to generate the Office Script **perfectly**, make sure you mention that the output should be an **array of JSON objects with headers** — not just an array of arrays.

---

### **Why This is Amazing**
- No 5000-row limit.
- No pagination.
- No heavy looping.
- Just **one lightweight action**.
- **Simpler, faster, and future-proof** for any automation involving big Excel files.

---

Would you also like me to give you a ready-to-use **Office Script code** sample to plug into your flow? 🚀  
(Just say the word!)
