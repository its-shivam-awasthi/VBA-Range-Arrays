# 📊 VBA Array Management Framework

A robust and extensible VBA framework for managing, indexing, filtering, and transforming arrays—designed for advanced Excel automation and data manipulation.

> ⚠️ Work in Progress.
> This project is actively under development. Some features may be incomplete, unstable, or subject to change. Contributions, suggestions, and bug reports are welcome!

---

## 🚀 Features

- Dynamic Array Indexing via ArrayIndexes

  - Row and column index tracking
  
  - Index-based filtering, slicing, and cell-level access

  - Centralised Array Management via Arrays
  
  - Dictionary-based reference system for multiple arrays

  - Array filtering, transposing, appending, and region extraction

  - Type-checking: numeric, text, date, boolean, jagged arrays

  - Excel Range Integration via RangeArray

  - Convert Excel ranges to arrays with header trimming

  - Export arrays back to Excel ranges

  -  Metadata tracking: workbook, worksheet, and range address

## 🛠️ Usage

### Load Excel Range into Array

<pre><code>```Dim rngManager As New RangeArray```</code></pre>
<pre><code>```rngManager.ArraysFromRanges "SalesData", "A1:D100", , , vbCurrentregion, True, vbRowHeade```</code></pre>

- Access and Manipulate Array

```vba
Dim arrObj As Arrays
Set arrObj = rngManager.Arrays("SalesData")

Debug.Print arrObj.Dimension
Debug.Print arrObj.CellValue(2, 3)
```

- Filter Array by Criteria

```vba
arrObj.Filter ">1000", 3, xlByRows, vbAutoDetect
```
- Export Back to Excel

```vba
Dim targetRange As Range
Set targetRange = Sheet1.Range("F1")
RngManager.ExportToRange "SalesData", targetRange
```

---

## 📚 Requirements

- Excel with VBA enabled

- No external dependencies

- Designed for 2D arrays and Excel ranges

---

##🔭 Roadmap & Future Plans

### Here’s what’s coming next:

- Header-Aware IndexingSupport for named headers and automatic mapping of header labels to column indexes.

- Dynamic Indexing EnhancementsSmarter row/column selection using conditions, labels, and expressions.

- Large Dataset OptimizationPerformance tuning for arrays with thousands of rows—minimizing memory and CPU overhead.

- Flexible Filtering & IFS LogicSupport for multi-condition filters, nested logic, and IFS-style output arrays.

- Improved Error Handling & DebuggingMore descriptive error messages and optional debug logging.


---

## 🧪 Testing

- Use the `ArrayIndexes` and `Arrays` classes independently for unit testing array logic.
- Integrate `RangeArray` for end-to-end Excel workflows.

---

## 📄 License

- MIT License.
-  Feel free to use, modify, and distribute.

- Let me know if you'd like to add badges, contributor guidelines, or sample files to this README. I can also help you write a changelog or GitHub Actions workflow if you're planning CI/CD.
