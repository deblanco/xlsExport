# xls-export

Javascript library for exporting object arrays to Excel XLS and CSV.

## Usage

xlsExport is defined as a class, so has to be instantiated with **data** (objects array) and an optional **title**.

```javascript
var xls = new xlsExport([..., Object], String);
```

### Methods:
- *exportToXLS(String fileName)*: convert data and force download of a Excel XLS file.
- *exportToCSV(String fileName)*: convert data separate by semi-colons and force download of a CSV file.

*fileName parameter is **optional**, if its not defined the file will be named "export.xls".*

### Example
```javascript
var xls = new xlsExport([..., Object], String);
xls.exportToXLS('export2017.xls');
xls.exportToCSV('export2017.xls');
```

### Demo: https://jsfiddle.net/3xvb2g5w/

---

License: MIT