# xlsexport

Javascript library for exporting object arrays to Excel XLS and CSV.

## Installation

Clone or download the Github repo or via bower:

```bash
bower install xlsexport

or

npm install xlsexport
```

## Usage

xlsExport is defined as a class, so has to be instantiated with **data** (objects array) and an optional **title**.

```javascript
var xls = new XlsExport([..., Object], String);
```
Since Chromium(v61) supports ES6 Modules, XlsExport is available with 'import' syntax 😎. For older browsers I also include an ES5 version inside the package.

### Methods:
- *exportToXLS(String fileName)*: convert data and force download of a Excel XLS file.
- *exportToCSV(String fileName)*: convert data separate by semi-colons and force download of a CSV file.

*fileName parameter is **optional**, if it's not defined, the file will be named "export.xls".*

### Example
```javascript
import XlsExport from './xls-export.js';

var xls = new XlsExport([..., Object], String);
xls.exportToXLS('export2017.xls');
xls.exportToCSV('export2017.xls');
```

### Demo: https://jsfiddle.net/3xvb2g5w/

---

License: MIT