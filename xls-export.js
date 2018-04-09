/**
 * __  ___     _____                       _
 * \ \/ / |___| ____|_  ___ __   ___  _ __| |_
 *  \  /| / __|  _| \ \/ / '_ \ / _ \| '__| __|
 *  /  \| \__ \ |___ >  <| |_) | (_) | |  | |_
 * /_/\_\_|___/_____/_/\_\ .__/ \___/|_|   \__|
 *                       |_|
 * 6/12/2017
 * Daniel Blanco Parla
 * https://github.com/deblanco/xlsExport
 */

class XlsExport {
  // data: array of objects with the data for each row of the table
  // name: title for the worksheet
  constructor(data, title = 'Worksheet') {
    // input validation: new xlsExport([], String)
    if (!Array.isArray(data) || (typeof title !== 'string' || Object.prototype.toString.call(title) !== '[object String]')) {
      throw new Error('Invalid input types: new xlsExport(Array [], String)');
    }

    this._data = data;
    this._title = title;
  }

  set setData(data) {
    if (!Array.isArray(data)) throw new Error('Invalid input type: setData(Array [])');

    this._data = data;
  }

  get getData() {
    return this._data;
  }

  exportToXLS(fileName = 'export.xls') {
    if (typeof fileName !== 'string' || Object.prototype.toString.call(fileName) !== '[object String]') {
      throw new Error('Invalid input type: exportToCSV(String)');
    }

    const TEMPLATE_XLS = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"/>
        <head><!--[if gte mso 9]><xml>
        <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
        <![endif]--></head>
        <body>{table}</body></html>`;
    const MIME_XLS = 'application/vnd.ms-excel;base64,';

    const parameters = {
      title: this._title,
      table: this.objectToTable(),
    };
    const computeOutput = TEMPLATE_XLS.replace(/{(\w+)}/g, (x, y) => parameters[y]);

    const computedXLS = new Blob([computeOutput], {
      type: MIME_XLS,
    });
    const xlsLink = window.URL.createObjectURL(computedXLS);
    this.downloadFile(xlsLink, fileName);
  }

  exportToCSV(fileName = 'export.csv') {
    if (typeof fileName !== 'string' || Object.prototype.toString.call(fileName) !== '[object String]') {
      throw new Error('Invalid input type: exportToCSV(String)');
    }
    const computedCSV = new Blob([this.objectToSemicolons()], {
      type: 'text/csv;charset=utf-8',
    });
    const csvLink = window.URL.createObjectURL(computedCSV);
    this.downloadFile(csvLink, fileName);
  }

  downloadFile(output, fileName) {
    const link = document.createElement('a');
    document.body.appendChild(link);
    link.download = fileName;
    link.href = output;
    link.click();
  }

  toBase64(string) {
    return window.btoa(unescape(encodeURIComponent(string)));
  }

  objectToTable() {
    // extract keys from the first object, will be the title for each column
    const colsHead = `<tr>${Object.keys(this._data[0]).map(key => `<td>${key}</td>`).join('')}</tr>`;

    const colsData = this._data.map(obj => [`<tr>
                ${Object.keys(obj).map(col => `<td>${obj[col] ? obj[col] : ''}</td>`).join('')}
            </tr>`]) // 'null' values not showed
      .join('');

    return `<table>${colsHead}${colsData}</table>`.trim(); // remove spaces...
  }

  objectToSemicolons() {
    const colsHead = Object.keys(this._data[0]).map(key => [key]).join(';');
    const colsData = this._data.map(obj => [ // obj === row
      Object.keys(obj).map(col => [
        obj[col], // row[column]
      ]).join(';'), // join the row with ';'
    ]).join('\n'); // end of row

    return `${colsHead}\n${colsData}`;
  }
}

export default XlsExport; // comment this line to babelize
