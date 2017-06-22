/**
 * 21/06/2017
 * Daniel Blanco Parla
 * http://github.com/deblanco
 */

'use strict';

class xlsExport {

  // data: array of objects with the data for each row of the table
  // name: title for the worksheet
  constructor(data, title = 'Worksheet') {
    this._data = data;
    this._title = title;
  }

  set setData(data) {
    this._data = data;
  }

  get getData() {
    return this._data;
  }

  exportToXLS(fileName = 'export.xls') {
    const TEMPLATE_XLS = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><!--[if gte mso 9]><xml>
        <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
        <![endif]--><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head>
        <body>{table}</body></html>`;
    const MIME_XLS = 'data:application/vnd.ms-excel;base64,';

    const parameters = { title: this._title, table: this.objectToTable() };
    const computeOutput = TEMPLATE_XLS.replace(/{(\w+)}/g, (x, y) => parameters[y]);

    this.downloadFile(MIME_XLS + this.toBase64(computeOutput), fileName);
  }

  exportToCSV(fileName = 'export.csv') {
    const MIME_CSV = 'data:attachament/csv,';
    this.downloadFile(MIME_CSV + encodeURIComponent(this.objectToSemicolons()), fileName);
  }

  downloadFile(output, fileName) {
    const link = document.createElement('a');
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
                ${Object.keys(obj).map(col => `<td>${obj[col]}</td>`).join('')}
            </tr>`])
      .join('');

    return `
        <table>
            ${colsHead}
            ${colsData}
        </table>`.trim(); // remove spaces...
  }

  objectToSemicolons() {
    const colsHead = Object.keys(this._data[0]).map(key => [key]).join(';');
    const colsData = this._data.map(obj => [ // obj === row
                            Object.keys(obj).map(col => [
                                obj[col] // row[column]
                            ]).join(';') // join the row with ';'
                        ]).join('\n'); // end of row

    return `${colsHead}\n${colsData}`;
  }

}