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

    this.MIME = 'data:application/vnd.ms-excel;base64,';
    this.TEMPLATE_XLS = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><!--[if gte mso 9]><xml>
        <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
        <![endif]--><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head>
        <body>{table}</body></html>`;
  }

  exportToXLS(fileName = 'export') {
    const parameters = { title: this._title, table: this.objectToTable() };
    const computeOutput = this.TEMPLATE_XLS.replace(/{(\w+)}/g, (x, y) => parameters[y]);
    const link = document.createElement('a');
    link.download = `${fileName}.xls`;
    link.href = this.MIME + this.toBase64(computeOutput);
    link.click();
  }

  exportToCSV() {

  }

  toBase64(string) {
    return window.btoa(unescape(encodeURIComponent(string)));
  }

  objectToTable() {
    // extract keys from the first object, will be the title for each column
    const colsHead = `<tr>${Object.keys(this._data[0]).map(key => `<td>${key}</td>`).join('')}</tr>`;

    const colsData = this._data.map(obj => [`<tr>
                ${Object.keys(obj).map(col => `<td>${obj[col]}</td>`)}
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

    return `${colsHead}\n
            ${colsData}`;
  }

  set setData(data) {
    this._data = data;
  }

  get getData() {
    return this._data;
  }

}