/**
 * 21/06/2017
 * Daniel Blanco Parla
 * http://github.com/deblanco
 */

'use strict';

class xlsExport {

  // data: array of objects with the data for each row of the table
  // name: title for the worksheet
  constructor(data, title) {
    this.data = data;
    this.title = name;

    this.MIME = 'data:application/vnd.ms-excel;base64,';
    this.TEMPLATE_XLS = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><!--[if gte mso 9]><xml>
        <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
        <![endif]--><meta charset="utf-8"></head>
        <body>{table}</body></html>`;
  }

  exportXLS() {

  }

  exportCSV() {

  }

  toBase64(string) {
    return window.btoa(unescape(encodeURIComponent(string)));
  }

  objectToTable() {
    // extract keys from the first object, will be the title for each column
    const colsHead = `<tr>${Object.keys(this.data[0]).map(key => `<td>${key}</td>`).join('')}</tr>`;

    const colsData = this.data.map(obj => [`<tr>
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
    const colsHead = Object.keys(this.data[0]).map(key => [key]).join(';');
    const colsData = this.data.map(obj => [ // obj === row
                            Object.keys(obj).map(col => [
                                obj[col] // row[column]
                            ]).join(';') // join the row with ';'
                        ]).join('\n'); // end of row

    return `${colsHead}\n
            ${colsData}`;
  }

  // set data(data) {
  //   this.data = data;
  // }

  // get data() {
  //   return this.data;
  // }

}