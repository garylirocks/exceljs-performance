const Excel = require('exceljs');
const data = require('./data.json');
const columns = require('./columns.json');
const { logMemory } = require('./memory.js');

console.log('==== Start ====');
logMemory();

const filename = `./output/excel-test.${new Date().toISOString()}.xlsx`;
const method = 'Normal';

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('My Sheet');

worksheet.columns = columns;

console.time(method);
data.forEach(row => {
  worksheet.addRow(row);
});

workbook.xlsx.writeFile(filename).then(function() {
  console.log('==== End ====');
  logMemory();

  console.timeEnd(method);
});
