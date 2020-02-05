const Excel = require('exceljs');
const data = require('./data.json');
const columns = require('./columns.json');
const { logMemory } = require('./memory.js');

const filename = `./output/excel-test.${new Date().toISOString()}.xlsx`;
const method = 'Streaming';

const options = {
  filename
  // useStyles: true,
  // useSharedStrings: true
};

const workbook = new Excel.stream.xlsx.WorkbookWriter(options);
const worksheet = workbook.addWorksheet('My Sheet');

worksheet.columns = columns;

console.time(method);
data.forEach(row => {
  worksheet.addRow(row).commit();
});

// Finished the workbook.
workbook.commit().then(function() {
  // the stream has been written
  logMemory();
  console.timeEnd(method);
});
