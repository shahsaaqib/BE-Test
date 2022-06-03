const XLSX = require('xlsx');

const workBook = XLSX.readFile('./data/beTest.xlsx');

const workSheet = workBook.Sheets[workBook.SheetNames[0]];

const data = XLSX.utils.sheet_to_json(workSheet);
const s1 = [];

for (let i = 0; i < data.length; i++) {
  const keys = Object.keys(data[i]).join('').split('|');
  const values = Object.values(data[i]).join('').split('|');
  const result = Object.assign.apply(
    {},
    keys.map((v, i) => ({ [v]: values[i] }))
  );

  const des = result.Description.split('/');
  let flag = '';
  if (des[0].includes('AEPS')) {
    flag = 'AEPS';
  } else if (des[0].includes('FEE CHG')) {
    flag = 'FEE CHG';
  }
  result.Flag = flag;
  result.DescrPtion = des[0];
  result[' '] = des[1];

  s1.push(result);
}
const newWorkBook = XLSX.utils.book_new();
const newWorkSheet = XLSX.utils.json_to_sheet(s1);
XLSX.utils.book_append_sheet(newWorkBook, newWorkSheet, 'Sheet1');
XLSX.writeFile(newWorkBook, 'BE Test Result.xlsx');
