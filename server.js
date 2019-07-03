const express = require('express');
const app = express();
const XLSX = require('xlsx');

const sheet2arr = (sheet) => {
  let rowNum, colNum, range = XLSX.utils.decode_range(sheet['!ref']);
  const result = {};

  for(rowNum = 4; rowNum <= range.e.r; rowNum++){
    for(colNum=15; colNum<=range.e.c; colNum++){
        const nextCell = sheet[
          XLSX.utils.encode_cell({r: rowNum, c: colNum})
        ];
        if( typeof nextCell !== 'undefined' ){ 
          const header = getHeaderName(colNum, sheet);
          if (result[header]) {
            result[header].push(nextCell.w);
          } else {
            result[header] = [];
          }
        };
    }
  }
  return result;
};

const getHeaderName = (colNum, sheet) => {
  const headers = getAllHeaders(sheet);
  let name = '';
  loopHeader: for(let header of headers){
    const key = Object.keys(header).pop();
    if (key == colNum) {
      name = Object.values(header).pop();
      break loopHeader;
    }
  }
  return name;
}

const getAllHeaders = sheet => {
  var range = XLSX.utils.decode_range(sheet['!ref']);
  const headers = [];
  for(let colNum = 0; colNum <= range.e.c; colNum++) {
    const nextCell = sheet[
      XLSX.utils.encode_cell({r: 2, c: colNum})
    ];
    if (nextCell) {
      const { v } = nextCell;
      headers.push({[colNum]:v});
    }
  }
  return headers;
}

app.get('/rowWithHeaders', (req, res, next) => {

  const workbook = XLSX.readFile('./fileCsv.xlsx');
  const first_sheet_name = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[first_sheet_name];
  
  const data = sheet2arr(worksheet);
  res.json({ data });
  }
);

app.get('/collapsed', (req, res, next) => {

  const workbook = XLSX.readFile('./fileCsv.xlsx');
  const first_sheet_name = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[first_sheet_name];
  const range = XLSX.utils.decode_range(worksheet['!ref']);

  const result = {};
  for (let rowNum = 5; rowNum <= 137; rowNum++) {
    const desired_cell = worksheet[`AN${rowNum}`];
    if (typeof desired_cell !== 'undefined') {
      // const values = desired_cell.v.match(/[^\r\n|,|\s]+$/g);
      const values = desired_cell.v.match(/[^\r\n|,|\s]+/g);
      // const values = desired_cell.v.match(/^(\s*)(\S+(?:\s+([\S|\r|\n])+)*)(\s*)$/g);
      for (let val of values) {
        let valLower = val.toLowerCase();
        if (valLower[valLower.length-1] === ".") 
          valLower = valLower.slice(0,-1);
          
          for (let colNum = 15; colNum <= range.e.c; colNum++){ 
            let header = getHeaderName(colNum, worksheet);
            if (typeof header === 'string') {
              header = header.replace(/(\r\n)/g, " ");
            }
            const cellValues = worksheet[
              XLSX.utils.encode_cell({r: (rowNum-1), c: colNum})
            ];
            if (!result[valLower]) {
              if (typeof cellValues !== 'undefined') {
                let { v } = cellValues;
                  v = v.replace(/(\r\n)+/g, ", ");
                result[valLower] = [{[header]: [v]}]
              }
            } else {
              if (typeof cellValues !== 'undefined') {
                for(let node of result[valLower]) {
                  if (node[header]) {
                    let { v } = cellValues;
                    if (typeof v === 'string') {
                      v = v.replace(/(\r\n)+/g, ", ");
                    }
                    // console.log(typeof v)
                    node[header].push(v);
                  } else {
                    //someText.replace(/(\r\n|\n|\r)/gm, "");
                    // const value = cellValues.v.replace(/(\r\n)/g, " ");
                    let { v } = cellValues;
                    if (typeof v === 'string') {
                      v = v.replace(/(\r\n)+/g, ", ");
                    }
                    // v = v.replace(/(\r\n)/g, " ");
                    node[header] = [v];
                  }
                }
              }
            }
        }
      }
    }
  };
  delete result.clauses;
  delete result.clause;
  delete result.partial;
  delete result.annex;
  delete result.sections;
  delete result.deleted;
  delete result.all;
  delete result.in;
  delete result.section;
  delete result.plus;
  delete result[4];
  delete result[5];
  delete result[6];
  delete result[7];
  delete result[8];
  delete result[9];
  delete result[10];
  delete result['(partial)'];
  delete result['a'];
  delete result['c'];
  delete result['(b)'];
  delete result['(c)'];
  delete result['b)'];
  delete result['e'];
  delete result['(a)'];
  delete result['(2)'];
  for (let key of Object.keys(result)) {
    result[key] = result[key].pop();
    for(let secondKey of Object.keys(result[key])){
      result[key][secondKey] = result[key][secondKey].pop();
    }
  }
  res.json({ result });
  }
);

app.get('/', (req, res, next) => {

  const workbook = XLSX.readFile('./fileCsv.xlsx');
  const first_sheet_name = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[first_sheet_name];
  const result = {};
  for (let i = 5; i <= 137; i++) {
    const desired_cell = worksheet[`AN${i}`];
    if (typeof desired_cell !== 'undefined') {
      const values = desired_cell.v.match(/[^\r\n|,|\s]+/g);
      for (let val of values) {
        let valLower = val.toLowerCase();
        if (valLower[valLower.length-1] === ".") 
          valLower = valLower.slice(0,-1);
        if (!result[valLower]) {
          result[valLower] = [i]
        } else {
          result[valLower].push(i)
        }
      }
    }
  };
  res.json({ result });
}
);

app.listen(3001, () => {
  console.log('Listening 3001!');
});
