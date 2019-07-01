const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const XLSX = require('xlsx');
app.use(bodyParser.json());

app.use(bodyParser.urlencoded({ extended: true }));


app.use('/',(err, req, res) => {
  if (err) {
    console.log(err)
    // res.status(500).send(err.toString());
  } else {
    console.log('acat otototo')

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
});

// console.log(result);
app.listen(3001, () => {
  console.log('Listening 3001!');
});
