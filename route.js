const express = require('express');
const router = express.Router();

const readXlsxFile = require('read-excel-file/node');
const docxTem = require('docx-templates');

const app = express();

router.get('/convert', (req, res) => {
    readXlsxFile('D:\\MyRPA\\contract-list\\ContractList.xlsx')
        .then((rows) => {
            headers = rows[0];
            rows.shift();
            rows.forEach(row => {
                data = {};
                for (i = 0; i < headers.length; i++) {
                    data[`${headers[i]}`] = (`${row[i]}`);
                }
                console.log(data)
                docxTem({
                    template: 'D:\\MyRPA\\templates\\contract.docx',
                    output: `D:\\MyRPA\\outputs\\${row[0]}.docx`,
                    data: data
                });
                console.log(`D:\\MyRPA\\outputs\\${row[0]}.docx`);
                console.log('Downloaded');
            });
            res.send("Files are already prepared.")
        })
        .catch(err => console.log(err));
});
module.exports = router;
