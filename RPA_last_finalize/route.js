const express = require('express');
const router = express.Router();

// Excel
const readXlsxFile = require('read-excel-file/node');

// Docx-templates
const docxTem = require('docx-templates');

// Convert local file
router.get('/convert', (req, res) => {
    readXlsxFile(`${__dirname}/../contract-list/contract-list.xlsx`)
        .then((rows) => {
            headers = rows[0];
            rows.shift();
            rows.forEach(row => {
                data = {};
                for (i = 0; i < headers.length; i++) {
                    data[`${headers[i]}`] = (`${row[i]}`);
                }
                docxTem({
                    template: `${__dirname}/../templates/contract.docx`,
                    output: `${__dirname}/../outputs/${row[0]}.docx`,
                    data: data
                });
            });
            // res.send("Files are already prepared.")
            // res.redirect('/download');
            res.render('finished');
        })
        .catch(err => console.log(err));
});

module.exports = router;