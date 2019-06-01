const express = require('express');
const exphbs = require('express-handlebars');
const formidable = require('formidable');
const fs = require('fs');
const zip = require('express-zip');
const archiver = require('archiver');
const bodyParser = require('body-parser');
const archive = archiver('zip', {
    gzip: true,
    zlib: { level: 9 } // Sets the compression level.
});

// Excel
const readXlsxFile = require('read-excel-file/node');

// Docx-templates
const docxTem = require('docx-templates');

const app = express();

const PORT = process.env.PORT || 5656;

// Handlebars
app.engine('handlebars', exphbs({ defaultLayout: 'landing' }));
app.set('view engine', 'handlebars');

// Body-parser
app.use(bodyParser.urlencoded({ extended: false }));

app.listen(PORT, console.log(`Server is started on port ${PORT}`));

// Index Page
app.get('/', (req, res) => res.render('layouts/main'));

// Upload Files
app.post('/convert', (req, res) =>{
    let form = new formidable.IncomingForm();
    form.parse(req);

    // Statically set the .xlsx and .docx file name
    form.on('fileBegin', function (name, file){
        if (name === 'template_file') {
            file.path = __dirname + '/templates/contract.docx';
        } else if (name === 'data_file') {
            file.path = __dirname + '/contract-list/contract-list.xlsx';
        }
    });
    res.redirect('/local-files/convert');

});

// Handle POST request
app.post('/download', (req, res) => {
    dir = `${__dirname}/outputs`;

    var output = fs.createWriteStream(`${__dirname}/zips/result.zip`);

    archive.pipe(output);

    fs.readdir(dir, (err, files) => {
        files.forEach(file => {
            let file1 = `${__dirname}/outputs/${file}`;
            archive.append(fs.createReadStream(file1), { name: `${file}` });
        });
        archive.finalize();
        res.redirect('/download-file');
    });

});

// Handle GET request
app.get('/download', (req, res) => {
    dir = `${__dirname}\\outputs`;

    var output = fs.createWriteStream(`${__dirname}/zips/result.zip`);

    archive.pipe(output);

    fs.readdir(dir, (err, files) => {
        files.forEach(file => {
            let file1 = `${__dirname}/outputs/${file}`;
            archive.append(fs.createReadStream(file1), { name: `${file}` });
        });
        archive.finalize();
        res.redirect('/download-file');
    });

});

app.get('/download-file', (req, res) =>{
    res.download(`${__dirname}/zips/result.zip`);
});

// Local files route
app.use('/local-files', require('./routes/route'));

