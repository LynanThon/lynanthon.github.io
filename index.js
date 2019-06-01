const express = require('express');
const exphbs = require('express-handlebars');
const formidable = require('formidable');
const fs = require('fs');
const zip = require('express-zip');
const archiver = require('archiver');
const bodyParser = require('body-parser');
const path = require('path');
var session = require('express-session');
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

app.use(express.static(path.join(__dirname, 'public')));

// Handlebars
app.engine('handlebars', exphbs({ defaultLayout: 'landing' }));
app.set('view engine', 'handlebars');

// Body-parser
app.use(bodyParser.urlencoded({ extended: false }));

//session ID
app.use(session({ secret: 'contract', cookie: {maxAge: 60000}}))

app.listen(PORT, console.log(`Server is started on port ${PORT}`));

// Index Page
// app.get('/', (req, res) => res.render('layouts/main'));
app.get('/', (req, res) => {
    res.render('layouts/main');
    dir = `${__dirname}/outputs`;
    // fs.readdir(dir, (err, files) => {
    //     if (!(files === undefined || files.length === 0)) {
    //         files.forEach(file => {
    //             fs.unlinkSync(`${__dirname}/outputs/${file}`)
    //         });
    //     }
    // });
    // fs.readdir(`${__dirname}/contract-list/`, (err, file) => {
    //     if (!(file === undefined || file.length === 0)) {
    //         fs.unlinkSync(`${__dirname}/contract-list/${req.sessionID}.xlsx`);
    //     }
    // });
    // fs.readdir(`${__dirname}/zips`, (err, files) => {
    //     files.forEach(file => {
    //         fs.unlinkSync(`${__dirname}/zips/${file}`);
    //         // if(file.replace(".zip","") != req.sessionID)
    //         // {
    //         //     fs.unlinkSync(`${__dirname}/zips/${file}`);
    //         // }
    //     });
    // });
    // fs.readdir(`${__dirname}/templates/`, (err, file) => {
    //     if (!(file === undefined || file.length === 0)) {
    //         fs.unlinkSync(`${__dirname}/templates/${req.sessionID}.docx`);
    //     }
    // });
});

// Upload Files
app.post('/convert', (req, res) =>{
    let form = new formidable.IncomingForm();
    form.parse(req);

    // Statically set the .xlsx and .docx file name
    form.on('fileBegin', function (name, file){
        console.log(`${__dirname}/templates/${req.sessionID}.docx`)
        if (name === 'template_file') {
            file.path = `${__dirname}/templates/${req.sessionID}.docx`;
        } else if (name === 'data_file') {
            file.path = `${__dirname}/contract-list/${req.sessionID}.xlsx`;
        }
    });
    res.redirect('/local-files/convert');
});
 
// Handle POST request
app.post('/download', (req, res) => {
    dir = `${__dirname}/outputs`;

    var output = fs.createWriteStream(`${__dirname}/zips/${req.sessionID}.zip`);

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

    var output = fs.createWriteStream(`${__dirname}/zips/${req.sessionID}.zip`);

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
    // res.download(`${__dirname}/zips/result.zip`);
    fs.readdir(`${__dirname}/zips/`, (err, file) => {
        if (!(file === undefined || file.length === 0)) {
            res.download(`${__dirname}/zips/${req.sessionID}.zip`);
        }
        else {
            res.redirect('/')
        }
    });
});

// Local files route
app.use('/local-files', require('./routes/route'));

//this is default in case of unmatched routes
app.use((req, res) => {
    res.redirect('/')
});

