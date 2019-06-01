const express = require('express');
const formidabel = require('formidable');
const router = express.Router();
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const docxTem = require('docx-templates');
const archiver = require('archiver');
const app = express();
const download = require('download-file');
const download1 = function(url, dest, callback)
{
    request.get(url)
    .on('error', function(err) {console.log(err)} )
    .pipe(fs.createWriteStream(dest))
    .on('close', callback);
}

const PORT = process.env.PORT || 5600;

app.listen(PORT, console.log(`Server is started on port ${PORT}`));

app.get('/', function (req, res){
    res.sendFile(__dirname + '/index.html');
});

//Upload Files
app.post('/excel', function (req, res){
    var form = new formidabel.IncomingForm();
    form.parse(req);

    form.on('fileBegin', function (name, file){
        if(name == 'upload1'){
            file.path = __dirname + '/contract-list/ContractList.xlsx';
        }else{
            file.path = __dirname + '/templates/contract.docx';
        }
        console.log('File Uploaded');
    });

    // form.on('file', function (name, file){
    //     console.log('Upload :' + '/contract-list/' + file.name);
    // });
    res.sendFile(__dirname + '/index.html');
});

//Upload Template
// app.post('/template', function (req, res){
//     var form = new formidabel.IncomingForm();

//     form.parse(req);
    
//     console.log("#####################");

//     form.on('fileBegin', function (name, file){
//         console.log("====",name);
//         console.log("----",file);
//         console.log("****",__dirname);
//         file.path = __dirname + '/templates/' + file.name;
//     });

//     form.on('file', function (name , file){
//         console.log('Upload :' + '/templates/' + file.name);
//     });

//     res.sendFile(__dirname + '/index.html');
// });

//Convert File
app.post('/convert', function(req, res){
    console.log("Files converted");
    readXlsxFile(__dirname + '/contract-list/ContractList.xlsx')
        .then((rows) => {
            headers = rows[0];
            rows.shift();
            rows.forEach(row => {
                data = {};
                for (i =0; i < headers.length; i++) {
                    data[`${headers[i]}`] = (`${row[i]}`);
                }
                console.log(data)
                docxTem({
                    template: __dirname + '/templates/contract.docx',
                    output: `D:\\MyRPA\\outputs\\${row[0]}.docx`,
                    data: data
                });
            });
            res.send("Files are already prepared.")
        })
        .catch(err => console.log(err));
});

app.use('/local-files', require('./route'));

//Download File

// fs.readdir(__dirname+ '/outputs', (error, files) => {
//     let totalFiles = file.length;
//     console.log(totalFiles);
// })

// app.post('/download', function(req, res){
//     var file = __dirname + '/outputs';
//     fs.reddir(file, function (error, file) {
//         res.download(file);
//     })
//     let totalFiles = file.length;
//     console.log(totalFiles);
//     console.log("Downloaded")    
// })

app.post('/download', function(req, res){
    fs.reddir()
    console.log("Downloaded")
    const file = __dirname + '\\outputs\\Lynan.docx';
    res.download(file);
    // fs.readdir(file, (err, files) => {
    //     files.forEach(file => {
    //         res.download(file);
    //     })
    // })

    // for (var i=0; i<file.length; i++)
    // {
    //     res.download(file);
    // }

    // var dir = __dirname + '/outputs';

    // console.log("=======");
    // console.log(dir);

    // fs.readdir(dir, (err, files) => {
    //     console.log(files);
    //   //console.log(files.length);
    // });  
    // const testFolder = __dirname + '\\outputs\\';
    // fs.readdir(testFolder, (err, files) => {
    //     console.log("----"+testFolder);
    //     console.log("****"+files);
    //     files.forEach(file => {
    //         console.log("____"+file);
    //         const down = testFolder + file;
    //         console.log(down);
    //         res.download(down);
    //     });
    
        // for(num = 1; num <= 2; num++)
        // {
        //     const fileName = testFolder + files[num];
        //     const fileName1 = testFolder + files[num + 1];
        //     const fileName0 = testFolder + files[num - 1];
        //     if(num == 1)
        //     {
        //     console.log(fileName);
        //     res.download(fileName0);
        //     // res.download(fileName1);
        //     console.log('done' + fileName);
        //     console.log('done' + fileName1);
        //     }
        //     else
        //     {
        //         console.log('Done All');
        //     }
        // }
        // res.download(testFolder + '\\Lynan.docx\\');
        
    // })
});
