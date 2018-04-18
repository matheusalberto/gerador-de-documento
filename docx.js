var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

//Load the docx file as a binary
var content = fs
    .readFileSync(path.resolve(__dirname, 'template_doc_despachante.docx'), 'binary');

var zip = new JSZip(content);

var doc = new Docxtemplater();
doc.loadZip(zip);

var data = JSON.parse(fs.readFileSync('output.json', 'utf8')); 
doc.setData(data);

try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
}
catch (error) {
    var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
    }
    console.log(JSON.stringify({error: e}));
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
    throw error;
}

var buf = doc.getZip()
                .generate({type: 'nodebuffer'});

var today = new Date();
var dd = today.getDate();
var mm = today.getMonth()+1; //January is 0!
var yyyy = today.getFullYear();
var hour = today.getHours();
var minute = today.getMinutes();
var seconds = today.getSeconds();
var outputFile = 'docs/output' + hour + 'h__' + minute + 'min__' + seconds + 'sec__' + dd + '__' + mm + '__' + yyyy + '.docx';

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, outputFile), buf);
fs.writeFileSync(path.resolve(__dirname, 'filename.txt'), outputFile);