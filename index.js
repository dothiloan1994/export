const express = require("express"); // goi module express de su dung
const app = express(); //xay nha-tao dich vu host
const port = process.env.PORT || 3000;//su dung port cua file env, neu khong co file nay thi su dung port 5000
/*const multer = require("multer");*/
const Excel = require('exceljs');
const upload = require("express-fileupload");
const bodyParser = require('body-parser')
const fs = require('fs');
const parse = require('csv-parser');
/*const { pipeline } = require('stream/promises');*/

/*const csvtojson = require("csvtojson");*/

/*var storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads')
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + '-' + Date.now());
    }
})
var upload = multer({ storage });*/

app.listen(port,function (){
    console.log("Server is running...");
});

app.use(express.static("public"));
app.use(upload());
app.use( bodyParser.json() );       // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
    extended: true
}));

app.set("view engine","ejs");

app.get("/",function (req,res){
    res.render("home");
});
/*
app.post("/upload-done", (req,res) => {
    var fs = require('fs');
    var parse = require('csv-parser');
    var csvData=[];
    /!*console.log(uploadfile.path);*!/
    /!*console.log(req.files.uploadfile.data);*!/
 var csvData = req.files.uploadfile.data.toString('utf8');
/!*    var csvtojson = csvtojson().fromString(csvData);*!/
 var data3 = csvData.replace(/\r\n/g, "\r").replace(/\n/g, "\r").split(/\r/);
 console.log(data3);
/!*var files = req.files;*!/
    /!*console.log();*!/
    var pattern = req.query.pattern;
    var err = false;
    var wb = new Excel.Workbook();

    wb.xlsx.readFile('uploads/template.xlsx').then(function() {
        fs.createReadStream("uploads/temple.csv")
        .pipe(parse({delimiter: ':'}))
        .on('data', function(csvrow) {
            var datas = JSON.stringify(csvrow).split(',');
            for (let data of datas) {
                var values = data.split(':');
                csvData.push(values[1]);
                var items = values[1].split(';');
                if (items.length == 5) {
                    var tmpws = wb.getWorksheet('Template');
                    var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                    ws.model = tmpws.model;
                    ws.name = items[0].split('"')[1] + '_' + items[1];
                    ws.getCell('3', '6').value = items[1];
                    ws.getCell('4', '6').value = items[2];
                    ws.getCell('5', '6').value = items[3];
                    ws.mergeCells('A8', 'S13');
                    var ketqua = items[4].split('"')[0];
                    if(ketqua=='1') {
                        ws.getCell('8', '1').value = 'Benh ung thu vu giai doan 1. Can dieu tri...';
                    } else if(ketqua=='2') {
                        ws.getCell('8', '1').value = 'Benh ung thu vu giai doan 2. Can dieu tri...';
                    }
                } else{
                    err = true;
                }
            }
        })
        .on('end',function() {
            //do something with csvData
            if(err==false) {
                res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                res.setHeader('Content-Disposition', 'attachment; filename=' + 'export.xlsx');

                return wb.xlsx.write(res)
                    .then(function () {
                        res.status(200).end();
                    });
            } else{
                res.send('File upload bi loi');
            }
        });
    })
});*/

app.post("/upload-done", async function(req,res) {
    var csvData = req.files.uploadfile.data.toString('utf8');
    var dataRaw = csvData.replace(/\r\n/g, "\r").replace(/\n/g, "\r").split(/\r/);
    var timeInMss = new Date();
    var pattern = req.body.pattern;
    var err = false;
    var wb = new Excel.Workbook();
    try {
        if (pattern == "1") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                    var datas = JSON.stringify(dataRaw).split(',');
                    for (let i = 1; i < datas.length; i++) {
                        var items = datas[i].split(';');
                        if (items.length == 11) {
                            var tmpws = wb.getWorksheet('Template');
                            var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                            ws.model = tmpws.model;
                            ws.name = items[0].split('"')[1] + '_' + items[1];
                            /* ws.getCell('7', '8').value = items[1];*/

                            ws.mergeCells('A1', 'K1');
                            ws.mergeCells('A2', 'K2');
                            ws.mergeCells('C4', 'W4');
                            ws.mergeCells('C5', 'W5');
                            ws.mergeCells('C7', 'F7');
                            ws.mergeCells('C8', 'F8');
                            ws.mergeCells('C9', 'F9');
                            ws.mergeCells('C10', 'F10');
                            ws.mergeCells('C11', 'F11');
                            ws.mergeCells('07', 'R7');
                            ws.mergeCells('010', 'R10');
                            ws.mergeCells('E13', 'M13');
                            ws.mergeCells('N13', 'U13');
                            ws.mergeCells('E14', 'M14');
                            ws.mergeCells('N14', 'U14');
                            ws.mergeCells('D17', 'U21');
                            ws.mergeCells('G22', 'U23');
                            ws.mergeCells('G24', 'U24');
                            ws.mergeCells('M27', 'V27');
                            ws.mergeCells('M32', 'V32');
                            ws.mergeCells('M33', 'V33');

                            /*ws.getCell('E13').style.border = {
                                top: { style: 'thick' },
                                left: { style: 'thick' },
                                bottom: { style: 'thick' },
                                right: { style: 'thick' }
                            };*/
                            ws.getCell('S7').value = items[2];
                            ws.getCell('H7').value = items[1];
                            ws.getCell('V7').value = items[6];
                            ws.getCell('H8').value = items[3];
                            ws.getCell('H9').value = items[9];
                            ws.getCell('S10').value = items[5];
                            ws.getCell('H11').value = items[4];
                            ws.getCell('H11').value = items[4];
                            ws.getCell('E14').value = items[4];
                            ws.getCell('N14').value = items[10].split('"')[0];
                            ws.getCell('H10').value = items[8];
                            var timeInMss = new Date();
                            var times = JSON.stringify(timeInMss).split('-');
                            ws.getCell('V26').value = times[0].split('"')[1];
                            ws.getCell('T26').value = times[1];
                            ws.getCell('Q26').value = times[2].split('T')[0];
                            for(let i=0;i<csvData.length;i++){
                                var ketluan = JSON.stringify(csvData[i]).split(",");
                                if(ketluan[1]=='"'+items[7]+'"'){
                                   if(ketluan[2].length>1) {
                                       ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                   }
                                    if(ketluan[3].length>1) {
                                        ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                    }
                                    if(ketluan[4].length>5) {
                                        ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                    }
                                }
                            }
                        } else{
                            err=true;
                        }
                    }

                    if(err==false){
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                        res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                        return wb.xlsx.write(res)
                            .then(function () {
                                res.status(200).end();
                            });
                    }else {
                        err = true;
                        res.send('File upload bi loi');
                    }
                })
            })
        } else if (pattern == "2") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template_img.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                var datas = JSON.stringify(dataRaw).split(',');
                for (let i = 1; i < datas.length; i++) {
                    var items = datas[i].split(';');
                    if (items.length == 11) {
                        var tmpws = wb.getWorksheet('Template');
                        var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                        ws.model = tmpws.model;
                        ws.name = items[0].split('"')[1] + '_' + items[1];
                        /* ws.getCell('7', '8').value = items[1];*/

                        ws.mergeCells('A1', 'K1');
                        ws.mergeCells('A2', 'K2');
                        ws.mergeCells('C4', 'W4');
                        ws.mergeCells('C5', 'W5');
                        ws.mergeCells('C7', 'F7');
                        ws.mergeCells('C8', 'F8');
                        ws.mergeCells('C9', 'F9');
                        ws.mergeCells('C10', 'F10');
                        ws.mergeCells('C11', 'F11');
                        ws.mergeCells('07', 'R7');
                        ws.mergeCells('010', 'R10');
                        ws.mergeCells('E13', 'M13');
                        ws.mergeCells('N13', 'U13');
                        ws.mergeCells('E14', 'M14');
                        ws.mergeCells('N14', 'U14');
                        ws.mergeCells('D17', 'M21');
                        ws.mergeCells('N17', 'U21');
                        ws.mergeCells('G22', 'U23');
                        ws.mergeCells('G24', 'U24');
                        ws.mergeCells('M27', 'V27');
                        ws.mergeCells('M32', 'V32');
                        ws.mergeCells('M33', 'V33');

                        /*ws.getCell('E13').style.border = {
                            top: { style: 'thick' },
                            left: { style: 'thick' },
                            bottom: { style: 'thick' },
                            right: { style: 'thick' }
                        };*/
                        ws.getCell('S7').value = items[2];
                        ws.getCell('H7').value = items[1];
                        ws.getCell('V7').value = items[6];
                        ws.getCell('H8').value = items[3];
                        ws.getCell('H9').value = items[9];
                        ws.getCell('S10').value = items[5];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('E14').value = items[4];
                        ws.getCell('N14').value = items[10].split('"')[0];
                        ws.getCell('H10').value = items[8];
                        var timeInMss = new Date();
                        var times = JSON.stringify(timeInMss).split('-');
                        ws.getCell('V26').value = times[0].split('"')[1];
                        ws.getCell('T26').value = times[1];
                        ws.getCell('Q26').value = times[2].split('T')[0];
                        for(let i=0;i<csvData.length;i++){
                            var ketluan = JSON.stringify(csvData[i]).split(",");
                            if(ketluan[1]=='"'+items[7]+'"'){
                                if(ketluan[2].length>1) {
                                    ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                }
                                if(ketluan[3].length>1) {
                                    ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                }
                                if(ketluan[4].length>5) {
                                    ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                }
                            }
                        }
                    } else{
                        err=true;
                    }
                }

                    if(err==false){
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                        res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                        return wb.xlsx.write(res)
                            .then(function () {
                                res.status(200).end();
                            });
                    }else {
                        err = true;
                        res.send('File upload bi loi');
                    }
                })
            })
        } else if (pattern == "3") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template_ubieuhn.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                var datas = JSON.stringify(dataRaw).split(',');
                for (let i = 1; i < datas.length; i++) {
                    var items = datas[i].split(';');
                    if (items.length == 11) {
                        var tmpws = wb.getWorksheet('Template');
                        var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                        ws.model = tmpws.model;
                        ws.name = items[0].split('"')[1] + '_' + items[1];
                        /* ws.getCell('7', '8').value = items[1];*/

                        ws.mergeCells('A1', 'H1');
                        ws.mergeCells('A2', 'H2');
                        ws.mergeCells('I1', 'W1');
                        ws.mergeCells('I2', 'W2');
                        ws.mergeCells('C4', 'W4');
                        ws.mergeCells('C5', 'W5');
                        ws.mergeCells('C7', 'F7');
                        ws.mergeCells('C8', 'F8');
                        ws.mergeCells('C9', 'F9');
                        ws.mergeCells('C10', 'F10');
                        ws.mergeCells('C11', 'F11');
                        ws.mergeCells('07', 'R7');
                        ws.mergeCells('010', 'R10');
                        ws.mergeCells('E13', 'M13');
                        ws.mergeCells('N13', 'U13');
                        ws.mergeCells('E14', 'M14');
                        ws.mergeCells('N14', 'U14');
                        ws.mergeCells('D17', 'U21');
                        ws.mergeCells('G22', 'U23');
                        ws.mergeCells('G24', 'U24');
                        ws.mergeCells('M27', 'V27');
                        ws.mergeCells('M32', 'V32');
                        ws.mergeCells('M33', 'V33');

                        /*ws.getCell('E13').style.border = {
                            top: { style: 'thick' },
                            left: { style: 'thick' },
                            bottom: { style: 'thick' },
                            right: { style: 'thick' }
                        };*/
                        ws.getCell('S7').value = items[2];
                        ws.getCell('H7').value = items[1];
                        ws.getCell('V7').value = items[6];
                        ws.getCell('H8').value = items[3];
                        ws.getCell('H9').value = items[9];
                        ws.getCell('S10').value = items[5];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('E14').value = items[4];
                        ws.getCell('N14').value = items[10].split('"')[0];
                        ws.getCell('H10').value = items[8];
                        var times = JSON.stringify(timeInMss).split('-');
                        ws.getCell('V26').value = times[0].split('"')[1];
                        ws.getCell('T26').value = times[1];
                        ws.getCell('Q26').value = times[2].split('T')[0];

                        for(let i=0;i<csvData.length;i++){
                            var ketluan = JSON.stringify(csvData[i]).split(",");
                            if(ketluan[1]=='"'+items[7]+'"'){
                                if(ketluan[2].length>1) {
                                    ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                }
                                if(ketluan[3].length>1) {
                                    ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                }
                                if(ketluan[4].length>5) {
                                    ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                }
                            }
                        }
                    } else{
                        err=true;
                    }
                }

                if(err==false){
                    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                    res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua_ubieuhn' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                    return wb.xlsx.write(res)
                        .then(function () {
                            res.status(200).end();
                        });
                }else {
                    err = true;
                    res.send('File upload bi loi');
                }
            })
        })
        } else if (pattern == "4") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template_dainghia.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                var datas = JSON.stringify(dataRaw).split(',');
                for (let i = 1; i < datas.length; i++) {
                    var items = datas[i].split(';');
                    if (items.length == 11) {
                        var tmpws = wb.getWorksheet('Template');
                        var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                        ws.model = tmpws.model;
                        ws.name = items[0].split('"')[1] + '_' + items[1];
                        /* ws.getCell('7', '8').value = items[1];*/

                        /*ws.mergeCells('A1', 'H1');
                        ws.mergeCells('A2', 'H2');*/
                        ws.mergeCells('I1', 'W1');
                        ws.mergeCells('I2', 'W2');
                        ws.mergeCells('C4', 'W4');
                        ws.mergeCells('C5', 'W5');
                        ws.mergeCells('C7', 'F7');
                        ws.mergeCells('C8', 'F8');
                        ws.mergeCells('C9', 'F9');
                        ws.mergeCells('C10', 'F10');
                        ws.mergeCells('C11', 'F11');
                        ws.mergeCells('07', 'R7');
                        ws.mergeCells('010', 'R10');
                        ws.mergeCells('E13', 'M13');
                        ws.mergeCells('N13', 'U13');
                        ws.mergeCells('E14', 'M14');
                        ws.mergeCells('N14', 'U14');
                        ws.mergeCells('D17', 'U21');
                        ws.mergeCells('G22', 'U23');
                        ws.mergeCells('G24', 'U24');
                        ws.mergeCells('M27', 'V27');
                        ws.mergeCells('M32', 'V32');
                        ws.mergeCells('M33', 'V33');

                        /*ws.getCell('E13').style.border = {
                            top: { style: 'thick' },
                            left: { style: 'thick' },
                            bottom: { style: 'thick' },
                            right: { style: 'thick' }
                        };*/
                        ws.getCell('S7').value = items[2];
                        ws.getCell('H7').value = items[1];
                        ws.getCell('V7').value = items[6];
                        ws.getCell('H8').value = items[3];
                        ws.getCell('H9').value = items[9];
                        ws.getCell('S10').value = items[5];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('E14').value = items[4];
                        ws.getCell('N14').value = items[10].split('"')[0];
                        ws.getCell('H10').value = items[8];
                        var times = JSON.stringify(timeInMss).split('-');
                        ws.getCell('V26').value = times[0].split('"')[1];
                        ws.getCell('T26').value = times[1];
                        ws.getCell('Q26').value = times[2].split('T')[0];

                        for(let i=0;i<csvData.length;i++){
                            var ketluan = JSON.stringify(csvData[i]).split(",");
                            if(ketluan[1]=='"'+items[7]+'"'){
                                if(ketluan[2].length>1) {
                                    ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                }
                                if(ketluan[3].length>1) {
                                    ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                }
                                if(ketluan[4].length>5) {
                                    ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                }
                            }
                        }
                    } else{
                        err=true;
                    }
                }

                    if(err==false){
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                        res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua_dainghia' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                        return wb.xlsx.write(res)
                            .then(function () {
                                res.status(200).end();
                            });
                    }else {
                        err = true;
                        res.send('File upload bi loi');
                    }
                })
            })
        } else if (pattern == "5") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template_ductin.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                var datas = JSON.stringify(dataRaw).split(',');
                for (let i = 1; i < datas.length; i++) {
                    var items = datas[i].split(';');
                    if (items.length == 11) {
                        var tmpws = wb.getWorksheet('Template');
                        var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                        ws.model = tmpws.model;
                        ws.name = items[0].split('"')[1] + '_' + items[1];
                        /* ws.getCell('7', '8').value = items[1];*/

                        /*ws.mergeCells('A1', 'H1');
                        ws.mergeCells('A2', 'H2');*/
                        ws.mergeCells('I1', 'W1');
                        ws.mergeCells('I2', 'W2');
                        ws.mergeCells('C4', 'W4');
                        ws.mergeCells('C5', 'W5');
                        ws.mergeCells('C7', 'F7');
                        ws.mergeCells('C8', 'F8');
                        ws.mergeCells('C9', 'F9');
                        ws.mergeCells('C10', 'F10');
                        ws.mergeCells('C11', 'F11');
                        ws.mergeCells('07', 'R7');
                        ws.mergeCells('010', 'R10');
                        ws.mergeCells('E13', 'M13');
                        ws.mergeCells('N13', 'U13');
                        ws.mergeCells('E14', 'M14');
                        ws.mergeCells('N14', 'U14');
                        ws.mergeCells('D17', 'U21');
                        ws.mergeCells('G22', 'U23');
                        ws.mergeCells('G24', 'U24');
                        ws.mergeCells('M27', 'V27');
                        ws.mergeCells('M32', 'V32');
                        ws.mergeCells('M33', 'V33');

                        /*ws.getCell('E13').style.border = {
                            top: { style: 'thick' },
                            left: { style: 'thick' },
                            bottom: { style: 'thick' },
                            right: { style: 'thick' }
                        };*/
                        ws.getCell('S7').value = items[2];
                        ws.getCell('H7').value = items[1];
                        ws.getCell('V7').value = items[6];
                        ws.getCell('H8').value = items[3];
                        ws.getCell('H9').value = items[9];
                        ws.getCell('S10').value = items[5];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('E14').value = items[4];
                        ws.getCell('N14').value = items[10].split('"')[0];
                        ws.getCell('H10').value = items[8];
                        var times = JSON.stringify(timeInMss).split('-');
                        ws.getCell('V26').value = times[0].split('"')[1];
                        ws.getCell('T26').value = times[1];
                        ws.getCell('Q26').value = times[2].split('T')[0];

                        for(let i=0;i<csvData.length;i++){
                            var ketluan = JSON.stringify(csvData[i]).split(",");
                            if(ketluan[1]=='"'+items[7]+'"'){
                                if(ketluan[2].length>1) {
                                    ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                }
                                if(ketluan[3].length>1) {
                                    ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                }
                                if(ketluan[4].length>5) {
                                    ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                }
                            }
                        }
                    } else{
                        err=true;
                    }
                }

                if(err==false){
                    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                    res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua_ductin' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                    return wb.xlsx.write(res)
                        .then(function () {
                            res.status(200).end();
                        });
                }else {
                    err = true;
                    res.send('File upload bi loi');
                }
            })
        })
        } else if (pattern == "6") {
            var csvData = [];
            wb.xlsx.readFile('uploads/template_vincare.xlsx').then(function () {
                fs.createReadStream("public/csv/ketluanpattern.csv")
                    .pipe(parse({delimiter: ':'}))
                    .on('data', function (csvrow) {
                        var datas = JSON.stringify(csvrow).split(',');
                        for (let data of datas) {
                            var values = data.split(':');
                            var csvitems = values[1].split(';');
                            if (csvitems.length == 5) {
                                csvData.push(csvitems);
                            }
                        }
                    }).on('end', function () {
                var datas = JSON.stringify(dataRaw).split(',');
                for (let i = 1; i < datas.length; i++) {
                    var items = datas[i].split(';');
                    if (items.length == 11) {
                        var tmpws = wb.getWorksheet('Template');
                        var ws = wb.addWorksheet(items[0] + '_' + items[1]);
                        ws.model = tmpws.model;
                        ws.name = items[0].split('"')[1] + '_' + items[1];
                        /* ws.getCell('7', '8').value = items[1];*/

                        /*ws.mergeCells('A1', 'H1');
                        ws.mergeCells('A2', 'H2');*/
                        ws.mergeCells('I1', 'W1');
                        ws.mergeCells('I2', 'W2');
                        ws.mergeCells('C4', 'W4');
                        ws.mergeCells('C5', 'W5');
                        ws.mergeCells('C7', 'F7');
                        ws.mergeCells('C8', 'F8');
                        ws.mergeCells('C9', 'F9');
                        ws.mergeCells('C10', 'F10');
                        ws.mergeCells('C11', 'F11');
                        ws.mergeCells('07', 'R7');
                        ws.mergeCells('010', 'R10');
                        ws.mergeCells('E13', 'M13');
                        ws.mergeCells('N13', 'U13');
                        ws.mergeCells('E14', 'M14');
                        ws.mergeCells('N14', 'U14');
                        ws.mergeCells('D17', 'M21');
                        ws.mergeCells('N17', 'U21');
                        ws.mergeCells('G22', 'U23');
                        ws.mergeCells('G24', 'U24');
                        ws.mergeCells('M27', 'V27');
                        ws.mergeCells('M32', 'V32');
                        ws.mergeCells('M33', 'V33');

                        /*ws.getCell('E13').style.border = {
                            top: { style: 'thick' },
                            left: { style: 'thick' },
                            bottom: { style: 'thick' },
                            right: { style: 'thick' }
                        };*/
                        ws.getCell('S7').value = items[2];
                        ws.getCell('H7').value = items[1];
                        ws.getCell('V7').value = items[6];
                        ws.getCell('H8').value = items[3];
                        ws.getCell('H9').value = items[9];
                        ws.getCell('S10').value = items[5];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('H11').value = items[4];
                        ws.getCell('E14').value = items[4];
                        ws.getCell('N14').value = items[10].split('"')[0];
                        ws.getCell('H10').value = items[8];
                        var timeInMss = new Date();
                        var times = JSON.stringify(timeInMss).split('-');
                        ws.getCell('V26').value = times[0].split('"')[1];
                        ws.getCell('T26').value = times[1];
                        ws.getCell('Q26').value = times[2].split('T')[0];

                        for(let i=0;i<csvData.length;i++){
                            var ketluan = JSON.stringify(csvData[i]).split(",");
                            if(ketluan[1]=='"'+items[7]+'"'){
                                if(ketluan[2].length>1) {
                                    ws.getCell('D17').value = ketluan[2].substring(1, ketluan[2].length - 1);
                                }
                                if(ketluan[3].length>1) {
                                    ws.getCell('G22').value = ketluan[3].substring(1, ketluan[3].length - 1);
                                }
                                if(ketluan[4].length>5) {
                                    ws.getCell('G24').value = ketluan[4].substring(1, ketluan[4].length - 5);
                                }
                            }
                        }
                    } else{
                        err=true;
                    }
                }

                    if(err==false){
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                        res.setHeader('Content-Disposition', 'attachment; filename=' + 'ketqua_vincare' + JSON.stringify(timeInMss).split('T')[0] + '.xlsx');

                        return wb.xlsx.write(res)
                            .then(function () {
                                res.status(200).end();
                            });
                    }else {
                        err = true;
                        res.send('File upload bi loi');
                    }
                })
            })
        }
    }catch(e){
            res.send("Khong co don hang nao ca");
        }
});
