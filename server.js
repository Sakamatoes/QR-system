const path = require('path');
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const ejs = require('ejs');
const qr = require('qrcode');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module');
const PizZip = require("pizzip");
const Jimp = require("jimp");

const app = express();
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname + "/uploads/");
  },
  filename: (req, file, cb) => {
    console.log(file)
    cb(null, file.originalname);
  }
});

const upload = multer({storage: storage});

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.get('/', (req, res) => {
  res.render('index');
});

app.post('/generate', upload.fields([{name: 'spreadsheet', maxCount: 1}, {name: 'template', maxCount: 1}]), (req, res) => {

  const workbook = new ExcelJS.Workbook();
    const spreadsheet = workbook.getWorksheet('spreadsheet');
    console.log(workbook.spreadsheet[0]);
  });

app.get("/test", async (req, res) => {
  const rowsArray = [];
  const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(__dirname + '/uploads/spreadsheet.xlsx').then(data => {
      const worksheet = workbook.getWorksheet(1); // Assuming you want to access the first worksheet
  
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowData = [];
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        rowData.push(cell.value);
      });
      rowsArray.push(rowData);
    });
  });
  
async function addTextToImage(imagePath, text) {
    const image = await Jimp.read(imagePath);
    const font = await Jimp.loadFont(Jimp.FONT_SANS_32_BLACK);
    const imageWidth = image.bitmap.width;
    const imageHeight = image.bitmap.height;
    image.print(font, imageWidth / 2, imageHeight / 2, text);
    image.write(imagePath);
}

addTextToImage(__dirname + '/uploads/template.png', 'Hello World!');
});
app.use(express.static(path.join(__dirname, 'public')));

app.listen(3000, () => {
  console.log('App listening on http://localhost:3000');
});
