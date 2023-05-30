const path = require('path');
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const ejs = require('ejs');
const qr = require('qrcode');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.get('/', (req, res) => {
  res.render('index');
});

app.post('/generate', upload.fields([{ name: 'spreadsheet', maxCount: 1 }, { name: 'template', maxCount: 1 }]), (req, res) => {
    try {
      const spreadsheetPath = req.files['spreadsheet'][0].path;
      const templatePath = req.files['template'][0].path;
  
      const workbook = new ExcelJS.Workbook();
      workbook.xlsx.readFile(spreadsheetPath);
  
      const worksheet = workbook.worksheets[0];
      const names = worksheet.getColumn('A').values.slice(1); // Assuming the names start from row 3 in column A
  
      const templateContent = fs.readFileSync(templatePath, 'utf-8');
  
      const certificates = [];
  
      for (const name of names) {
        const qrCode = qr.toDataURL(name.toString());
  
        const certificate = ejs.render(templateContent, { name, qrCode });
        certificates.push(certificate);
      }
  
      console.log('Certificates:', certificates);
  
      res.render('certificates', { certificates: certificates.join('') });
    } catch (error) {
      console.error(error);
      res.sendStatus(500);
    }
  });
app.use(express.static(path.join(__dirname, 'public')));

app.listen(3000, () => {
  console.log('App listening on http://localhost:3000');
});
