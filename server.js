const path = require("path");
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const ejs = require("ejs");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module");
const PizZip = require("pizzip");
const Jimp = require("jimp");
const cors = require("cors");
const qr = require("qr-image");

const { google } = require("googleapis");
const dotenv = require("dotenv");
dotenv.config();
const scopes = require("./config/scopes");

const app = express();
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname + "/uploads/");
  },
  filename: (req, file, cb) => {
    console.log(file);
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

app.set("view engine", "ejs");
app.use(cors());
app.set("views", path.join(__dirname, "views"));
app.use("/public", express.static(__dirname + "/public"));

const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

const drive = google.drive({ version: "v3", auth: oauth2Client });

try {
  const creds = fs.readFileSync("creds.json");
  oauth2Client.setCredentials(JSON.parse(creds));
} catch (err) {
  console.log("No creds found");
}

app.get("/auth/google", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: scopes,
  });
  res.redirect(url);
});

app.get("/google/redirect", async (req, res) => {
  const { code } = req.query;
  const { tokens } = await oauth2Client.getToken(code);
  oauth2Client.setCredentials(tokens);
  fs.writeFileSync("creds.json", JSON.stringify(tokens));
  res.send("success");
});

app.post(
  "/generate",
  upload.fields([
    { name: "spreadsheet", maxCount: 1 },
    { name: "template", maxCount: 1 },
  ]),
  async (req, res) => {
    const rowsArray = [];
    const workbook = new ExcelJS.Workbook();
    const rowData = [];
    const spreadsheet = req.files["spreadsheet"][0].originalname;
    const template = req.files["template"][0].originalname;
    await workbook.xlsx
      .readFile(__dirname + "/uploads/" + spreadsheet)
      .then((data) => {
        const worksheet = workbook.getWorksheet(1); // Assuming you want to access the first worksheet

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            rowData.push(cell.value);
          });
          rowsArray.push(rowData);
        });
      });

    addTextToImage(__dirname + "/uploads/" + template, rowData);

    async function addTextToImage(imagePath, array) {
      const date = new Date().toJSON().slice(0, 10);
      var index = 0;
      for (const data of array) {
        index++;
        const image = await Jimp.read(imagePath);
        const font = await Jimp.loadFont(Jimp.FONT_SANS_64_BLACK);
        const imageWidth = image.bitmap.width;
        const imageHeight = image.bitmap.height;
        image.print(font, imageWidth / 3 - 100, imageHeight / 2, data);
        image.write(__dirname + "/output/" + data + ".png");
      }
      for (const item of array) {
        drive.files
          .create({
            requestBody: {
              name: date + "-" + item + ".png",
              mimeType: "image/png",
            },
            media: {
              mimeType: "image/png",
              body: fs.createReadStream(__dirname + "/output/" + item + ".png"),
            },
          })
          .then((res) => {
            console.log(`https://drive.google.com/uc?id=${res.data.id}`);
            const qrData = `https://drive.google.com/uc?id=${res.data.id}`;
            const outputPath = __dirname + "/qr/" + item + "-qr.png";

            const qrCode = qr.image(qrData, { type: "png" });
            qrCode.pipe(fs.createWriteStream(outputPath));
          });
      }
      res.send("Boop");
    }
  }
);

app.listen(3000, () => {
  console.log("App listening on http://localhost:3000");
});
