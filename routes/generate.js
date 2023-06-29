const express = require("express");
const router = express.Router();
const Jimp = require("jimp");
const { google } = require("googleapis");
const ExcelJS = require("exceljs");
const oauth2Client = require("../config/oauth");
const drive = google.drive({ version: "v3", auth: oauth2Client });
const Document = require("../models/Document");
const qr = require("qr-image");
const upload = require("../utils/upload");
const { join } = require("path");
const fs = require("fs");
const outputPath = join(__dirname, "..");

router.post(
  "/generate",
  upload.fields([
    { name: "spreadsheet", maxCount: 1 },
    { name: "template", maxCount: 1 },
  ]),
  async (req, res) => {
    const names = [],
      eventTitles = [],
      eventLocations = [],
      dateOfEvent = [];
    const workbook = new ExcelJS.Workbook();
    const spreadsheet = req.files["spreadsheet"][0].originalname;
    const template = req.files["template"][0].originalname;
    await workbook.xlsx
      .readFile(join(__dirname, "..", "/uploads/", spreadsheet))
      .then((data) => {
        const worksheet = workbook.getWorksheet(1); // Assuming you want to access the first worksheet

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            if (rowNumber != 1) {
              if (colNumber == 1) {
                names.push(cell.value);
              } else if (colNumber == 2) {
                eventTitles.push(cell.value);
              } else if (colNumber == 3) {
                eventLocations.push(cell.value);
              } else if (colNumber == 4) {
                dateOfEvent.push(cell.value);
              }
            }
          });
        });
      });

    addTextToImage(
      join(__dirname, "..", "/uploads/", template),
      names,
      eventTitles,
      eventLocations,
      dateOfEvent
    );

    async function addTextToImage(
      imagePath,
      names,
      eventTitles,
      eventLocations,
      dateOfEvent
    ) {
      const date = new Date().toJSON().slice(0, 10);

      const promises = names.map(async (data, index) => {
        const image = await Jimp.read(imagePath);
        const font = await Jimp.loadFont(Jimp.FONT_SANS_64_BLACK);
        const imageWidth = image.bitmap.width;
        const imageHeight = image.bitmap.height;
        image.print(font, imageWidth / 3 - 100, imageHeight / 2, data);
        image.write(outputPath + "/output/" + date + "-" + data + ".png");
        const newDoc = new Document({
          fullName: data,
          eventTitle: eventTitles[index],
          eventLocation: eventLocations[index],
          dateOfEvent: dateOfEvent[index],
        });
        return newDoc.save();
      });

      Promise.all(promises)
        .then((savedDocuments) => {
          savedDocuments.forEach((document) => {
            const docId = document._id.toString();
            const qrData = `localhost:3000/verify/${docId}`;
            const output = outputPath + "/qr/" + docId + "-qr.png";
            const qrCode = qr.image(qrData, { type: "png" });
            qrCode.pipe(fs.createWriteStream(output));
          });
          console.log("QR's created");
        })
        .catch((error) => {
          console.error("Error saving documents:", error);
        });

      //   for (const item of names) {
      //     drive.files
      //       .create({
      //         requestBody: {
      //           name: date + "-" + item + ".png",
      //           mimeType: "image/png",
      //         },
      //         media: {
      //           mimeType: "image/png",
      //           body: fs.createReadStream(
      //             outputPath + "/output/" + item + ".png"
      //           ),
      //         },
      //       })
      //       .then((res) => {
      //         console.log(`https://drive.google.com/uc?id=${res.data.id}`);
      //         const qrData = `https://drive.google.com/uc?id=${res.data.id}`;
      //         const output = outputPath + "/qr/" + item + "-qr.png";

      //       });
      //   }
      //   res.send("Boop");
    }
  }
);

module.exports = router;
