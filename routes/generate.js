const express = require("express");
const router = express.Router();
const Jimp = require("jimp");
const { google } = require("googleapis");
const ExcelJS = require("exceljs");
const oauth2Client = require("../config/oauth");
const drive = google.drive({ version: "v3", auth: oauth2Client });
const mongoose = require("mongoose");
const Document = require("../models/Document");
const qr = require("qr-image");
const upload = require("../utils/upload");
const { join } = require("path");
const fs = require("fs");
const outputPath = join(__dirname, "..");
const app = express();
app.use(express.json());

router.post(
  "/generate",
  upload.fields([
    { name: "spreadsheet", maxCount: 1 },
    { name: "template", maxCount: 1 },
  ]),
  async (req, res) => {
    const name_coordinates = JSON.parse(req.body.name_coords);
    const qr_coordinates = JSON.parse(req.body.qr_coords);
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
        const font = await Jimp.loadFont(Jimp.FONT_SANS_128_BLACK);
        const textWidth = Jimp.measureText(font, data);
        const centerX = (name_coordinates.startX + name_coordinates.endX) / 2;
        const startingX = centerX - textWidth / 2;

        image.print(font, startingX, name_coordinates.startY, data);
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
          savedDocuments.forEach(async (document) => {
            const date = new Date().toJSON().slice(0, 10);
            const docId = document._id.toString();
            const qrData = `localhost:3000/verify/${docId}`;
            const output = outputPath + "/qr/" + docId + "-qr.png";
            const qrCode = qr.image(qrData, {
              type: "png",
              ec_level: "L",
            });
            qrCode.pipe(fs.createWriteStream(output));
          });
          return savedDocuments;
        })
        .then((data) => {
          const promise = data.map(async (document) => {
            const date = new Date().toJSON().slice(0, 10);

            const image = await Jimp.read(
              join(
                __dirname,
                "..",
                "/output/",
                date + "-" + document.fullName + ".png"
              )
            );

            const qr = await Jimp.read(
              join(__dirname, "..", "/qr/", document._id + "-qr.png")
            );

            image.composite(qr, qr_coordinates.endX, qr_coordinates.endY);

            await image.writeAsync(
              join(
                __dirname,
                "..",
                "/drive/",
                date + "-" + document.fullName + ".png"
              )
            );

            const driveFile = await drive.files.create({
              requestBody: {
                name: date + "-" + document.fullName + ".png",
                mimeType: "image/png",
              },
              media: {
                mimeType: "image/png",
                body: fs.createReadStream(
                  join(
                    __dirname,
                    "..",
                    "/drive/",
                    date + "-" + document.fullName + ".png"
                  )
                ),
              },
            });

            const url = `https://drive.google.com/uc?id=${driveFile.data.id}`;
            await Document.updateOne(
              { _id: document._id },
              { certificate: url }
            );

            console.log(url);
          });
          Promise.all(promise)
            .then(() => {
              console.log("All files created and uploaded successfully.");
              res.send("Success");
            })
            .catch((error) => {
              console.error("Error:", error);
            });
        })
        .catch((error) => {
          console.error("Error saving documents:", error);
        });
    }
  }
);

module.exports = router;
