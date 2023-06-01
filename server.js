const path = require("path");
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const ejs = require("ejs");
const qr = require("qrcode");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module");
const PizZip = require("pizzip");
const Jimp = require("jimp");

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
app.set("views", path.join(__dirname, "views"));
app.use("/public", express.static(__dirname + "/public"));

app.get("/", (req, res) => {
	res.render("index");
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
		await workbook.xlsx
			.readFile(__dirname + "/uploads/spreadsheet.xlsx")
			.then((data) => {
				const worksheet = workbook.getWorksheet(1); // Assuming you want to access the first worksheet

				worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
					row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
						rowData.push(cell.value);
					});
					rowsArray.push(rowData);
				});
			});

		addTextToImage(__dirname + "/uploads/template.png", rowData);

		async function addTextToImage(imagePath, array) {
			var index = 0;
			for (const data of array) {
				index++;
				const image = await Jimp.read(imagePath);
				const font = await Jimp.loadFont(Jimp.FONT_SANS_64_BLACK);
				const imageWidth = image.bitmap.width;
				const imageHeight = image.bitmap.height;
				image.print(font, imageWidth / 3 - 100, imageHeight / 2, data);
				image.write(__dirname + "/output/newImage" + index + ".png");
			}
		}
	}
);

app.listen(3000, () => {
	console.log("App listening on http://localhost:3000");
});
