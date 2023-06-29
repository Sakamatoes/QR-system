const express = require("express");
const fs = require("fs");
const cors = require("cors");
const mongoose = require("mongoose");
const dotenv = require("dotenv");
dotenv.config();
const oauth2Client = require("./config/oauth");

mongoose.connect("mongodb://127.0.0.1:27017/InTTO-QR-System").then(() => {
  console.log("Database Connected!");
});

const app = express();
app.use(cors());
app.use("/public", express.static(__dirname + "/public"));
app.use(require("./routes/auth"));
app.use(require("./routes/redirect"));
app.use(require("./routes/generate"));

try {
  const creds = fs.readFileSync("creds.json");
  oauth2Client.setCredentials(JSON.parse(creds));
} catch (err) {
  console.log("No creds found");
}

app.listen(3000, () => {
  console.log("App listening on http://localhost:3000");
});
