const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const { google } = require("googleapis");
require("dotenv").config();

const app = express();

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Google Sheets Auth
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.SERVICE_ACCOUNT_KEY_FILE,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

app.post("/export-user-data", async (req, res) => {
  try {
    const { username, email, value1, value2 } = req.body;

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const sheetTitle = username;

    // 1. Check if sheet tab exists
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const exists = meta.data.sheets.some(
      (sheet) => sheet.properties.title === sheetTitle
    );

    // 2. Create sheet tab if it doesn't exist
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: { properties: { title: sheetTitle } },
            },
          ],
        },
      });

      // Add header row to new sheet
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!A1`,
        valueInputOption: "RAW",
        requestBody: {
          values: [["Username", "Email", "Value1", "Value2"]],
        },
      });

      console.log(`Created sheet tab "${sheetTitle}" and added header.`);
    }

    // 3. Append new row to the sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!A1`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[username, email, value1, value2]],
      },
    });

    res.status(200).json({ message: "Data exported successfully." });
  } catch (error) {
    console.error("Export failed:", error);
    res.status(500).json({ error: "Failed to export data." });
  }
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Okay Zach, Backend running on port ${PORT}`));
