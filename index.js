const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const { google } = require("googleapis");
require("dotenv").config();

const app = express();

app.use(cors());
app.use(bodyParser.json());

const rawServiceAccount = process.env.SERVICE_ACCOUNT_JSON;
const serviceAccount = JSON.parse(rawServiceAccount);
serviceAccount.private_key = serviceAccount.private_key.replace(/\\n/g, "\n");

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

app.post("/export-user-data", async (req, res) => {
  try {
    const { username, email, value1, value2 } = req.body;

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const sheetTitle = username;

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const exists = meta.data.sheets.some(
      (sheet) => sheet.properties.title === sheetTitle
    );

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

      const sheetId = (
        await sheets.spreadsheets.get({ spreadsheetId })
      ).data.sheets.find((s) => s.properties.title === sheetTitle).properties
        .sheetId;

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!B2`,
        valueInputOption: "RAW",
        requestBody: {
          values: [["Username", "Email", "Value1", "Value2"]],
        },
      });

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateDimensionProperties: {
                range: {
                  sheetId,
                  dimension: "COLUMNS",
                  startIndex: 0,
                  endIndex: 1,
                },
                properties: {
                  pixelSize: 50,
                },
                fields: "pixelSize",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: 0,
                  endRowIndex: 500,
                  startColumnIndex: 0,
                  endColumnIndex: 26,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: {
                      red: 0.9,
                      green: 0.9,
                      blue: 0.9,
                    },
                  },
                },
                fields: "userEnteredFormat.backgroundColor",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: 1,
                  endRowIndex: 2,
                  startColumnIndex: 1,
                  endColumnIndex: 5,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: {
                      red: 0.8,
                      green: 1,
                      blue: 0.8,
                    },
                    borders: {
                      top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                      bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                      left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                      right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                    },
                  },
                },
                fields: "userEnteredFormat(backgroundColor,borders)",
              },
            },
          ],
        },
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!B2`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[username, email, value1, value2]],
      },
    });

    const dataRange = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetTitle}!B3:B`,
    });
    const rowCount = dataRange.data.values?.length || 0;
    const rowIndex = 2 + rowCount;

    const sheetId = meta.data.sheets.find(
      (s) => s.properties.title === sheetTitle
    ).properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: rowIndex,
                endRowIndex: rowIndex + 1,
                startColumnIndex: 1,
                endColumnIndex: 5,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 0.8,
                    green: 1,
                    blue: 0.8,
                  },
                  borders: {
                    top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                    bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                    left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                    right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor,borders)",
            },
          },
        ],
      },
    });

    res.status(200).json({ message: "Data exported successfully." });
  } catch (error) {
    console.error("Export failed:", error);
    res.status(500).json({ error: "Failed to export data." });
  }
});

app.get("/health", (req, res) => {
  console.log("💓 Health check ping received");
  res.setHeader("Content-Type", "text/plain");
  res.status(200).send("OK");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log(`✅ Okay Zach, Backend running on port ${PORT}`)
);
