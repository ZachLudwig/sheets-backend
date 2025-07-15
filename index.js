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
    const {
      username,
      age,
      gender,
      preSkill,
      postSkill,
      challenged,
      attentionFocus,
      enjoyment,
      control,
      wantedToAchieve,
      challengeSkillBalance,
      goalsDefined,
      awarePerformance,
      timeAlter,
      performedAutomatically,
      goal,
      goalAchievedTF,
      goalAchievedReason,
      comments,
    } = req.body;

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const sheetTitle = username;

    // Check if the sheet exists
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetData = meta.data.sheets;
    const existingSheet = sheetData.find((s) => s.properties.title === sheetTitle);
    let sheetId;

    if (!existingSheet) {
      // Add new sheet
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

      // Get the new sheetId
      const newMeta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheet = newMeta.data.sheets.find((s) => s.properties.title === sheetTitle);
      sheetId = sheet.properties.sheetId;

      // Set header row values
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!B2:S2`,
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            "Age",
            "Gender",
            "Skill chosen before activity",
            "Skill chosen after activity",
            "I was challenged, but I believed my skills would allow me to meet the challenge.",
            "My attention was focused entirely on what I was doing.",
            "I really enjoyed the experience.",
            "I felt like I could control what I was doing.",
            "I knew what I wanted to achieve.",
            "The challenge and my skills were at an equally high level.",
            "My goals were clearly defined.",
            "I was aware of how well I was performing.",
            "Time seemed to alter (either slowed down or sped up).",
            "I performed automatically.",
            "Goal of activity",
            "Goal Achieved?",
            "Why goal was/wasn't met?",
            "Further comments"
          ]],
        },
      });

      // Set column widths and style header row (bold, borders, no bg)
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateDimensionProperties: {
                range: {
                  sheetId,
                  dimension: "COLUMNS",
                  startIndex: 1,
                  endIndex: 19,
                },
                properties: { pixelSize: 150 },
                fields: "pixelSize",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: 1,  // row 2 (header)
                  endRowIndex: 2,
                  startColumnIndex: 1,
                  endColumnIndex: 19,
                },
                cell: {
                  userEnteredFormat: {
                    textFormat: {
                      bold: true,
                      fontFamily: "Times New Roman",
                    },
                    borders: {
                      top: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                      bottom: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                      left: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                      right: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                    },
                  },
                },
                fields: "userEnteredFormat(textFormat,borders)",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: 0,
                  endRowIndex: 1000,
                  startColumnIndex: 0,
                  endColumnIndex: 26,
                },
                cell: {
                  userEnteredFormat: {
                    textFormat: {
                      fontFamily: "Times New Roman",
                    },
                  },
                },
                fields: "userEnteredFormat.textFormat.fontFamily",
              },
            },
          ],
        },
      });
    } else {
      sheetId = existingSheet.properties.sheetId;
    }

    // Append new data row
    const appendResult = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!B3:S3`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          age,
          gender,
          preSkill,
          postSkill,
          challenged,
          attentionFocus,
          enjoyment,
          control,
          wantedToAchieve,
          challengeSkillBalance,
          goalsDefined,
          awarePerformance,
          timeAlter,
          performedAutomatically,
          goal,
          goalAchievedTF,
          goalAchievedReason,
          comments,
        ]],
      },
    });

    // Get the starting row index of appended data (Google Sheets API returns it as 1-based)
    const updatedRange = appendResult.data.updates.updatedRange; 
    // Example updatedRange: "SheetName!B3:S3" or "SheetName!B5:S5"
    const match = updatedRange.match(/!B(\d+):S(\d+)/);
    let firstRowIndex;
    if (match) {
      firstRowIndex = parseInt(match[1], 10) - 1; // convert to zero-based index for batchUpdate
    } else {
      firstRowIndex = 2; // fallback to row 3 zero-based index
    }

    // Apply green background styling to only the newly appended row
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: firstRowIndex,
                endRowIndex: firstRowIndex + 1,
                startColumnIndex: 1,
                endColumnIndex: 19,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 0.85,
                    green: 1.0,
                    blue: 0.85,
                  },
                  borders: {
                    top: { style: "SOLID", color: { red: 0, green: 0, blue: 0 } },
                    bottom: { style: "SOLID", color: { red: 0, green: 0, blue: 0 } },
                    left: { style: "SOLID", color: { red: 0, green: 0, blue: 0 } },
                    right: { style: "SOLID", color: { red: 0, green: 0, blue: 0 } },
                  },
                  textFormat: {
                    fontFamily: "Times New Roman",
                    bold: false,
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor,borders,textFormat)",
            },
          },
        ],
      },
    });

    res.status(200).json({ message: "Data exported and styled successfully." });
  } catch (error) {
    console.error("Export failed:", error);
    res.status(500).json({ error: "Failed to export data." });
  }
});

app.get("/health", (req, res) => {
  res.setHeader("Content-Type", "text/plain");
  res.status(200).send("OK");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log(`✅ Okay Zach, Backend running on port ${PORT}`)
);
