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

const paddedCols = [1, 2, 3]; // B, C, D
const autoResizeColsAdjusted = [4, 5, 19, 20]; // E, F, T, U
const wrapTextCols = [];
for (let i = 3; i <= 20; i++) {
  if (![...autoResizeColsAdjusted, ...paddedCols].includes(i)) {
    wrapTextCols.push(i);
  }
}

app.get("/health", (req, res) => {
  console.log("💓 Health check ping received");
  res.setHeader("Content-Type", "text/plain");
  res.status(200).send("OK");
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

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetData = meta.data.sheets;
    let sheetId;

    const existingSheet = sheetData.find((s) => s.properties.title === sheetTitle);

    if (!existingSheet) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: sheetTitle } } }],
        },
      });

      const newMeta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheet = newMeta.data.sheets.find((s) => s.properties.title === sheetTitle);
      sheetId = sheet.properties.sheetId;

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!B2:U2`,
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            "Date",
            "Time Recorded",
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
            "Further comments",
          ]],
        },
      });

      const requests = [
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 1,
              endIndex: 21,
            },
            properties: { pixelSize: 180 },
            fields: "pixelSize",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: 1,
              endColumnIndex: 21,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.85, green: 1.0, blue: 0.85 },
                textFormat: { bold: true, fontFamily: "Times New Roman" },
                borders: {
                  top: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                  bottom: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                  left: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                  right: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,borders)",
          },
        },
        ...wrapTextCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        ...autoResizeColsAdjusted.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
      ];

      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
    } else {
      sheetId = existingSheet.properties.sheetId;
    }

    const now = new Date();
    const estDateStr = now.toLocaleDateString("en-US", {
      timeZone: "America/New_York",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).split("/").reverse().join("-");

    const estTimeStr = now.toLocaleTimeString("en-US", {
      timeZone: "America/New_York",
      hour12: false,
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });

    const appendResult = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!B3:U3`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          estDateStr,
          estTimeStr,
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

    const match = appendResult.data.updates.updatedRange.match(/!B(\d+):U(\d+)/);
    const appendedRowIndex = match ? parseInt(match[1], 10) - 1 : 2;

    const appendRequests = [
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: appendedRowIndex,
            endRowIndex: appendedRowIndex + 1,
            startColumnIndex: 1,
            endColumnIndex: 21,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.85, green: 1.0, blue: 0.85 },
              borders: {
                top: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                bottom: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                left: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
                right: { style: "SOLID", width: 2, color: { red: 0, green: 0, blue: 0 } },
              },
            },
          },
          fields: "userEnteredFormat(backgroundColor,borders)",
        },
      },
      ...wrapTextCols.map((colIndex) => ({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: appendedRowIndex,
            endRowIndex: appendedRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
          fields: "userEnteredFormat.wrapStrategy",
        },
      })),
      ...autoResizeColsAdjusted.map((colIndex) => ({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: appendedRowIndex,
            endRowIndex: appendedRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
          fields: "userEnteredFormat.wrapStrategy",
        },
      })),
    ];

    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: appendRequests } });

    // Auto-resize and pad B, C, D columns
    for (const colIndex of paddedCols) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            autoResizeDimensions: {
              dimensions: {
                sheetId,
                dimension: "COLUMNS",
                startIndex: colIndex,
                endIndex: colIndex + 1,
              },
            },
          }],
        },
      });

      // Add 10px padding (estimated default width + 10, e.g. 110)
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            updateDimensionProperties: {
              range: {
                sheetId,
                dimension: "COLUMNS",
                startIndex: colIndex,
                endIndex: colIndex + 1,
              },
              properties: { pixelSize: 110 },
              fields: "pixelSize",
            },
          }],
        },
      });
    }

    res.status(200).json({ message: "User data exported successfully." });
  } catch (error) {
    console.error("Error exporting user data:", error);
    res.status(500).json({ error: "Internal server error." });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
