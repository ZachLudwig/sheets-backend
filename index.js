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

// Columns to auto resize (indexes start from B=1, so D=3 etc.)
// Adjusted columns for fewer columns, up to U (index 21)
const autoResizeCols = [4, 5, 19, 20]; // D=3, E=4, S=18, T=19 originally? Adjusted below
// Let’s explicitly use 1-based indices converted to zero-based here for clarity:
const autoResizeColsAdjusted = [4, 5, 19, 20]; // stays same as before but max 20 for U column

// Wrap text columns: from D(3) to U(20), excluding autoResizeCols
const wrapTextCols = [];
for (let i = 3; i <= 20; i++) {
  if (!autoResizeColsAdjusted.includes(i)) wrapTextCols.push(i);
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

      // Header row from B2 to U2 (columns 2 to 21 inclusive)
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

      // Formatting requests
      const requests = [
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 1, // B = 1
              endIndex: 21, // U = 20 zero-based + 1
            },
            properties: { pixelSize: 180 },
            fields: "pixelSize",
          },
        },
        // Header row formatting with green background and borders for B2:U2
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
        // Wrap text on wrapTextCols (all but autoResizeCols) for header row
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
        // Wrap text on autoResizeCols for header row
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
        // Data rows font and wrap text B3:U1000
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: 1,
              endColumnIndex: 21,
            },
            cell: {
              userEnteredFormat: {
                textFormat: { fontFamily: "Times New Roman" },
              },
            },
            fields: "userEnteredFormat.textFormat",
          },
        },
        // Wrap text on wrapTextCols in data rows
        ...wrapTextCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        // Wrap text on autoResizeCols in data rows
        ...autoResizeColsAdjusted.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
      ];

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

      // Auto resize columns individually
      for (const colIndex of autoResizeColsAdjusted) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                autoResizeDimensions: {
                  dimensions: {
                    sheetId,
                    dimension: "COLUMNS",
                    startIndex: colIndex,
                    endIndex: colIndex + 1,
                  },
                },
              },
            ],
          },
        });
      }
    } else {
      sheetId = existingSheet.properties.sheetId;
    }

    // Get current EST date/time strings
    const now = new Date();
    const estDateStr = now.toLocaleDateString("en-US", {
      timeZone: "America/New_York",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).split("/").reverse().join("-"); // YYYY-MM-DD

    const estTimeStr = now.toLocaleTimeString("en-US", {
      timeZone: "America/New_York",
      hour12: false,
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });

    // Append data starting from B3:U3
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

    // Extract appended row index from updatedRange, zero-based
    const updatedRange = appendResult.data.updates.updatedRange;
    const match = updatedRange.match(/!B(\d+):U(\d+)/);
    let appendedRowIndex;
    if (match) {
      appendedRowIndex = parseInt(match[1], 10) - 1; // zero-based
    } else {
      appendedRowIndex = 2; // fallback for row 3 zero-based
    }

    // Format appended row: green background, borders, wrap text on cols
    const appendRequests = [
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: appendedRowIndex,
            endRowIndex: appendedRowIndex + 1,
            startColumnIndex: 1, // B
            endColumnIndex: 21,  // U + 1
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
      // Wrap text on wrapTextCols in appended row
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
      // Wrap text on autoResizeCols in appended row
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

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: appendRequests },
    });

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
