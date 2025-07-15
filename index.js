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

// Columns to auto resize (shifted by +3 since A is blank and data starts at D=3)
const autoResizeCols = [4, 5, 19, 20, 21, 22];
// Wrapped columns = all from D(3) to W(22) excluding autoResizeCols
const wrapTextCols = [];
for (let i = 3; i <= 22; i++) {
  if (!autoResizeCols.includes(i)) wrapTextCols.push(i);
}

// Health check endpoint for uptime monitoring
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
    const existingSheet = sheetData.find((s) => s.properties.title === sheetTitle);
    let sheetId;

    if (!existingSheet) {
      // Create new sheet
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: sheetTitle } } }],
        },
      });

      const newMeta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheet = newMeta.data.sheets.find((s) => s.properties.title === sheetTitle);
      sheetId = sheet.properties.sheetId;

      // Set header row in row 2, columns B to W (2 to 22 zero-based indices)
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!B2:W2`,
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

      // Requests for formatting & resizing
      const requests = [
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 1, // B column index
              endIndex: 23, // W is 22 zero-based, endIndex is exclusive so 23
            },
            properties: { pixelSize: 180 },
            fields: "pixelSize",
          },
        },
        // Header row green background, bold, borders, font for B2:W2
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: 1,
              endColumnIndex: 23,
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
        // Wrap text for header row on wrapTextCols only
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
        // No wrap on autoResizeCols for header row
        ...autoResizeCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "OVERFLOW" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        // Data rows: Times New Roman font on B3:W1000
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: 1,
              endColumnIndex: 23,
            },
            cell: {
              userEnteredFormat: {
                textFormat: { fontFamily: "Times New Roman" },
              },
            },
            fields: "userEnteredFormat.textFormat",
          },
        },
        // Wrap on wrapTextCols for data rows
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
        // No wrap on autoResizeCols for data rows
        ...autoResizeCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: { userEnteredFormat: { wrapStrategy: "OVERFLOW" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
      ];

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

      // Auto resize specified columns individually for better fit
      for (const colIndex of autoResizeCols) {
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

    // Prepare current date and time strings (e.g. YYYY-MM-DD and HH:mm:ss)
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const timeStr = now.toTimeString().slice(0, 8);

    // Append the data row, shifted by 2 columns to start at B (Date) and C (Time)
    const appendResult = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!B3:W3`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          dateStr,
          timeStr,
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

    // Determine the appended row index (0-based)
    const updatedRange = appendResult.data.updates.updatedRange;
    const match = updatedRange.match(/!B(\d+):W(\d+)/);
    let firstRowIndex;
    if (match) {
      firstRowIndex = parseInt(match[1], 10) - 1; // zero-based
    } else {
      firstRowIndex = 2; // fallback to row 3 zero-based
    }
    const lastRowIndex = firstRowIndex;

    // Format appended row with green background, borders, and wrap logic
    const appendRequests = [
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: lastRowIndex + 1,
            startColumnIndex: 1,
            endColumnIndex: 23,
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
      // Wrap text on wrapTextCols
      ...wrapTextCols.map((colIndex) => ({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: lastRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
          fields: "userEnteredFormat.wrapStrategy",
        },
      })),
      // No wrap on autoResizeCols
      ...autoResizeCols.map((colIndex) => ({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: lastRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: { userEnteredFormat: { wrapStrategy: "OVERFLOW" } },
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
