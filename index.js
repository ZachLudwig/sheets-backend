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

const autoResizeCols = [1, 2, 16, 17, 18, 19];
const wrapTextCols = [];
for (let i = 1; i <= 19; i++) {
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
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            { addSheet: { properties: { title: sheetTitle } } },
          ],
        },
      });

      const newMeta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheet = newMeta.data.sheets.find((s) => s.properties.title === sheetTitle);
      sheetId = sheet.properties.sheetId;

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

      const requests = [
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 0,
              endIndex: 26,
            },
            properties: { pixelSize: 100 },
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
              endColumnIndex: 19,
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
        // Set wrap strategy for wrapped columns
        ...wrapTextCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                wrapStrategy: "WRAP",
              },
            },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        // Set wrap strategy for auto resize columns
        ...autoResizeCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: 2,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                wrapStrategy: "WRAP",
              },
            },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
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
            fields: "userEnteredFormat.textFormat",
          },
        },
        // Wrap for rows after header (body)
        ...wrapTextCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                wrapStrategy: "WRAP",
              },
            },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        // Wrap for auto resize cols in body
        ...autoResizeCols.map((colIndex) => ({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 1000,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                wrapStrategy: "WRAP",
              },
            },
            fields: "userEnteredFormat.wrapStrategy",
          },
        })),
        // Set wrapped columns width explicitly to 180 pixels
        ...wrapTextCols.map((colIndex) => ({
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: colIndex,
              endIndex: colIndex + 1,
            },
            properties: { pixelSize: 180 },
            fields: "pixelSize",
          },
        })),
      ];

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

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

    const updatedRange = appendResult.data.updates.updatedRange;
    const match = updatedRange.match(/!B(\d+):S(\d+)/);
    let firstRowIndex;
    if (match) {
      firstRowIndex = parseInt(match[1], 10) - 1;
    } else {
      firstRowIndex = 2;
    }
    const lastRowIndex = firstRowIndex;

    const appendRequests = [];

    appendRequests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: firstRowIndex,
          endRowIndex: lastRowIndex + 1,
          startColumnIndex: 1,
          endColumnIndex: 19,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.85, green: 1.0, blue: 0.85 },
            borders: {
              top: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
              left: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
              right: { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } },
            },
            textFormat: { fontFamily: "Times New Roman" },
          },
        },
        fields: "userEnteredFormat(backgroundColor,borders,textFormat)",
      },
    });

    for (const colIndex of wrapTextCols) {
      appendRequests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: lastRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: {
            userEnteredFormat: {
              wrapStrategy: "WRAP",
            },
          },
          fields: "userEnteredFormat.wrapStrategy",
        },
      });
    }

    for (const colIndex of autoResizeCols) {
      appendRequests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: firstRowIndex,
            endRowIndex: lastRowIndex + 1,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: {
            userEnteredFormat: {
              wrapStrategy: "WRAP",
            },
          },
          fields: "userEnteredFormat.wrapStrategy",
        },
      });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: appendRequests },
    });

    res.status(200).json({ message: "User data appended successfully." });
  } catch (error) {
    console.error("Error appending user data:", error);
    res.status(500).json({ error: error.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
