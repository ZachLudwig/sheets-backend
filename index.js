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

    // Get metadata to check if sheet exists
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetData = meta.data.sheets;
    const existingSheet = sheetData.find((s) => s.properties.title === sheetTitle);
    const exists = !!existingSheet;
    let sheetId = exists ? existingSheet.properties.sheetId : undefined;

    if (!exists) {
      // Create new sheet with username as title
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

      // Get new sheetId after creation
      sheetId = (
        await sheets.spreadsheets.get({ spreadsheetId })
      ).data.sheets.find((s) => s.properties.title === sheetTitle).properties.sheetId;

      // Set headers in B2:Q2 with your custom titles
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!B2:Q2`,
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

      // Format columns width and background for header row
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateDimensionProperties: {
                range: {
                  sheetId,
                  dimension: "COLUMNS",
                  startIndex: 1, // B column
                  endIndex: 18,  // Q column (B to Q is 17 columns)
                },
                properties: { pixelSize: 150 },
                fields: "pixelSize",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: 1,  // Row 2 (0-indexed)
                  endRowIndex: 2,
                  startColumnIndex: 1,
                  endColumnIndex: 18,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: {
                      red: 0.9,
                      green: 0.9,
                      blue: 0.9,
                    },
                    textFormat: {
                      bold: true,
                      fontFamily: "Times New Roman",
                    },
                    borders: {
                      bottom: { style: "SOLID_MEDIUM", width: 2, color: { red: 0, green: 0, blue: 0 } },
                    },
                  },
                },
                fields: "userEnteredFormat(backgroundColor,textFormat,borders.bottom)",
              },
            },
            {
              // Set entire sheet font to Times New Roman
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
    }

    // Append the new row (without username)
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!B3:Q3`,
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
