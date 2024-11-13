const express = require("express");
const path = require("path");
const cors = require("cors");
const axios = require("axios");
const xlsx = require("xlsx");
const fs = require("fs");
const multer = require("./multer");

const app = express();
app.use(express.json());
app.use(
  cors({
    origin: "*",
  })
);

app.use("/uploads", express.static(path.join(__dirname, "uploads/")));
app.use("/uploads", express.static(path.join(__dirname, "./uploads/")));

app.post("/upload-excel", multer.single("file"), async (req, res) => {
  try {
    const filePath = req.file?.path;

    if (!filePath) {
      return res.status(400).json({ error: "File upload failed" });
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert the worksheet to JSON for easy manipulation
    let data = xlsx.utils.sheet_to_json(worksheet, {
      raw: false,
      defval: null,
    });

    // Parse dates and calculate Work Hours and Rest Hours
    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      // Parse 'POB' and 'Completion Date' if they exist
      const pob = row["POB"] ? new Date(row["POB"]) : null;
      const completionDate = row["Completion Date"]
        ? new Date(row["Completion Date"])
        : null;

      // Calculate Work Hours as (Completion Date - POB) in hours
      if (pob && completionDate) {
        const workHours = (completionDate - pob) / (1000 * 60 * 60); // Convert ms to hours
        row["Work Hours"] = workHours;
      } else {
        row["Work Hours"] = null;
      }

      // Calculate Rest Hours only if there's a previous Completion Date
      if (i > 0 && completionDate) {
        const previousCompletionDate = data[i - 1]["Completion Date"]
          ? new Date(data[i - 1]["Completion Date"])
          : null;
        if (previousCompletionDate) {
          const restHours = (pob - previousCompletionDate) / (1000 * 60 * 60); // Convert ms to hours
          row["Rest Hours"] = restHours;
        } else {
          row["Rest Hours"] = null;
        }
      } else {
        row["Rest Hours"] = null;
      }
    }

    // Filter out any keys that start with '__EMPTY'
    data = data.map((row) => {
      const filteredRow = {};
      Object.keys(row).forEach((key) => {
        if (!key.startsWith("__EMPTY")) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    // Create a new worksheet from the modified data
    const newWorksheet = xlsx.utils.json_to_sheet(data);

    // Replace the old sheet with the new one in the workbook
    workbook.Sheets[sheetName] = newWorksheet;

    // Replace the old sheet with the new one in the workbook
    workbook.Sheets[sheetName] = newWorksheet;
    // Save the updated workbook
    const outputFilePath = `./uploads/results.xlsx`;
    xlsx.writeFile(workbook, outputFilePath);

    res.status(200).json({
      success: true,
      result: { excelUrl: `/uploads/results.xlsx` },
    });
  } catch (error) {
    console.error("Error processing file upload:", error);
    res.status(500).json({ error: "Internal server error" });
  }
});

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
