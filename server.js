const express = require("express");
const multer = require("multer");
const unzipper = require("unzipper");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
app.use(cors());
const PORT = 5000;

const UPLOAD_DIR = "uploads";
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const safeName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, "_");
    cb(null, Date.now() + "-" + safeName);
  },
});
const upload = multer({ storage, limits: { fileSize: 200 * 1024 * 1024 } });

function cleanOldFiles() {
  if (fs.existsSync(UPLOAD_DIR)) {
    fs.readdirSync(UPLOAD_DIR).forEach(file => {
      const filePath = path.join(UPLOAD_DIR, file);
      if (file.startsWith("merged-") || file.startsWith("extracted-")) {
        fs.rmSync(filePath, { recursive: true, force: true });
      }
    });
  }
}

async function processZip(file) {
  const zipPath = file.path;
  const extractDir = path.join(
    UPLOAD_DIR,
    `extracted-${Date.now()}-${Math.floor(Math.random() * 1000)}`
  );
  fs.mkdirSync(extractDir);

  await fs.createReadStream(zipPath)
    .pipe(unzipper.Extract({ path: extractDir }))
    .promise();

  const excelFiles = fs
    .readdirSync(extractDir)
    .filter(f => f.endsWith(".xlsx") || f.endsWith(".xls"));

  let allData = [];

  for (const excel of excelFiles) {
    const workbook = XLSX.readFile(path.join(extractDir, excel));

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet, { defval: "", header: 1 });

      if (sheetData.length === 0) continue;

      // Include header only once
      if (allData.length === 0) {
        allData = allData.concat(sheetData);
      } else {
        allData = allData.concat(sheetData.slice(1));
      }
    }
  }

  fs.rmSync(extractDir, { recursive: true, force: true });
  fs.unlinkSync(zipPath);

  return allData;
}

async function processInBatches(files, batchSize = 5) {
  let mergedData = [];
  for (let i = 0; i < files.length; i += batchSize) {
    const batch = files.slice(i, i + batchSize);
    const results = await Promise.allSettled(batch.map(file => processZip(file)));
    for (const r of results) {
      if (r.status === "fulfilled") {
        mergedData = mergedData.concat(r.value);
      } else {
        console.error("Error processing file:", r.reason);
      }
    }
  }
  return mergedData;
}

const MAX_CELL_LENGTH = 32767;
function truncateCellValue(value) {
  if (typeof value === "string" && value.length > MAX_CELL_LENGTH) {
    return value.slice(0, MAX_CELL_LENGTH);
  }
  return value;
}

app.post("/upload-zip", upload.any(), async (req, res) => {
  try {
    cleanOldFiles();

    if (!req.files || req.files.length === 0)
      return res.status(400).send("No files uploaded");

    const mergedData = await processInBatches(req.files, 5);

    if (mergedData.length === 0) {
      return res.status(400).send("No data found in uploaded files");
    }

    // Truncate cells to avoid Excel limits
    const truncatedData = mergedData.map(row => row.map(cell => truncateCellValue(cell)));

    // Create workbook & worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(truncatedData);

    XLSX.utils.book_append_sheet(wb, ws, "Merged Data");

    const mergedFilePath = path.join(UPLOAD_DIR, `merged-${Date.now()}.xlsx`);
    XLSX.writeFile(wb, mergedFilePath);

    res.download(mergedFilePath, err => {
      if (!err) fs.unlinkSync(mergedFilePath);
    });
  } catch (error) {
    console.error(error);
    res.status(500).send("Server error: " + error.message);
  }
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
