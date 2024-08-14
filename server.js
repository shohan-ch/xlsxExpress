const express = require("express");
const XLSX = require("xlsx");
const path = require("path");

const app = express();

app.get("/xlsx", (req, res) => {
  const filePath = path.join(__dirname, "public/toilet.xlsx");
  const workbook = XLSX.readFile(filePath);
  const sheetNames = workbook.SheetNames;
  const sheet = workbook.Sheets[sheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const columns = data[0];
  const uniqueValues = {};

  columns.forEach((col, index) => {
    const values = data.slice(1).map((row) => row[index]);
    uniqueValues[col] = [...new Set(values)];
  });

  const newWorkbook = XLSX.utils.book_new();

  const newSheetData = [columns];
  const maxLength = Math.max(
    ...Object.values(uniqueValues).map((arr) => arr.length)
  );

  for (let i = 0; i < maxLength; i++) {
    const row = columns.map((col) => uniqueValues[col][i] || "");
    newSheetData.push(row);
  }

  const newSheet = XLSX.utils.aoa_to_sheet(newSheetData);
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "UniqueValues");

  const outputFilePath = path.join(__dirname, "public/unique_values.xlsx");
  XLSX.writeFile(newWorkbook, outputFilePath);

  res.download(outputFilePath, "unique_values.xlsx", (err) => {
    if (err) {
      res.status(500).send("Error downloading file.");
    }
  });

  //   Read data raher than download

  //   const workbook = XLSX.readFile(filePath);
  //   const sheetNames = workbook.SheetNames;
  //   const sheet = workbook.Sheets[sheetNames[0]]; // Assuming you are interested in the first sheet
  //   const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  //   const columns = data[0];
  //   const uniqueValues = {};

  //   columns.forEach((col, index) => {
  //     const values = data.slice(1).map((row) => row[index]);
  //     uniqueValues[col] = [...new Set(values)];
  //   });
  //   res.json(uniqueValues);
});

app.listen(4000, () => console.log("app listen at port" + 4000));
