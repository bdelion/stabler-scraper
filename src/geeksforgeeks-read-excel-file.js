#!/usr/bin/env node

"use strict";

// Requiring the module
const XLSX = require("xlsx");

const getData = (excelFilePath, sheetName, firstRow) => {
  // Reading our test file
  const wb = XLSX.readFile(excelFilePath);
  // Select sheet
  const ws = wb.Sheets[sheetName];
  // Read data in string format, begin at a specific raw
  const json = XLSX.utils.sheet_to_json(ws, {
    blankRows: false,
    range: firstRow,
    raw: false,
  });
  return json;
};

const jsonResult = getData("assets/InputData.xlsx", "Suivi Conso New", 2);
// console.log(jsonResult);

let previousDate = "";
const dateArray = [];

jsonResult.forEach((entry) => {
  // Initialize an empty object to store the row data
  const rowData = {};
  if (previousDate !== "") {
    rowData["begin"] = previousDate;
    rowData["end"] = entry.Date;
    dateArray.push(rowData);
  }
  previousDate = entry.Date;
});

console.log(JSON.stringify(dateArray));
