#!/usr/bin/env node

"use strict";

// Requiring the module
const XLSX = require("xlsx");
const axios = require("axios");
const cheerio = require("cheerio");
const moment = require("moment");

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

const formatData = (jsonArray) => {
  let previousDate = "";
  const dateArray = [];

  jsonArray.forEach((entry) => {
    // Initialize an empty object to store the row data
    const rowData = {};
    if (previousDate !== "") {
      rowData["begin"] = previousDate;
      rowData["end"] = entry.Date;
      dateArray.push(rowData);
    }
    previousDate = entry.Date;
  });

  return dateArray;
};

async function performScraping(idCommune, date) {
  const momentDate = moment(date, "DD/MM/YY hh:mm");
  let day = momentDate.date();
  let month = momentDate.month();
  let year = momentDate.year();
  let url = `https://www.meteociel.fr/temps-reel/obs_villes.php?code2=${idCommune}&jour2=${day}&mois2=${month}&annee2=${year}&affint=1`;

  // downloading the target web page by performing an HTTP GET request in Axios
  const axiosResponse = await axios.request({
    method: "GET",
    url: url,
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    },
  });

  // parsing the HTML source of the target web page with Cheerio
  const $ = cheerio.load(axiosResponse.data);

  // Select the table element
  const table = $('table:nth-child(3)[width="100%"]');

  // initializing the data structures that will contain the scraped data
  const dataWeather = [];

  // Iterate over each row of the table using the find and each methods
  table
    .find("tbody")
    .find("tr")
    .each((i, row) => {
      // Initialize an empty object to store the row data
      const rowData = {};

      // Iterate over each cell of the row using the find and each methods
      const dataLine = [];
      $(row)
        // .find('td[align="center"], div[align="center"]')
        .find("td")
        .each((j, cell) => {
          // Add the cell data to the row data object
          dataLine.push($(cell).text());
        });

      if (!(dataLine[0] === "Heurelocale")) {
        rowData["idCommune"] = idCommune;
        rowData["jour"] = momentDate.format("DD/MM/YYYY");
        rowData["heure"] = dataLine[0].replace("h", ":");
        rowData["moment"] = moment(
          rowData["jour"] + " " + rowData["heure"],
          "DD/MM/YYYY hh:mm"
        );
        rowData["temperature"] = dataLine[2].substring(
          0,
          dataLine[2].indexOf(" �C")
        );

        // Add the row data to the table data array
        dataWeather.push(rowData);
      }
    });

  // Return the table data
  return dataWeather;
}

async function getWeatherDataBetween2Dates(idCommune, startDate, endDate) {
  // initializing the data structures that will contain all scraped data
  let datasWeather = [];
  const dateStart = moment(startDate, "DD/MM/YYYY hh:mm");
  const dateEnd = moment(endDate, "DD/MM/YYYY hh:mm");
  const dateEndIteration = dateEnd.clone().add(1, "days");

  let dateIteration = dateStart.clone();
  while (dateIteration < dateEndIteration) {
    datasWeather = datasWeather.concat(
      await performScraping(idCommune, dateIteration)
    );
    dateIteration.add(1, "days");
  }

  // Sort data by date
  datasWeather.sort((a, b) => a.moment - b.moment);
  // Initializing the data structures that will contain filtered data
  let filteredDatasWeather = [];
  // Filter data on date range
  datasWeather.forEach(function (value) {
    if (value["moment"] >= dateStart && value["moment"] <= dateEnd) {
      // Add the row data to the table data array
      filteredDatasWeather.push(value);
    }
  });

  // Sort data by temperature
  filteredDatasWeather.sort((a, b) => a.temperature - b.temperature);
  // Initialize an empty object to store the row data
  const rowData = {};
  rowData["idCommune"] = idCommune;
  rowData["date"] = endDate;
  rowData["moment"] = dateEnd;
  rowData["temperatureMin"] = filteredDatasWeather[0].temperature;
  rowData["temperatureMax"] =
    filteredDatasWeather[filteredDatasWeather.length - 1].temperature;
  return rowData;
}

// initializing the data structures that will contain result data
let weatherDatas = [];

async function main() {
  const jsonResult = getData("assets/InputData.xlsx", "Suivi Conso New", 2);
  const inputDatas = formatData(jsonResult);
  console.log(inputDatas);

  let previousValue = "";

  for (const currentValue of inputDatas) {
    console.log(currentValue);
    weatherDatas.push(
      await getWeatherDataBetween2Dates(
        79049004,
        currentValue.begin,
        currentValue.end
      )
    );

    // if (!(previousValue.trim().length === 0)) {
    //   console.log(previousValue + " --> " + currentValue);
    //   weatherDatas.push(
    //     await getWeatherDataBetween2Dates(79049004, previousValue, currentValue)
    //   );
    // }
    // previousValue = currentValue;
  }

  // Sort data by date
  weatherDatas.sort((a, b) => a.moment - b.moment);
  // console.log(JSON.stringify(weatherDatas));
  return weatherDatas;
}

main()
  .then((result) => {
    console.log(result);
    const worksheet = XLSX.utils.json_to_sheet(result);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Temperatures");
    XLSX.writeFile(workbook, "assets/OutputData.xlsx", { compression: true });
  })
  .catch((err) => console.error(err));
