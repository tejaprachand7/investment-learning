const xlsx = require('xlsx');
const fs = require('fs');

/**
 * Reads NIFTY 50 data from an Excel file
 * @param {string} filePath - Path to the Excel file
 * @returns {Array} Array of data objects with Date, Open, High, Low, Close
 */
function readNiftyDataFromExcel(filePath) {
  try {
    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert to JSON with headers
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract headers and data rows
    const headers = data[0];
    const rows = data.slice(1);

    // Map rows to objects with proper date parsing
    return rows.map((row) => {
      const entry = {};
      headers.forEach((header, index) => {
        if (header === 'Date') {
          // Handle date parsing - assuming DD-MMM-YY format
          entry[header] = row[index];
        } else {
          // Convert numeric values to numbers with 2 decimal precision
          entry[header] = parseFloat(parseFloat(row[index]).toFixed(2));
        }
      });
      return entry;
    });
  } catch (error) {
    console.error('Error reading Excel file:', error);
    throw error;
  }
}

/**
 * Calculates Simple Moving Average for a specified period
 * @param {Array} data - Array of data objects with Close prices
 * @param {number} period - Period for SMA calculation (e.g., 20, 50)
 * @returns {Array} Array of SMA values
 */
function calculateSMA(data, period) {
  const smaValues = [];

  // For each data point
  for (let i = 0; i < data.length; i++) {
    if (i < period - 1) {
      // Not enough data points yet for the period
      smaValues.push(null);
    } else {
      // Calculate sum of closing prices for the period
      let sum = 0;
      for (let j = i - period + 1; j <= i; j++) {
        sum += data[j].Close;
      }

      // Calculate average and round to 2 decimal places
      const sma = parseFloat((sum / period).toFixed(2));
      smaValues.push(sma);
    }
  }

  return smaValues;
}

/**
 * Adds SMA columns to the data and writes to a new Excel file
 * @param {Array} data - Original data array
 * @param {Array} sma20Values - SMA 20 values
 * @param {Array} sma50Values - SMA 50 values
 * @param {string} outputPath - Path for the output Excel file
 */

function writeDataWithSMAToExcel(data, sma20Values, sma50Values, inputFilePath) {
  try {
    // Create a new array with original data plus SMA values
    const enhancedData = data.map((entry, index) => {
      return {
        ...entry,
        SMA_20: sma20Values[index],
        SMA_50: sma50Values[index],
      };
    });

    // Read the existing workbook
    let workbook;
    try {
      workbook = xlsx.readFile(inputFilePath);
    } catch (error) {
      // If file doesn't exist or can't be read, create a new workbook
      workbook = xlsx.utils.book_new();
    }

    // Convert enhanced data to worksheet
    const worksheet = xlsx.utils.json_to_sheet(enhancedData);

    // Add the new worksheet to the existing workbook (or replace if it already exists)
    const sheetName = 'NIFTY50_with_SMA';
    if (workbook.SheetNames.includes(sheetName)) {
      workbook.Sheets[sheetName] = worksheet;
    } else {
      xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    // Write the updated workbook back to the same file
    xlsx.writeFile(workbook, inputFilePath);

    console.log(`Data with SMA values written to sheet '${sheetName}' in ${inputFilePath}`);
  } catch (error) {
    console.error('Error writing Excel file:', error);
    throw error;
  }
}

/**
 * Main function to process NIFTY 50 data
 * @param {string} inputFilePath - Path to the input Excel file
 * @param {string} outputFilePath - Path for the output Excel file
 */
function processNiftyData(inputFilePath, outputFilePath) {
  try {
    // Read data from Excel
    const data = readNiftyDataFromExcel(inputFilePath);

    // Sort data by date (oldest to newest) if needed
    // Assuming data is already sorted as mentioned

    // Calculate SMA values
    const sma20Values = calculateSMA(data, 20);
    const sma50Values = calculateSMA(data, 50);

    // Write enhanced data to new Excel file
    writeDataWithSMAToExcel(data, sma20Values, sma50Values, outputFilePath);

    console.log('NIFTY 50 data processing completed successfully');
  } catch (error) {
    console.error('Error processing NIFTY 50 data:', error);
  }
}

// Usage example
const inputFile = '../files/NIFTY_50_Feb_2015_to_March_2025.xlsx';

const outputFile = '../files/NIFTY_50_Feb_2015_to_March_2025.xlsx';
processNiftyData(inputFile, outputFile);
