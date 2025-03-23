const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to convert Excel serial date to JavaScript Date object
function excelDateToJSDate(serial) {
  // Excel's epoch starts on January 1, 1900
  // 1900 is not a leap year in reality, but Excel treats it as one
  // So we need to adjust for this Excel bug
  const excelEpoch = new Date(1899, 11, 30);
  const offsetDays = serial;
  const offsetMilliseconds = offsetDays * 24 * 60 * 60 * 1000;
  const jsDate = new Date(excelEpoch.getTime() + offsetMilliseconds);
  // Format the date as needed, e.g., "DD-MMM-YY"
  const formattedDate = jsDate.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: '2-digit',
  });

  return formattedDate;
}

/**
 * Analyzes Nifty 50 data to identify trading signals and calculate statistics
 * @param {string} inputFilePath - Path to the Excel file
 */
function analyzeNiftyData(inputFilePath, sheetName, outputFilePath) {
  try {
    // Read the Excel file
    const workbook = xlsx.readFile(inputFilePath);
    const worksheet = workbook.Sheets[sheetName];

    // Convert to JSON
    const data = xlsx.utils.sheet_to_json(worksheet);

    // Sort data by date (oldest to newest)
    // data.sort((a, b) => {
    //   const dateA = new Date(a.Date.split('-').reverse().join('-'));
    //   const dateB = new Date(b.Date.split('-').reverse().join('-'));
    //   return dateA - dateB;
    // });

    // Filter out rows with null or 0 SMA values
    const filteredData = data.filter(
      (row) => row.SMA_20 && row.SMA_50 && row.SMA_20 !== 0 && row.SMA_50 !== 0
    );

    // Initialize arrays to store trade information
    const trades = [];
    const bullishSignalDates = [];
    const bearishSignalDates = [];

    // Analyze data for trading signals
    for (let i = 3; i < filteredData.length - 1; i++) {
      // Check for bullish setup
      const bullishSetup = checkBullishSetup(filteredData, i);
      if (bullishSetup) {
        bullishSignalDates.push(filteredData[i].Date);

        // Simulate trade
        const trade = simulateTrade(filteredData, i, true);
        if (trade) {
          trades.push(trade);
        }
      }

      // Check for bearish setup
      //   const bearishSetup = checkBearishSetup(filteredData, i);
      //   if (bearishSetup) {
      //     bearishSignalDates.push(filteredData[i].Date);

      //     // Simulate trade
      //     const trade = simulateTrade(filteredData, i, false);
      //     if (trade) {
      //       trades.push(trade);
      //     }
      //   }
    }

    // Calculate statistics
    const stats = calculateStatistics(trades);

    // Generate output
    const output = generateOutput(trades, stats, bullishSignalDates, bearishSignalDates);

    // Write to file
    const outputPath = outputFilePath;
    fs.writeFileSync(outputPath, output);

    console.log(`Analysis complete. Results saved to: ${outputPath}`);
  } catch (error) {
    console.error('Error analyzing Nifty data:', error);
  }
}

/**
 * Checks if a bullish setup is present at the given index
 * @param {Array} data - The filtered data array
 * @param {number} index - Current index to check
 * @returns {boolean} - Whether a bullish setup is present
 */
function checkBullishSetup(data, index) {
  // Check if SMA 20 is rising for past 3 days
  const sma20Rising =
    data[index].SMA_20 > data[index - 1].SMA_20 &&
    data[index - 1].SMA_20 > data[index - 2].SMA_20 &&
    data[index - 2].SMA_20 > data[index - 3].SMA_20;

  // Check if SMA 50 is rising for past 3 days
  const sma50Rising =
    data[index].SMA_50 > data[index - 1].SMA_50 &&
    data[index - 1].SMA_50 > data[index - 2].SMA_50 &&
    data[index - 2].SMA_50 > data[index - 3].SMA_50;

  // Check if SMA 20 is above SMA 50
  const sma20AboveSma50 = data[index].SMA_20 > data[index].SMA_50;

  // Check if price opened below SMA 20 and closed above it OR opened above SMA 20 and closed above opening price
  const priceCondition =
    (data[index].Open < data[index].SMA_20 && data[index].Close > data[index].SMA_20) ||
    (data[index].Open > data[index].SMA_20 && data[index].Close > data[index].Open);

  // Check if close price is within +1.5% of SMA 20
  const priceDiffPercent =
    Math.abs((data[index].Close - data[index].SMA_20) / data[index].SMA_20) * 100;
  const closeNearSma20 = priceDiffPercent <= 1.5;

  return sma20Rising && sma50Rising && sma20AboveSma50 && priceCondition && closeNearSma20;
}

/**
 * Checks if a bearish setup is present at the given index
 * @param {Array} data - The filtered data array
 * @param {number} index - Current index to check
 * @returns {boolean} - Whether a bearish setup is present
 */
function checkBearishSetup(data, index) {
  // Check if SMA 20 is falling for past 3 days
  const sma20Falling =
    data[index].SMA_20 < data[index - 1].SMA_20 &&
    data[index - 1].SMA_20 < data[index - 2].SMA_20 &&
    data[index - 2].SMA_20 < data[index - 3].SMA_20;

  // Check if SMA 50 is falling for past 3 days
  const sma50Falling =
    data[index].SMA_50 < data[index - 1].SMA_50 &&
    data[index - 1].SMA_50 < data[index - 2].SMA_50 &&
    data[index - 2].SMA_50 < data[index - 3].SMA_50;

  // Check if SMA 20 is below SMA 50
  const sma20BelowSma50 = data[index].SMA_20 < data[index].SMA_50;

  // Check if price opened above SMA 20 and closed below it OR opened below SMA 20 and closed below opening price
  const priceCondition =
    (data[index].Open > data[index].SMA_20 && data[index].Close < data[index].SMA_20) ||
    (data[index].Open < data[index].SMA_20 && data[index].Close < data[index].Open);

  // Check if close price is within 1-1.5% of SMA 20
  const priceDiffPercent =
    Math.abs((data[index].Close - data[index].SMA_20) / data[index].SMA_20) * 100;
  const closeNearSma20 = priceDiffPercent >= 1 && priceDiffPercent <= 1.5;

  return sma20Falling && sma50Falling && sma20BelowSma50 && priceCondition && closeNearSma20;
}

/**
 * Simulates a trade based on the signal at the given index
 * @param {Array} data - The filtered data array
 * @param {number} signalIndex - Index of the signal
 * @param {boolean} isBullish - Whether the signal is bullish
 * @returns {Object|null} - Trade information or null if trade couldn't be completed
 */
function simulateTrade(data, signalIndex, isBullish) {
  // Ensure there's enough data after the signal
  if (signalIndex + 1 >= data.length) {
    return null;
  }

  const signalDate = excelDateToJSDate(data[signalIndex].Date);
  const entryDate = data[signalIndex + 1].Date;
  const entryPrice = data[signalIndex + 1].Open;
  const initialStopLoss = data[signalIndex].SMA_20;

  if (
    entryPrice < data[signalIndex].SMA_20 ||
    entryPrice < data[signalIndex].Open ||
    entryPrice < data[signalIndex].Close
  ) {
    return null;
  }

  // Calculate target based on risk (2x risk for profit factor)
  const risk = Math.abs(entryPrice - initialStopLoss);
  const targetPrice = isBullish ? entryPrice + risk * 2 : entryPrice - risk * 2;

  let currentStopLoss = initialStopLoss;
  let exitDate = null;
  let exitPrice = null;
  let tradeDays = 0;
  let tradeStatus = null;

  // Simulate the trade day by day
  for (let i = signalIndex + 1; i < data.length; i++) {
    tradeDays++;

    // Update trailing stop loss for bullish trades (only moves up)
    if (isBullish && data[i].SMA_20 > currentStopLoss) {
      currentStopLoss = data[i].SMA_20;
    }
    // Update trailing stop loss for bearish trades (only moves down)
    else if (!isBullish && data[i].SMA_20 < currentStopLoss) {
      currentStopLoss = data[i].SMA_20;
    }

    // Check if stop loss was hit
    if (
      (isBullish && data[i].Close <= currentStopLoss) ||
      (!isBullish && data[i].Close >= currentStopLoss)
    ) {
      exitDate = data[i + 1].Date;
      exitPrice = data[i + 1].Open;
      if (isBullish) {
        tradeStatus = entryPrice < exitPrice ? 'Profit' : 'Loss';
      } else {
        tradeStatus = entryPrice > exitPrice ? 'Profit' : 'Loss';
      }
      break;
    }

    // Check if target was hit
    if (
      (isBullish && data[i].Close >= targetPrice) ||
      (!isBullish && data[i].Close <= targetPrice)
    ) {
      exitDate = data[i + 1].Date;
      exitPrice = data[i + 1].Open;
      if (isBullish) {
        tradeStatus = entryPrice < exitPrice ? 'Profit' : 'Loss';
      } else {
        tradeStatus = entryPrice > exitPrice ? 'Profit' : 'Loss';
      }
      break;
    }

    // If we've reached the end of the data
    if (i === data.length - 1) {
      exitDate = data[i].Date;
      exitPrice = data[i].Close;
      tradeStatus = isBullish
        ? exitPrice > entryPrice
          ? 'Profit'
          : 'Loss'
        : exitPrice < entryPrice
        ? 'Profit'
        : 'Loss';
    }
  }

  // If no exit was found
  if (!exitDate) {
    return null;
  }

  // Calculate P/L
  const pl = isBullish ? exitPrice - entryPrice : entryPrice - exitPrice;
  const plPercentage = (pl / entryPrice) * 100;

  return {
    signalDate,
    type: isBullish ? 'Bullish' : 'Bearish',
    entryDate,
    entryPrice,
    initialStopLoss,
    targetPrice,
    exitDate,
    exitPrice,
    tradeDays,
    status: tradeStatus,
    pl,
    plPercentage,
  };
}

/**
 * Calculates statistics from the trades
 * @param {Array} trades - Array of trade objects
 * @returns {Object} - Statistics
 */
function calculateStatistics(trades) {
  const totalTrades = trades.length;
  const profitableTrades = trades.filter((t) => t.status === 'Profit');
  const lossMakingTrades = trades.filter((t) => t.status === 'Loss');

  const percentProfitable = (profitableTrades.length / totalTrades) * 100;
  const percentLossMaking = (lossMakingTrades.length / totalTrades) * 100;

  const avgProfitPercent =
    profitableTrades.length > 0
      ? profitableTrades.reduce((sum, t) => sum + t.plPercentage, 0) / profitableTrades.length
      : 0;

  const avgLossPercent =
    lossMakingTrades.length > 0
      ? lossMakingTrades.reduce((sum, t) => sum + t.plPercentage, 0) / lossMakingTrades.length
      : 0;

  const avgTradeDays = trades.reduce((sum, t) => sum + t.tradeDays, 0) / totalTrades;

  const avgProfitableTradeDays =
    profitableTrades.length > 0
      ? profitableTrades.reduce((sum, t) => sum + t.tradeDays, 0) / profitableTrades.length
      : 0;

  const avgLossMakingTradeDays =
    lossMakingTrades.length > 0
      ? lossMakingTrades.reduce((sum, t) => sum + t.tradeDays, 0) / lossMakingTrades.length
      : 0;

  return {
    totalTrades,
    profitableTrades: profitableTrades.length,
    lossMakingTrades: lossMakingTrades.length,
    percentProfitable,
    percentLossMaking,
    avgProfitPercent,
    avgLossPercent,
    avgTradeDays,
    avgProfitableTradeDays,
    avgLossMakingTradeDays,
  };
}

/**
 * Generates the output text
 * @param {Array} trades - Array of trade objects
 * @param {Object} stats - Statistics object
 * @param {Array} bullishDates - Array of bullish signal dates
 * @param {Array} bearishDates - Array of bearish signal dates
 * @returns {string} - Formatted output text
 */
function generateOutput(trades, stats, bullishDates, bearishDates) {
  let output = '* NIFTY 50 BULLISH SMA TRADES STATISTICAL ANALYSIS RESULTS *\n\n';

  // Add bullish signal dates
  //   output += '=== BULLISH SIGNAL DATES ===\n';
  //   bullishDates.forEach((date) => {
  //     output += date + '\n';
  //   });
  //   output += '\n';

  //   // Add bearish signal dates
  //   output += '=== BEARISH SIGNAL DATES ===\n';
  //   bearishDates.forEach((date) => {
  //     output += date + '\n';
  //   });
  //   output += '\n';

  // Add overall statistics
  output += '=== OVERALL STATISTICS ===\n\n';
  output += `Total Trades: ${stats.totalTrades}\n`;
  output += `Profitable Trades: ${stats.profitableTrades}\n`;
  output += `Loss-Making Trades: ${stats.lossMakingTrades}\n`;
  output += `Percentage of Profitable Trades: ${stats.percentProfitable.toFixed(2)}%\n`;
  output += `Percentage of Loss-Making Trades: ${stats.percentLossMaking.toFixed(2)}%\n`;
  output += `Average Profit Percentage (Profitable Trades): ${stats.avgProfitPercent.toFixed(
    2
  )}%\n`;
  output += `Average Loss Percentage (Loss-Making Trades): ${stats.avgLossPercent.toFixed(2)}%\n`;
  output += `Average Trade Duration (All Trades): ${stats.avgTradeDays.toFixed(1)} days\n`;
  output += `Average Duration of Profitable Trades: ${stats.avgProfitableTradeDays.toFixed(
    1
  )} days\n`;
  output += `Average Duration of Loss-Making Trades: ${stats.avgLossMakingTradeDays.toFixed(
    1
  )} days\n`;

  // Add individual trade details
  output += '\n\n=== INDIVIDUAL TRADE DETAILS ===\n\n';
  for (let index = trades.length - 1; index >= 0; index--) {
    const trade = trades[index];
    output += `TRADE #${index + 1}:\n`;
    output += `SIGNAL DATE: ${trade.signalDate}\n`;
    output += `ENTRY DATE: ${excelDateToJSDate(trade.entryDate)}\n`;
    output += `ENTRY PRICE: ${trade.entryPrice.toFixed(2)}\n`;
    output += `INITIAL STOP LOSS: ${trade.initialStopLoss.toFixed(2)}\n`;
    output += `TARGET PRICE: ${trade.targetPrice.toFixed(2)}\n`;
    output += `EXIT DATE: ${excelDateToJSDate(trade.exitDate)}\n`;
    output += `EXIT PRICE: ${trade.exitPrice.toFixed(2)}\n`;
    output += `TRADE TIME IN DAYS: ${trade.tradeDays}\n`;
    output += `STATUS OF TRADE: ${trade.status}\n`;
    output += `P/L: ${trade.pl.toFixed(2)}\n`;
    output += `P/L PERCENTAGE: ${trade.plPercentage.toFixed(2)}%\n\n`;
  }

  //   trades.forEach((trade, index) => {
  //     output += `TRADE #${index + 1} (${trade.type}):\n`;
  //     output += `SIGNAL DATE: ${trade.signalDate}\n`;
  //     output += `ENTRY DATE: ${excelDateToJSDate(trade.entryDate)}\n`;
  //     output += `ENTRY PRICE: ${trade.entryPrice.toFixed(2)}\n`;
  //     output += `INITIAL STOP LOSS: ${trade.initialStopLoss.toFixed(2)}\n`;
  //     output += `TARGET PRICE: ${trade.targetPrice.toFixed(2)}\n`;
  //     output += `EXIT DATE: ${excelDateToJSDate(trade.exitDate)}\n`;
  //     output += `EXIT PRICE: ${trade.exitPrice.toFixed(2)}\n`;
  //     output += `TRADE TIME IN DAYS: ${trade.tradeDays}\n`;
  //     output += `STATUS OF TRADE: ${trade.status}\n`;
  //     output += `P/L: ${trade.pl.toFixed(2)}\n`;
  //     output += `P/L PERCENTAGE: ${trade.plPercentage.toFixed(2)}%\n\n`;
  //   });

  return output;
}

// Get the file path from command-line arguments
const inputFilePath = '../files/NIFTY_50_Feb_2015_to_March_2025.xlsx';
const sheetName = 'NIFTY50_with_SMA';
const outputFilePath = '../files/BULLISH_SMA_TRADE_STATS.txt';

// Run the analysis
analyzeNiftyData(inputFilePath, sheetName, outputFilePath);
