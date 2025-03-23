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

// 42054 is 19-02-2015
console.log(excelDateToJSDate(42054));

// 45568 is 03-10-24
console.log(excelDateToJSDate(45568));

// 45565 is 30-9-24
console.log(excelDateToJSDate(45565));
