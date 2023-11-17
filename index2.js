
// Read the existing workbook
const existingWorkbook = XLSX.readFile(filePath);

// Get the Sheet1 from the existing workbook
const ws = existingWorkbook.Sheets['Sheet1'];

// Add data to columns M and N
const data = [
  { M: 'TOP', N: 'BOTTOM' }
  // Add more rows as needed
];

// Add data to the worksheet, starting from cell M1
XLSX.utils.sheet_add_json(ws, data, { header: ['M', 'N'], origin: 'M1', skipHeader: true });

// Save the changes to the existing workbook
XLSX.writeFile(existingWorkbook, filePath);

console.log('Changes saved successfully!');
