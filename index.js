
const XLSX = require('xlsx');
const fs = require('fs');
// // Load the workbook
const workbook = XLSX.readFile('./VCL060-02-RFSOC_BP_2023AUG23_BOM.xlsx');

// Access the sheets
const sheet1Name = workbook.SheetNames[0]; // Get the name of the first sheet
const sheet1 = workbook.Sheets[sheet1Name];



// Read data from the text file
const textData = fs.readFileSync('./pcp_rep.rpt', 'utf-8');
const lines = textData.split('\n');
const headers = lines[0].trim().split(',');

// Create sheet2Data with headers
const sheet2Data = [headers];

// Parse text data and create sheet2Data
for (let i = 1; i < lines.length; i++) {
  const values = lines[i].trim().split(',');
  sheet2Data.push(values);
}

// Calculate the range for sheet2 based on the number of rows and columns
const range = {
  s: { c: 0, r: 0 },
  e: { c: headers.length - 1, r: sheet2Data.length - 1 },
};

// Create a new sheet (Sheet2) in the workbook
const sheet2 = XLSX.utils.aoa_to_sheet(sheet2Data, { cellDates: true, range });

// Append or replace 'Sheet2'
if (workbook.SheetNames.indexOf('Sheet2') !== -1) {
  workbook.Sheets['Sheet2'] = sheet2;
} else {
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2');
}

// Save the updated workbook
XLSX.writeFile(workbook, './VCL060-02-RFSOC_BP_2023AUG23_BOM.xlsx');

// Iterate over rows in Sheet1
const sheet1Data = XLSX.utils.sheet_to_json(sheet1, { header: 1 });

// Iterate over rows in Sheet1
for (let rowIndexInSheet1 = 0; rowIndexInSheet1 < sheet1Data.length; rowIndexInSheet1++) {
  const row1 = sheet1Data[rowIndexInSheet1];

  // Initialize top and bottom arrays for each row
  let topArray = [];
  let bottomArray = [];
  let columnDataArray = []; 

  if (Array.isArray(row1) && row1.length >= 3) {
    const refs = row1[2].split(',').map(ref => ref.trim()); 
    const part = row1[6]

    if (Array.isArray(refs)) {
      // Iterate over references
      refs.forEach((ref) => {
        // Find the corresponding row in Sheet2 based on Ref
        const rowIndexInSheet2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 }).findIndex((r) => r[0] === ref);

        if (rowIndexInSheet2 !== -1) {
          // Update "top" or "bottom" based on "s_m"
          const s_mValue = sheet2[XLSX.utils.encode_cell({ r: rowIndexInSheet2, c: 8 })];
          const refValue = sheet2[XLSX.utils.encode_cell({ r: rowIndexInSheet2, c: 0 })];
          if (s_mValue && s_mValue.v === 'NO') {
            topArray.push(ref);
          } else if (s_mValue && s_mValue.v === 'YES') {
            bottomArray.push(ref);
          }
          if (refValue && refValue.v === ref) {
            const columnIndexInSheet2 = headers.indexOf('Part Number') + 1;
            if (columnIndexInSheet2 === 0) {
              const data2 = [
                { J: 'Part Number' }
              ];

              XLSX.utils.sheet_add_json(sheet2, data2, { header: ['J'], origin: 'J1', skipHeader: true });
            }
            // Find the column index for the current reference
            // Update the corresponding cell in sheet2 for the current reference
            sheet2[XLSX.utils.encode_cell({ r: rowIndexInSheet2, c: columnIndexInSheet2 +9 })] = { t: 's', v: part, w: part };
          } else {
            console.log('not worked')

          }
        }
      });
    } else {
      console.log('Skipping row1 - refs is not an array:', refs);
    }
  } else {
    console.log('Skipping row1 ', row1);
  }

  const data1 = [
    { D: 'TOP', E: 'BOTTOM' } 
  ];
  // Add data to the worksheet, starting from cell M1
  XLSX.utils.sheet_add_json(sheet1, data1, { header: ['M', 'N'], origin: 'K1', skipHeader: true });
  // Add "top" and "bottom" columns to Sheet1 for the current row
  sheet1[`M${rowIndexInSheet1 + 1}`] = { t: 's', v: topArray.join(',').trim(), w: topArray.join(',').trim() };
  sheet1[`N${rowIndexInSheet1 + 1}`] = { t: 's', v: bottomArray.join(',').trim(), w: bottomArray.join(',').trim() };
}

// Save the updated workbook
XLSX.writeFile(workbook, './file_updated_with_top_bottom_column.xlsx');
