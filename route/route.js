const express = require('express')
const router = express.Router()
const multer = require('multer');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');

const upload = multer({ dest: 'uploads/' }); // Destination folder to store the uploaded files
// Variables to store the file paths temporarily
let uploadedExcelPath;
let createdWorksheetPath;
router.post('/upload', upload.fields([
    { name: 'excel', maxCount: 1 },
    { name: 'files', maxCount: 1 }
]), (req, res) => {
    try {
        // Handle the uploaded files
        const files = req.files;
        const excelFile = req.files.excel[0];
        const rptFile = req.files.files[0];
        console.log(excelFile);



        // // Load the workbook
        const workbook = XLSX.readFile(excelFile.path);

        // Access the sheets
        const sheet1Name = workbook.SheetNames[0]; // Get the name of the first sheet
        const sheet1 = workbook.Sheets[sheet1Name];



        // Read data from the text file
        const textData = fs.readFileSync(rptFile.path, 'utf-8');
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
        if (workbook.SheetNames.indexOf('PICK PLACE') !== -1) {
            workbook.Sheets['PICK PLACE'] = sheet2;
        } else {
            XLSX.utils.book_append_sheet(workbook, sheet2, 'PICK PLACE');
        }

        // Save the updated workbook
        const createdWorksheet = excelFile.originalname;
        XLSX.writeFile(workbook, createdWorksheet);

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

                                    XLSX.utils.sheet_add_json(sheet2, data2, { header: ['J'], origin: 'J5', skipHeader: true });
                                }
                                // Find the column index for the current reference
                                // Update the corresponding cell in sheet2 for the current reference
                                sheet2[XLSX.utils.encode_cell({ r: rowIndexInSheet2, c: columnIndexInSheet2 + 9 })] = { t: 's', v: part, w: part };
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
            XLSX.utils.sheet_add_json(sheet1, data1, { header: ['M', 'N'], origin: 'K3', skipHeader: true });
            // Add "top" and "bottom" columns to Sheet1 for the current row
            sheet1[`M${rowIndexInSheet1 + 1}`] = { t: 's', v: topArray.join(',').trim(), w: topArray.join(',').trim() };
            sheet1[`N${rowIndexInSheet1 + 1}`] = { t: 's', v: bottomArray.join(',').trim(), w: bottomArray.join(',').trim() };
        }


        // Save the updated workbook
        // const outputPath =  path.join(__dirname, excelFile.originalname); 
        // const outputPath = path.resolve(excelFile.originalname);

        const outputPath = (excelFile.originalname);
        console.log(outputPath)
        XLSX.writeFile(workbook, outputPath);
 // Save the paths to be used later in the GET endpoint
 uploadedExcelPath = excelFile.path;
createdWorksheetPath = outputPath;
        res.status(200).json({message:'File Sent Successfully!'})
    } catch (err) {
        console.log(err);
        res.status(500).json({ error: 'Internal Server Error' });
    }
});
// Separate GET endpoint to download the created worksheet
router.get('/download', (req, res) => {
    if (createdWorksheetPath) {
      res.download(createdWorksheetPath, createdWorksheetPath, (err) => {
        if (err) {
          console.error('Error downloading file:', err);
        } else {
          // Optionally, you can remove the temporary files after sending the response
          fs.unlinkSync(uploadedExcelPath);
          fs.unlinkSync(createdWorksheetPath);
        }
      });
    } else {
      res.status(404).json({ error: 'File not found' });
    }
  });
  
module.exports = router;