// Import necessary modules and libraries
const { google } = require('googleapis');

// Set the spreadsheet ID and range for student data
const spreadsheetId = '1g5w_VVbC6icNyMB2KaursHvcVnSGkkn1xFwRB_vAg2o';
const range = 'A2:H';

// Define constants for the semester
const totalClasses = 60;
const allowedAbsencePercentage = 0.25;

// Load Google Sheets API credentials
const credentials = require('./credentials.json');
const client = new google.auth.JWT(
  credentials.client_email,
  null,
  credentials.private_key,
  ['https://www.googleapis.com/auth/spreadsheets'],
);

// Main function to process the spreadsheet
async function processSpreadsheet() {
  const sheets = google.sheets({ version: 'v4', auth: client });

  try {
    // Fetch data from the spreadsheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = response.data.values;

    // Check if there is data in the spreadsheet
    if (rows.length > 3) {
      // Iterate through the rows, starting from row 4
      for (let i = 2; i < rows.length + 2; i++) {
        const row = rows[i];

        // Check if there is enough data in the row and it does not contain certain keywords
        if (row && row.length >= 6 && !row.includes('Status') && !row.includes('Final Grade')) {
          const registration = row[0];
          const student = row[1];
          const absences = parseFloat(row[2]);
          const grade1 = parseFloat(row[3]);
          const grade2 = parseFloat(row[4]);
          const grade3 = parseFloat(row[5]);

          // Calculate the average and determine the status and NAF
          const average = (grade1 + grade2 + grade3) / 3;
          const { status, naf } = calculateStatusAndNAF(average, absences);

          // Update the spreadsheet with the results, starting from row 4
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `G${i + 2}`, // Correction to start from row 4
            valueInputOption: 'RAW',
            resource: {
              values: [[status]],
            },
          });

          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `H${i + 2}`, // Correction to start from row 4
            valueInputOption: 'RAW',
            resource: {
              values: [[naf]],
            },
          });

          // Log the results for each student
          console.log(`Student ${student} - Status: ${status}, NAF: ${naf}`);
        }
      }
    } else {
      console.log('No data found in the spreadsheet.');
    }
  } catch (error) {
    console.error('Error processing the spreadsheet:', error.message);
  }
}

// Function to calculate student status and NAF
function calculateStatusAndNAF(average, absences) {
  let status, naf;

  if (absences >= totalClasses * allowedAbsencePercentage) {
    status = 'Failed due to Absence';
    naf = 0;
  } else if (average < 50) {
    status = 'Failed due to Grade';
    naf = 0;
  } else if (average < 70) {
    status = 'Final Exam';
    naf = Math.ceil(100 - average);
  } else {
    status = 'Passed';
    naf = 0;
  }

  return { status, naf };
}


processSpreadsheet();

// Command to execute in the terminal: node app.js
