// use "npm start" to start the program.

const { google } = require('googleapis');
const keys = require('./credentials.json');
const config = require('./config.json');
const spreadsheetId = '1rEXIEFjzMWMMYM956pSFRUGfl4__1v9jmH3MNycE2IM';
const client = new google.auth.JWT(
    keys.client_email,
    null,
    keys.private_key,
    ['https://www.googleapis.com/auth/spreadsheets']
);

client.authorize(function (err, tokens) {
    if (err) {
        return console.log(err);
    } else {
        console.log('Connected!');
        gsrun(client);
    }
});
console.log('loading spreadsheet information');
async function gsrun(cl) {
    const googleService = google.sheets({ version: 'v4', auth: cl })
    const params = {
        spreadsheetId, 
        range: 'A2:F27'
    }
    
    let sheet = await googleService.spreadsheets.values.get(params); // taking information from the spreadsheet
    let sheetArray = sheet.data.values; // taking the values ​​from the spreadsheet
    let p1 = [];
    let p2 = [];
    let p3 = [];
    let average = [];
    let averageStatus = [];
    let naf = [];
    let totalClasses;
    let absencesNumber = [];
    let fillSpreadsheet = [];



    totalClasses = sheetArray[0][0];
    totalClasses = totalClasses.split(':');
    totalClasses = parseInt(totalClasses[1]);




    sheetArray = sheetArray.slice(2); //Removing header of the spreadsheet 
console.log('calculating the values ​​to be filled in the spreadsheet');

    for (let i = 0; i < sheetArray.length; i++) { // filling the variable arrays
        p1[i] = parseInt(sheetArray[i][3]);
        p2[i] = parseInt(sheetArray[i][4]);
        p3[i] = parseInt(sheetArray[i][5]);
        absencesNumber[i] = parseInt(sheetArray[i][2]);
        average[i] = calculateAverage(p1[i], p2[i], p3[i]);
        averageStatus[i] = studentStatus(average[i], absencesNumber[i], totalClasses);
        averageStatus[i] = absenceStatus(totalClasses, absencesNumber[i]);
        naf[i] = calculateNAF(average[i], averageStatus[i]);
        fillSpreadsheet[i] = [averageStatus[i]];
        fillSpreadsheet[i][1] = naf[i];
    }
    console.log(fillSpreadsheet);

console.log('filling in new values ​​in the spreadsheet')
    googleService.spreadsheets.values.update({ // writing data to the spreadsheet
        spreadsheetId: "1rEXIEFjzMWMMYM956pSFRUGfl4__1v9jmH3MNycE2IM",
        range: 'G4:H27',
        valueInputOption: 'USER_ENTERED',
        resource: { values: fillSpreadsheet }
    })


}

function calculateAverage(p1, p2, p3) {
    average = (p1 + p2 + p3) / 3;
    return average;

}

function studentStatus(average, absencesNumber, totalClasses) { // sets the status of students
    if (average < 50) {
        return averageStatus = 'Reprovado por nota';
    } else if (average >= 50 && average < 70) {
        return averageStatus = 'Exame Final';
    } else if (average >= 70) {
        return averageStatus = 'Aprovado';
    } else if (absencesNumber > totalClasses * 0.25) {
        return averageStatus = 'Reprovado por faltas';
    }

}

function calculateNAF(average, averageStatus) { 
    if (averageStatus == 'Exame Final') {
        naf = 100 - average;
        return naf;
    } else {
        return naf = 0;
    }

}
function absenceStatus(totalClasses, absencesNumber) { // sets student status based on absences
    if (absencesNumber > ((totalClasses * 0.25))) {
        return averageStatus = 'reprovado por faltas';
    } else {
        return averageStatus = averageStatus;
    }

}