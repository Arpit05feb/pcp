const express = require('express');
const bodyParser = require('body-parser');
const ejs = require('ejs');
const path = require('path');
const xlsx = require('xlsx');
const _ = require('lodash');

const app = express();

// Set up middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Read data from Excel file
const excelFilePath = 'E:\\Readings\\web\\missing-slots-app\\views\\Book1.xlsx'; // Replace with your actual Excel file path
const workbook = xlsx.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Logic to find other slots and remaining slots
function findOtherSlots(identifier, value) {
    // Filter data based on CourseCode or Section
    const filteredData = _.filter(excelData, { [identifier]: value });

    // Extract registration numbers for the given CourseCode or Section
    const regNos = _.map(filteredData, 'RegdNo');

    // Filter data for the registration numbers excluding the entered identifier's value
    const otherSlotsData = _.filter(excelData, (row) => row[identifier] !== value && regNos.includes(row.RegdNo));

    // Extract distinct other slots for the filtered registration numbers
    const otherSlotsSet = new Set(_.map(otherSlotsData, 'New Time Slot'));

    // Convert set to array for consistent output
    const otherSlots = Array.from(otherSlotsSet);

    // Extract all available slots
    const allSlotsSet = new Set(_.map(excelData, 'New Time Slot'));

    // Find remaining slots by subtracting other busy slots from all available slots
    const remainingSlots = Array.from(new Set([...allSlotsSet].filter(x => !otherSlotsSet.has(x))));

    return { otherSlots, remainingSlots };
}

// Set up routes
app.get('/', (req, res) => {
    res.render('index', { result: null });
});

app.post('/process', (req, res) => {
    const identifier = req.body.identifier; // Assuming you have a form field with name 'identifier'
    const value = req.body.value; // Assuming you have a form field with name 'value'

    // Logic to find other slots and remaining slots
    const { otherSlots, remainingSlots } = findOtherSlots(identifier, value);

    // Render the result on the index page
    res.render('index', {
        result: `Selected Value: ${value} \n\nOther Busy Slots: ${otherSlots.join(', ') || 'Null'} \nRemaining Slots: ${remainingSlots.join(', ') || 'Null'}`
    
    });
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
