// Creating a function that works on opening the sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Creating a custom menu named 'Custom Imagine' for all processes
  ui.createMenu('Custom Imagine')

    // Adding the menu to create dummy datasets
    .addItem('Create Dummy Datasets', 'ImagineDummyData')

    // Adding the menu to check for data quality in all datasets
    .addSubMenu(ui.createMenu('Data Quality Checks')
      .addItem('Website Visits Quality Check', 'checkWebsiteVisitsDataQuality')
      .addItem('Application Rates Quality Check', 'checkApplicationRatesDataQuality')
      .addItem('Conversion Rates Quality Check', 'checkConversionRatesDataQuality'))

    // Adding the menu to visualize KPI
    .addSubMenu(ui.createMenu('Visualize KPIs')
      .addItem('Website Visits', 'visualizeWebsiteVisits'))

    // Adding the menu to monitor application processes
    .addItem('Monitor Application Processes', 'monitorApplicationProcess')

    // Adding the menu to send for Mail to Emma
    .addItem('A message to you Emma', 'sendApplicationEmail')

    // Adding the menu to the Google Sheet UI
    .addToUi();
}

// function to send email to Emma of Imagine
function sendApplicationEmail() {
  // Recipient's email address
  var recruiterEmail = 'adepejuoladapo@gmail.com';

  // Subject and body of the email
  var subject = 'Data Analyst and Google Sheets Developer Application By Google Apps Script';
  var body = 'Hi Emma,\n\n If you are reading this mail its because youve activated the cureview my project at the following link: [project link]';

  // Send the email
  GmailApp.sendEmail(recruiterEmail, subject, body);
}

function monitorApplicationProcess() {
  // writing a variable to access spreadsheet ui
  var ui = SpreadsheetApp.getUi();

  // Prompt the user to input a date
  var dateString = ui.prompt('Enter a Date', 'Please enter a date between 01-01-2023 and 01-31-2024 (MM/dd/yyyy):', ui.ButtonSet.OK_CANCEL);

  // Check if the user clicked "OK" and entered a valid date
  if (dateString.getSelectedButton() === ui.Button.OK) {
    var inputDate = new Date(dateString.getResponseText());

    // Check if the input date is within the valid range
    if (!isNaN(inputDate) && inputDate >= new Date('01/01/2023') && inputDate <= new Date('01/31/2024')) {
      var sheetName = 'application_rates'; // Update with the correct sheet name
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

      // Assuming data starts from the second row (header in the first row)
      var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      var data = dataRange.getValues();

      // creating variables
      var totalApplications = 0;
      var totalConversionRate = 0;
      var totalTime = 0;
      var count = 0;

      // Iterate through the data to calculate metrics for the input date
      for (var i = 0; i < data.length; i++) {
        var currentDate = new Date(data[i][0]); 

        if (currentDate.toDateString() === inputDate.toDateString()) {
          totalApplications += data[i][1]; 
          totalConversionRate += data[i][2];
          totalTime += data[i][3];
          count++;
        }
      }

      // Checking if there are records for the input date
      if (count > 0) {
        var averageConversionRate = totalConversionRate / count;
        var logString = 'Metrics for ' + Utilities.formatDate(inputDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') + ':\n' +
          'Average Conversion Rate: ' + averageConversionRate.toFixed(2) + '\n' +
          'Total Applications: ' + totalApplications + '\n' +
          'Total Completion Time: ' + totalTime + ' minutes';

        Logger.log(logString);
        ui.alert('Application Process Metrics', logString, ui.ButtonSet.OK);
      } else {
        ui.alert('No Data Found', 'No data available for the selected date.', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Invalid Date', 'Please enter a valid date within the specified range.', ui.ButtonSet.OK);
    }
  }
}


var selectedMetric = 'Total Visits'; // Default metric

function visualizeWebsiteVisits() {
  var sheetName = 'website_visits';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Clear existing charts
  var charts = sheet.getCharts();
  charts.forEach(function (chart) {
    sheet.removeChart(chart);
  });

  // Convert date values to JavaScript Date objects
  var dateRange = sheet.getRange('A2:A' + sheet.getLastRow());
  var dateValues = dateRange.getValues().flat().map(function (value) {
    return [new Date(value)];
  });
  dateRange.setValues(dateValues);

  // Create a line chart based on the selected metric
  var metricIndex = getMetricColumnIndex(selectedMetric) + 2; // Adding 2 to account for date column and 1-based indexing
  var chartBuilder = sheet.newChart().asLineChart();
  chartBuilder.addRange(sheet.getRange('A1:B' + sheet.getLastRow()));
  chartBuilder.setOption('useFirstColumnAsDomain', true);
  chartBuilder.setPosition(3, 1, 0, 0);

  // Set chart title and axis labels
  chartBuilder.setOption('title', selectedMetric + ' Over Time');
  chartBuilder.setOption('vAxes', {
    0: { title: selectedMetric },
  });
  chartBuilder.setOption('hAxis', {
    title: 'Date',
    format: 'MM/dd/yyyy', // Customizing date format as needed
  });

  // Insert chart into a new sheet named "Website Visit Visualization"
  var newSheetName = 'Website Visit Visualization';
  var newSheet = spreadsheet.getSheetByName(newSheetName);
  if (!newSheet) {
    newSheet = spreadsheet.insertSheet(newSheetName);
  } else {
    newSheet.clear(); // Clear existing data if sheet already exists
  }
  newSheet.insertChart(chartBuilder.build());

  // Open the sidebar
  showSidebar();
}
// function to generate metric column index
function getMetricColumnIndex(metric) {
  var headers = generate_headers('website_visits');
  return headers.indexOf(metric);
}

// Function to update the selected metric
function updateSelectedMetric(metric) {
  selectedMetric = metric;
  visualizeWebsiteVisits();
}

// Function to show the sidebar
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
    .setWidth(300)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}
// function to get metrics in sidebar
function getMetrics() {
  var metrics = generate_headers('website_visits').slice(1); // Exclude 'Date' from metrics
  return metrics;
}

function checkWebsiteVisitsDataQuality() {
  var sheetName = 'website_visits';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = dataRange.getValues();

  var nullCount = 0;
  var duplicateCount = 0;
  var visitedDates = [];

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      // Check for null values
      if (data[i][j] === null || data[i][j] === '') {
        nullCount++;
      }
    }

    // Check for duplicate dates
    var currentDate = data[i][0];
    if (visitedDates.includes(currentDate)) {
      duplicateCount++;
    } else {
      visitedDates.push(currentDate);
    }
  }

  var logString = 'There are ' + nullCount + ' null entries and ' + duplicateCount + ' duplicated entries in the website visits dataset.';
  Logger.log(logString);
  SpreadsheetApp.getUi().alert('Data Quality Check', logString, SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkApplicationRatesDataQuality() {
  var sheetName = 'application_rates';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = dataRange.getValues();

  var nullCount = 0;
  var duplicateCount = 0;
  var visitedDates = [];

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      // Check for null values
      if (data[i][j] === null || data[i][j] === '') {
        nullCount++;
      }
    }

    // Check for duplicate dates
    var currentDate = data[i][0];
    if (visitedDates.includes(currentDate)) {
      duplicateCount++;
    } else {
      visitedDates.push(currentDate);
    }
  }

  var logString = 'There are ' + nullCount + ' null entries and ' + duplicateCount + ' duplicated entries in the application rates dataset.';
  Logger.log(logString);
  SpreadsheetApp.getUi().alert('Data Quality Check', logString, SpreadsheetApp.getUi().ButtonSet.OK);
}


function checkConversionRatesDataQuality() {
  var sheetName = 'conversion_rates';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);


  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = dataRange.getValues();

  var nullCount = 0;
  var duplicateCount = 0;
  var visitedDates = [];

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      // Check for null values
      if (data[i][j] === null || data[i][j] === '') {
        nullCount++;
      }
    }

    // Check for duplicate dates
    var currentDate = data[i][0];
    if (visitedDates.includes(currentDate)) {
      duplicateCount++;
    } else {
      visitedDates.push(currentDate);
    }
  }

  var logString = 'There are ' + nullCount + ' null entries and ' + duplicateCount + ' duplicated entries in the conversion rates dataset.';
  Logger.log(logString);
  SpreadsheetApp.getUi().alert('Data Quality Check', logString, SpreadsheetApp.getUi().ButtonSet.OK);
}


function ImagineDummyData() {
  createDummyDataSheet('website_visits');
  createDummyDataSheet('application_rates');
  createDummyDataSheet('conversion_rates');
  deleteSheetsStartingWithSheet();
}

function deleteSheetsStartingWithSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (sheetName.toLowerCase().indexOf('sheet') === 0) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}


function createDummyDataSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  } else {
    sheet.clear(); // Clear existing data if sheet already exists
  }

  // Add headers
  var headers = generate_headers(sheetName);
  var data = [headers];

  // Generate dummy data for the specified date range
  var startDate = new Date('2023-01-01');
  var endDate = new Date('2024-01-31');
  var currentDate = new Date(startDate);

  while (currentDate <= endDate) {
    // Clone the date object to avoid modifying the original object
    var clonedDate = new Date(currentDate);
    
    // Format the date to MM/dd/yyyy
    var formattedDate = Utilities.formatDate(clonedDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy');

    var rowData = generate_row_values(sheetName);
    rowData.unshift(formattedDate);

    // Add data to the array
    data.push(rowData);

    // Move to the next day
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // Set the values for the entire data range
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('Dummy ' + sheetName + ' data created successfully.');
}


function generate_headers(sheet_name) {
  switch(sheet_name) {
    case 'website_visits':
      return ['Date', 'Total Visits', 'Unique Visitors', 'Pageviews', 'Bounce Rate', 'Avg Session Duration'];
    case 'application_rates':
      return ['Date', 'Total Applications', 'Conversion Rate', 'Application Completion Time'];
    case 'conversion_rates':
      return ['Date', 'Conversion Rate', 'Conversion Action'];
    default:
      return [];
  }
}

function generate_row_values(sheet_name) {
  switch(sheet_name) {
    case 'website_visits':
      return [
        Math.floor(Math.random() * 1000) + 500,
        Math.floor(Math.random() * 800) + 200,
        Math.floor(Math.random() * 1500) + 500,
        Math.random() * 100,
        Math.floor(Math.random() * 300) + 60
      ];
    case 'application_rates':
      return [
        Math.floor(Math.random() * 50) + 10,
        Math.random() * 10,
        Math.floor(Math.random() * 120) + 30
      ];
    case 'conversion_rates':
      return [
        Math.random() * 10,
        Math.floor(Math.random() * 3)
      ];
    default:
      return [];
  }
}