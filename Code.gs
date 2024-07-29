function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function submitAttendance(employeeId, action, userLocation) {
  var sheetName = getMonthName();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(['Timestamp(IN)', 'User Location(IN)', 'Location(IN)', 'Email Address', 'Your ID', 'Timestamp(OUT)', 'User Location(OUT)', 'Location(OUT)', 'Status', 'Working Hours']);
  }

  var email = Session.getActiveUser().getEmail();
  var timestamp = new Date();

  // validate ID and Email
  var validation = validateEmployee(employeeId, email);
  if (!validation.isValid) {
    return { message: action === 'checkin' ? 'Check in failed' : 'Check out failed' };
  }

  var locationStatus = isInArea(userLocation) ? 'In Area' : 'Out Area';

  if (action === 'checkin') {
    var status = getCheckInStatus(employeeId, timestamp);
    Logger.log('Check-in status: ' + status);
    handleCheckIn(sheet, timestamp, userLocation, locationStatus, email, employeeId, status);

    return { message: 'Check in successful', locationStatus: locationStatus };

  } else if (action === 'checkout') {
    var status = getCheckOutStatus(employeeId, timestamp);
    Logger.log('Check-out status: ' + status);
    handleCheckOut(sheet, timestamp, userLocation, locationStatus, email, employeeId, status);

    return { message: 'Check out successful', locationStatus: locationStatus };
  }
}


function validateEmployee(employeeId, email) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = spreadsheet.getSheetByName('Worker_Namelist');
  if (!employeeSheet) {
    Logger.log('Employee sheet not found');
    return { isValid: false };
  }

  var data = employeeSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === employeeId.toString() && data[i][3] === email) {
      return { isValid: true };
    }
  }
  return { isValid: false };
}

function handleCheckIn(sheet, timestamp, userLocation, locationStatus, email, employeeId, status) {
  var row = findRowForCheckIn(sheet, email, employeeId);
  if (row) {
    var currentLocationStatus = sheet.getRange(row, 3).getValue(); // Location(IN)
    if (currentLocationStatus === 'In Area') {
      Logger.log('Already checked in to In Area. No action taken.');
      return;
    } else {
      sheet.getRange(row, 1).setValue(timestamp); // Timestamp(IN)
      sheet.getRange(row, 2).setValue(userLocation); // User Location(IN)
      sheet.getRange(row, 3).setValue(locationStatus); // Location(IN)
      sheet.getRange(row, 9).setValue(status); // Status
      formatTimestampColumn(sheet, 1); // Ensure the format of the Timestamp(IN) column is set correctly
    }
  } else {
    var rowData = [timestamp, userLocation, locationStatus, email, employeeId, null, null, null, status, null];
    sheet.appendRow(rowData);
  }
}

function handleCheckOut(sheet, timestamp, userLocation, locationStatus, email, employeeId, status) {
  var row = findRow(sheet, email, employeeId, timestamp);
  if (row) {
    var currentLocationStatus = sheet.getRange(row, 8).getValue(); // Location(OUT)
    if (currentLocationStatus === 'In Area' && locationStatus === 'Out Area') {
      Logger.log('Already checked out to In Area. No action taken.');
      return;
    } else {
      sheet.getRange(row, 6).setValue(timestamp); // Timestamp(OUT)
      sheet.getRange(row, 7).setValue(userLocation); // User Location(OUT)
      sheet.getRange(row, 8).setValue(locationStatus); // Location(OUT)
      sheet.getRange(row, 9).setValue(status); // Status
      formatTimestampColumn(sheet, 6); // Ensure the format of the Timestamp(OUT) column is set correctly

      // Calculate working hours
      var checkInTimestamp = new Date(sheet.getRange(row, 1).getValue());
      var workingHours = (timestamp - checkInTimestamp) / (1000 * 60 * 60); // Convert ms to hours
      sheet.getRange(row, 10).setValue(workingHours.toFixed(2)); // Set working hours
    }
  } else {
    Logger.log('No matching row found for check out.');
  }
}

function getMonthName() {
  var date = new Date();
  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return monthNames[date.getMonth()];
}

function findRow(sheet, email, employeeId, timestamp) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === email && data[i][4].toString() === employeeId.toString()) {
      var inDate = new Date(data[i][0]);
      if (isSameDay(inDate, timestamp)) {
        return i + 1; // Return the row number
      }
    }
  }
  return null;
}

function findRowForCheckIn(sheet, email, employeeId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === email && data[i][4].toString() === employeeId.toString()) {
      var inDate = new Date(data[i][0]);
      if (isSameDay(inDate, new Date())) {
        return i + 1; // Return the row number
      }
    }
  }
  return null;
}

function isSameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

function formatTimestampColumn(sheet, column) {
  // Set the format of the specified column to show date and time
  sheet.getRange(2, column, sheet.getLastRow() - 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function isInArea(userLocation) {
  var centerLat = 3.0448612;
  var centerLng = 101.780167;
  var [userLat, userLng] = userLocation.split(',').map(Number);

  var distance = haversineDistance([centerLat, centerLng], [userLat, userLng]);

  return distance <= 0.7; // 0.7 km = 700 meters
}

function haversineDistance(coords1, coords2) {
  var R = 6371; // Earth's radius in kilometers
  var dLat = toRadians(coords2[0] - coords1[0]);
  var dLng = toRadians(coords2[1] - coords1[1]);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(toRadians(coords1[0])) * Math.cos(toRadians(coords2[0])) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  var distance = R * c;

  return distance;
}

function toRadians(degrees) {
  return degrees * (Math.PI / 180);
}

function getCheckInStatus(employeeId, timestamp) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = spreadsheet.getSheetByName('Worker_Namelist');
  if (!employeeSheet) {
    Logger.log('Employee sheet not found');
    return 'Absent';
  }

  var data = employeeSheet.getDataRange().getValues();
  Logger.log('Employee data: ' + JSON.stringify(data));

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === employeeId.toString()) {
      var shift = data[i][4]; // Shift time in format "HH:MM,HH:MM"
      Logger.log('Employee shift: ' + shift);
      var shiftStart = shift.split(',')[0]; // Get start time
      var shiftStartHour = parseInt(shiftStart.split(':')[0]);
      var shiftStartMinute = parseInt(shiftStart.split(':')[1]);
      var shiftStartTime = new Date(timestamp);
      shiftStartTime.setHours(shiftStartHour, shiftStartMinute, 0, 0);
      
      var gracePeriodEnd = new Date(shiftStartTime.getTime() + 10 * 60000); // 10 minutes grace period
      if (timestamp <= gracePeriodEnd) {
        return 'Present';
      } else {
        return 'Late';
      }
    }
  }
  Logger.log('Employee not found or shift time not set');
  return 'Absent'; // Default to absent if no matching employee or shift time found
}

function getCheckOutStatus(employeeId, timestamp) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = spreadsheet.getSheetByName('Worker_Namelist');
  if (!employeeSheet) {
    Logger.log('Employee sheet not found');
    return 'Absent';
  }

  var data = employeeSheet.getDataRange().getValues();
  Logger.log('Employee data: ' + JSON.stringify(data));

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === employeeId.toString()) {
      var shift = data[i][4]; // Shift time in format "HH:MM,HH:MM"
      Logger.log('Employee shift: ' + shift);
      var shiftEnd = shift.split(',')[1]; // Get end time
      var shiftEndHour = parseInt(shiftEnd.split(':')[0]);
      var shiftEndMinute = parseInt(shiftEnd.split(':')[1]);
      var shiftEndTime = new Date(timestamp);
      shiftEndTime.setHours(shiftEndHour, shiftEndMinute, 0, 0);
      
      if (timestamp < shiftEndTime) {
        return 'Excused';
      }
    }
  }
  Logger.log('Employee not found or shift time not set');
  return 'Absent'; // Default to absent if no matching employee or shift time found
}

function markAbsentEmployees() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = spreadsheet.getSheetByName('Worker_Namelist');
  var attendanceSheet = spreadsheet.getSheetByName(getMonthName());

  if (!employeeSheet || !attendanceSheet) {
    Logger.log('Employee sheet or attendance sheet not found');
    return;
  }

  var employeeData = employeeSheet.getDataRange().getValues();
  var attendanceData = attendanceSheet.getDataRange().getValues();
  var currentDate = new Date();

  for (var i = 1; i < employeeData.length; i++) {
    var employeeId = employeeData[i][0];
    var email = employeeData[i][3];
    var shift = employeeData[i][4];
    var shiftEnd = shift.split(',')[1];
    var shiftEndHour = parseInt(shiftEnd.split(':')[0]);
    var shiftEndMinute = parseInt(shiftEnd.split(':')[1]);
    var shiftEndTime = new Date(currentDate);
    shiftEndTime.setHours(shiftEndHour, shiftEndMinute, 0, 0);

    var hasCheckedIn = false;

    for (var j = 1; j < attendanceData.length; j++) {
      if (attendanceData[j][3] === email && attendanceData[j][4].toString() === employeeId.toString() && isSameDay(new Date(attendanceData[j][0]), currentDate)) {
        hasCheckedIn = true;
        break;
      }
    }

    if (!hasCheckedIn && currentDate > shiftEndTime) {
      attendanceSheet.appendRow([timestamp, null, null, email, employeeId, timestamp, null, null, 'Absent', null]);
    }
  }
}

function createIntervalTrigger() {
  // Delete any existing triggers to avoid duplicates
  deleteTriggers();

  ScriptApp.newTrigger('checkForNewMonth')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(59)
    .create();


  ScriptApp.newTrigger('sendDailyEmail')
    .timeBased()
    .atHour(17)
    .nearMinute(0)
    .everyDays(1)
    .create();
}



function createMonthlySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentMonthName = getMonthName();

  var sheet = spreadsheet.getSheetByName(currentMonthName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(currentMonthName);
    sheet.appendRow(['Timestamp(IN)', 'User Location(IN)', 'Location(IN)', 'Email Address', 'Your ID', 'Timestamp(OUT)', 'User Location(OUT)', 'Location(OUT)', 'Status', 'Working Hours']);
    Logger.log("Sheet for " + currentMonthName + " created.");
  } else {
    Logger.log("Sheet for " + currentMonthName + " already exists.");
  }
  return sheet;
}

function markAbsentEmployees() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = spreadsheet.getSheetByName('Worker_Namelist');
  var attendanceSheet = createMonthlySheet();

  if (!employeeSheet || !attendanceSheet) {
    Logger.log('Employee sheet or attendance sheet not found');
    return;
  }

  var employeeData = employeeSheet.getDataRange().getValues();
  var attendanceData = attendanceSheet.getDataRange().getValues();
  var currentDate = new Date();

  for (var i = 1; i < employeeData.length; i++) {
    var employeeId = employeeData[i][0];
    var email = employeeData[i][3];
    var shift = employeeData[i][4];
    var shiftEnd = shift.split(',')[1];
    var shiftEndHour = parseInt(shiftEnd.split(':')[0]);
    var shiftEndMinute = parseInt(shiftEnd.split(':')[1]);
    var shiftEndTime = new Date(currentDate);
    shiftEndTime.setHours(shiftEndHour, shiftEndMinute, 0, 0);

    var hasCheckedIn = false;

    for (var j = 1; j < attendanceData.length; j++) {
      if (attendanceData[j][3] === email && attendanceData[j][4].toString() === employeeId.toString() && isSameDay(new Date(attendanceData[j][0]), currentDate)) {
        hasCheckedIn = true;
        break;
      }
    }

    if (!hasCheckedIn && currentDate > shiftEndTime) {
      attendanceSheet.appendRow([null, null, null, email, employeeId, null, null, null, 'Absent', null]);
    }
  }

  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('LAST_RUN_DATE', currentDate.toDateString());
}

function checkForNewMonth() {
  var properties = PropertiesService.getScriptProperties();
  var lastRunDate = new Date(properties.getProperty('LAST_RUN_DATE'));
  var currentDate = new Date();

  if (lastRunDate.getMonth() !== currentDate.getMonth() || lastRunDate.getFullYear() !== currentDate.getFullYear()) {
    Logger.log("New month detected. Creating new sheet.");
    createMonthlySheet();
  }

  markAbsentEmployees();
}

function deleteTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}


//Email part
function getLatestDate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dateTimeColumn = 'A'; 
  var lastRow = sheet.getLastRow();
  var dateTimes = sheet.getRange(dateTimeColumn + '2:' + dateTimeColumn + lastRow).getValues();

  // Filter out invalid dates
  var validDates = dateTimes.map(function(dateTime) {
    var date = new Date(dateTime);
    return isNaN(date.getTime()) ? null : date;
  }).filter(function(date) {
    return date !== null;
  });

  // Ensure there are valid dates to process
  if (validDates.length === 0) {
    throw new Error("No valid dates found in the column.");
  }

  var latestDate = new Date(Math.max.apply(null, validDates));
  return latestDate;
}

function getAttendanceData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dateTimeColumn = 'A';
  var idColumn = 'E';
  var statusColumn = 'I';
  var locationInColumn = 'C';  
  var locationOutColumn = 'H'; 
  var lastRow = sheet.getLastRow();
  
  var dateTimes = sheet.getRange(dateTimeColumn + '2:' + dateTimeColumn + lastRow).getValues();
  var ids = sheet.getRange(idColumn + '2:' + idColumn + lastRow).getValues();
  var statuses = sheet.getRange(statusColumn + '2:' + statusColumn + lastRow).getValues();
  var locationsIn = sheet.getRange(locationInColumn + '2:' + locationInColumn + lastRow).getValues();
  var locationsOut = sheet.getRange(locationOutColumn + '2:' + locationOutColumn + lastRow).getValues();
  
  var latestDate = getLatestDate();
  var attendance = {
    present: [],
    late: [],
    absent: [],
    outAreaIn: [],
    outAreaOut: [],
    excused: []
  };
  
  for (var i = 0; i < dateTimes.length; i++) {
    var rowDate = new Date(dateTimes[i][0]);
    if (rowDate.toDateString() === latestDate.toDateString()) {
      switch (statuses[i][0].toLowerCase()) {
        case 'present':
          attendance.present.push(ids[i][0]);
          break;
        case 'late':
          attendance.late.push(ids[i][0]);
          break;
        case 'absent':
          attendance.absent.push(ids[i][0]);
          break;
        case 'excused':
          attendance.excused.push(ids[i][0]);
          break;
      }
      // Check location in for 'Out Area'
      if (locationsIn[i][0].toLowerCase() === 'out area') {
        attendance.outAreaIn.push(ids[i][0]);
      }
      // Check location out for 'Out Area'
      if (locationsOut[i][0].toLowerCase() === 'out area') {
        attendance.outAreaOut.push(ids[i][0]);
      }
    }
  }
  
  return attendance;
}

function getTotalEmployees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName("Worker_Namelist"); //the name of your sheet with total employees
  var lastRow = employeeSheet.getLastRow();
  return lastRow - 1; // Assuming the first row is a header
}

function getEmployeeNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName("Worker_Namelist");
  var lastRow = employeeSheet.getLastRow();
  var idColumn = 'A'; // Assuming employee IDs are in column A
  var nameColumn = 'B'; // Assuming employee names are in column B
  
  var ids = employeeSheet.getRange(idColumn + '2:' + idColumn + lastRow).getValues();
  var names = employeeSheet.getRange(nameColumn + '2:' + nameColumn + lastRow).getValues();
  
  var employeeMap = {};
  for (var i = 0; i < ids.length; i++) {
    employeeMap[ids[i][0]] = names[i][0];
  }
  
  return employeeMap;
}

function sendDailyEmail() {
  var latestDate = getLatestDate();
  var attendance = getAttendanceData();
  var totalEmployees = getTotalEmployees();
  var employeeNames = getEmployeeNames();
  
  var recipients = ['lowtengf@gmail.com','junyo@gmail.com','hawmi@gmail.com']; //enter email for management
  var subject = 'Daily Attendance Report for ' + Utilities.formatDate(latestDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  
  function formatEmployeeList(employees) {
    return employees.map(function(id) {
      return id + ' - ' + (employeeNames[id] || 'Unknown');
    }).join('\n');
  }
  
  var body = 'Date: ' + Utilities.formatDate(latestDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') + '\n' +
             'Total Employees: ' + totalEmployees + '\n\n' +
             'Present: ' + attendance.present.length + ' employees\n\n' +
             'Late (' + attendance.late.length + '):\n' + 
             formatEmployeeList(attendance.late) + '\n\n' +
             'Absent (' + attendance.absent.length + '):\n' + 
             formatEmployeeList(attendance.absent) + '\n\n' +
             'Excused (' + attendance.excused.length + '):\n' + 
             formatEmployeeList(attendance.excused) + '\n\n' +
             'Location(In) (out area) (' + attendance.outAreaIn.length + '):\n' + 
             formatEmployeeList(attendance.outAreaIn) + '\n\n' +
             'Location(Out) (out area) (' + attendance.outAreaOut.length + '):\n' + 
             formatEmployeeList(attendance.outAreaOut);
  
  MailApp.sendEmail(recipients.join(','), subject, body);
}
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance Tools')
    .addItem('Send Email Report', 'sendDailyEmail')
    .addToUi();
}

function setup() {
  createTrigger();
}
