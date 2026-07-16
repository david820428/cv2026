function doGet(e) {
  // Check if the URL has ?page=account
  var page = e.parameter.page;
  
  // If no page is specified, show the main resume, otherwise show the requested page
  var folder = page ? page : 'index'; 
  
  return HtmlService.createHtmlOutputFromFile(folder)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet) {
    ss.insertSheet("Users");
    sheet.appendRow(["Timestamp", "Username", "Password", "Email"]);
  }
}

/** * SIGN UP LOGIC 
 */
function signUpUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  // 1. Validation Logic: Check for blank fields
  // .trim() removes any accidental spaces the user might have typed
  if (!userData.username || userData.username.trim() === "" || 
      !userData.password || userData.password.trim() === "") {
    return { 
      success: false, 
      message: "Username and Password cannot be blank." 
    };
  }
  // Check if username already exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userData.username) {
      return { success: false, message: "Username already taken." };
    }
  }
  
  // Append new user
  sheet.appendRow([new Date(), userData.username, userData.password, userData.email]);
  return { success: true, message: "Account created successfully!" };
}

/** * LOGIN LOGIC (The "SQL Select" equivalent)
 */
function loginUser(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  // Debug: See what the user typed vs what is in the sheet
  console.log("Searching for User: " + username);

  for (let i = 1; i < data.length; i++) {
    let sheetUser = data[i][1].toString().trim(); // Column B
    let sheetPass = data[i][2].toString().trim(); // Column C
    
    console.log("Checking Row " + (i+1) + ": Found '" + sheetUser + "' with pass '" + sheetPass + "'");

    if (sheetUser === username.trim() && sheetPass === password.trim()) {
      console.log("MATCH FOUND!");
      return { success: true, username: username };
    }
  }
  
  console.log("NO MATCH FOUND after checking " + (data.length - 1) + " rows.");
  return { success: false, message: "Invalid username or password." };
}

