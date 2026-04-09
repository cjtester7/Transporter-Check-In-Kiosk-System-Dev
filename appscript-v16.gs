// CarsRUs Transporter Check-In System — Apps Script Backend
// Version: appscript-v16.gs
// Deploy as Web App: Execute as Me, Anyone can access

// ============================================================
// CONFIGURATION
const SHEET_URL = "https://docs.google.com/spreadsheets/d/1fAcPohNPc32egYUeLq3SbZGZnEufQNrggM0xTpn9a2w";
const MASTER_ADMIN_USERNAME = "superadmin";
const MASTER_ADMIN_PASSWORD = "changeme123";
// ============================================================

const SHEET_NAME = "TransporterLog";
const USERS_SHEET_NAME = "Users";
const AUDIT_SHEET_NAME = "AuditLog";
const NOTES_SHEET_NAME = "Notes";
const PHOTOS_SHEET_NAME = "Photos";
const DRIVE_FOLDER_NAME = "CarsRUs";

const NOTES_HEADERS = ["Note ID", "Row ID", "Note", "Added By", "Timestamp"];
const PHOTOS_HEADERS = ["Photo ID", "Row ID", "File Name", "Drive URL", "Uploaded By", "Timestamp"];

const AUDIT_HEADERS = [
  "Timestamp", "Action", "Row ID", "Driver Name", "Carrier",
  "Changed By", "Changes", "Previous Values"
];

const HEADERS = [
  "Date", "Driver Name", "Driver Phone", "Carrier", "Carrier Phone",
  "Lane", "Time In", "Time Out", "Drop Off", "Pickup",
  "Status", "Vehicle Types", "Comments", "Gate", "Queue Position",
  "Est. Wait (min)", "Signed In By", "Signed Out By", "Row ID", "Check-In Timestamp"
];

const USER_HEADERS = [
  "Username", "Password Hash", "Display Name", "Phone", "Carrier",
  "Group", "Status", "Created Date", "Last Login"
];

const COL = {};
HEADERS.forEach((h, i) => COL[h] = i);

const UCOL = {};
USER_HEADERS.forEach((h, i) => UCOL[h] = i);

// ============================================================
// REQUEST ROUTING
// ============================================================
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const params = e.parameter || {};
    let action, body;
    if (params.data) {
      body = JSON.parse(params.data);
      action = body.action;
    } else {
      action = params.action;
      body = params;
    }
    let result;
    switch (action) {
      case "getAll":         result = getAllRecords(); break;
      case "checkIn":        result = checkIn(body); break;
      case "checkOut":       result = checkOut(body); break;
      case "updateStatus":   result = updateStatus(body); break;
      case "updateRecord":   result = updateRecord(body); break;
      case "getQueue":       result = getQueueInfo(); break;
      case "fixQueue":       result = fixQueueNow(); break;
      case "getAuditLog":    result = getAuditLog(); break;
      case "addNote":        result = addNote(body); break;
      case "getNotes":       result = getNotes(body); break;
      case "deleteNote":     result = deleteNote(body); break;
      case "uploadPhoto":    result = uploadPhoto(body); break;
      case "getPhotos":      result = getPhotos(body); break;
      case "deletePhoto":    result = deletePhoto(body); break;
      case "login":          result = loginUser(body); break;
      case "register":       result = registerUser(body); break;
      case "changePassword": result = changePassword(body); break;
      case "getUsers":       result = getUsers(body); break;
      case "updateUser":     result = updateUser(body); break;
      case "deleteUser":     result = deleteUser(body); break;
      default:               result = { error: "Unknown action: " + action };
    }
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }
  return output;
}

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch (action) {
      case "uploadPhoto": result = uploadPhoto(body); break;
      default: result = doGet(e); // fall back to GET handler for other actions
    }
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }
  return output;
}

// ============================================================
// SHEET HELPERS
// ============================================================
function getSpreadsheet() {
  return SpreadsheetApp.openByUrl(SHEET_URL);
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, HEADERS.length);
    hr.setBackground("#1a1a2e");
    hr.setFontColor("#ffffff");
    hr.setFontWeight("bold");
  }
  return sheet;
}

function getUsersSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(USER_HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, USER_HEADERS.length);
    hr.setBackground("#1a1a2e");
    hr.setFontColor("#ffffff");
    hr.setFontWeight("bold");
    // Seed master admin
    seedMasterAdmin(sheet);
  }
  return sheet;
}

// ============================================================
// AUDIT LOG
// ============================================================
function getAuditSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET_NAME);
    sheet.appendRow(AUDIT_HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, AUDIT_HEADERS.length);
    hr.setBackground("#1a1a2e");
    hr.setFontColor("#ffffff");
    hr.setFontWeight("bold");
    // Widen columns for readability
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(7, 300); // Changes
    sheet.setColumnWidth(8, 300); // Previous Values
  }
  return sheet;
}

function appendAuditLog(action, rowId, driverName, carrier, changedBy, changes, previousValues) {
  try {
    const sheet = getAuditSheet();
    const tz = Session.getScriptTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy hh:mm:ss a");
    sheet.appendRow([
      ts,
      action,
      rowId || "",
      driverName || "",
      carrier || "",
      changedBy || "",
      changes || "",
      previousValues || ""
    ]);
  } catch(e) {
    // Never let audit logging failure break the main action
  }
}

function getAuditLog() {
  const sheet = getAuditSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { rows: [] };
  const headers = data[0];
  const rows = data.slice(1).reverse().map(row => { // newest first
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { rows };
}

// ============================================================
// NOTES
// ============================================================
function getNotesSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(NOTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTES_SHEET_NAME);
    sheet.appendRow(NOTES_HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, NOTES_HEADERS.length);
    hr.setBackground("#1a1a2e");
    hr.setFontColor("#ffffff");
    hr.setFontWeight("bold");
  }
  return sheet;
}

function getNotes(data) {
  const sheet = getNotesSheet();
  const rows = sheet.getDataRange().getValues();
  const notes = rows.slice(1)
    .filter(r => r[1] == data.rowId)
    .map(r => ({
      noteId:    r[0],
      rowId:     r[1],
      note:      r[2],
      addedBy:   r[3],
      timestamp: r[4]
    }));
  return { notes };
}

function addNote(data) {
  if (!data.rowId || !data.note) return { error: "Row ID and note are required." };
  const sheet = getNotesSheet();
  const tz = Session.getScriptTimeZone();
  const ts = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy hh:mm a");
  const noteId = "NOTE-" + new Date().getTime();
  sheet.appendRow([noteId, data.rowId, data.note, data.addedBy || "", ts]);
  appendAuditLog("Add Note", data.rowId, "", "", data.addedBy || "", "Note added", "");
  return { success: true, noteId, timestamp: ts };
}

function deleteNote(data) {
  if (!data.noteId) return { error: "Note ID required." };
  const sheet = getNotesSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.noteId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "Note not found." };
}

// ============================================================
// PHOTOS
// ============================================================
function getPhotosSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PHOTOS_SHEET_NAME);
    sheet.appendRow(PHOTOS_HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, PHOTOS_HEADERS.length);
    hr.setBackground("#1a1a2e");
    hr.setFontColor("#ffffff");
    hr.setFontWeight("bold");
  }
  return sheet;
}

function getOrCreateDriveFolder(rowId) {
  const root = DriveApp.getRootFolder();
  let carsrusFolder;
  const carsrusFolders = root.getFoldersByName(DRIVE_FOLDER_NAME);
  carsrusFolder = carsrusFolders.hasNext() ? carsrusFolders.next() : root.createFolder(DRIVE_FOLDER_NAME);
  let photoFolder;
  const photoFolders = carsrusFolder.getFoldersByName(rowId);
  photoFolder = photoFolders.hasNext() ? photoFolders.next() : carsrusFolder.createFolder(rowId);
  return photoFolder;
}

function uploadPhoto(data) {
  if (!data.rowId || !data.base64 || !data.fileName) return { error: "Row ID, file name and image data required." };
  try {
    const folder = getOrCreateDriveFolder(data.rowId);
    const mimeType = data.mimeType || "image/jpeg";
    const decoded = Utilities.base64Decode(data.base64);
    const blob = Utilities.newBlob(decoded, mimeType, data.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    const driveUrl = "https://drive.google.com/file/d/" + fileId + "/view";
    const thumbUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w300";
    const sheet = getPhotosSheet();
    const tz = Session.getScriptTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy hh:mm a");
    const photoId = "PHOTO-" + new Date().getTime();
    sheet.appendRow([photoId, data.rowId, data.fileName, driveUrl, data.uploadedBy || "", ts]);
    appendAuditLog("Upload Photo", data.rowId, "", "", data.uploadedBy || "", "Photo uploaded: " + data.fileName, "");
    return { success: true, photoId, driveUrl, thumbUrl, timestamp: ts };
  } catch(e) {
    return { error: e.message };
  }
}

function getPhotos(data) {
  const sheet = getPhotosSheet();
  const rows = sheet.getDataRange().getValues();
  const photos = rows.slice(1)
    .filter(r => r[1] == data.rowId)
    .map(r => {
      const fileId = r[3].match(/\/d\/([^\/]+)\//);
      return {
        photoId:    r[0],
        rowId:      r[1],
        fileName:   r[2],
        driveUrl:   r[3],
        thumbUrl:   fileId ? "https://drive.google.com/thumbnail?id=" + fileId[1] + "&sz=w300" : r[3],
        uploadedBy: r[4],
        timestamp:  r[5]
      };
    });
  return { photos };
}

function deletePhoto(data) {
  if (!data.photoId) return { error: "Photo ID required." };
  const sheet = getPhotosSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.photoId) {
      // Try to delete from Drive too
      try {
        const match = rows[i][3].match(/\/d\/([^\/]+)\//);
        if (match) DriveApp.getFileById(match[1]).setTrashed(true);
      } catch(e) {}
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "Photo not found." };
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function seedMasterAdmin(sheet) {
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy HH:mm");
  sheet.appendRow([
    MASTER_ADMIN_USERNAME,
    hashPassword(MASTER_ADMIN_PASSWORD),
    "Master Admin",
    "",
    "",
    "Admin",
    "Active",
    now,
    ""
  ]);
}

// ============================================================
// AUTH ACTIONS
// ============================================================
function loginUser(data) {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const username = (data.username || "").toLowerCase().trim();
  const passHash = hashPassword(data.password || "");

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][UCOL["Username"]].toLowerCase() === username &&
        rows[i][UCOL["Password Hash"]] === passHash) {
      const status = rows[i][UCOL["Status"]];
      if (status !== "Active") {
        return { error: "Account is " + status + ". Please contact an administrator." };
      }
      // Update last login
      const tz = Session.getScriptTimeZone();
      sheet.getRange(i + 1, UCOL["Last Login"] + 1).setValue(
        Utilities.formatDate(new Date(), tz, "MM/dd/yyyy HH:mm")
      );
      return {
        success: true,
        user: {
          username: rows[i][UCOL["Username"]],
          displayName: rows[i][UCOL["Display Name"]],
          phone: rows[i][UCOL["Phone"]],
          carrier: rows[i][UCOL["Carrier"]],
          group: rows[i][UCOL["Group"]]
        }
      };
    }
  }
  return { error: "Incorrect username or password." };
}

function registerUser(data) {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const username = (data.username || "").toLowerCase().trim();

  if (!username || !data.password || !data.displayName) {
    return { error: "Username, password and name are required." };
  }
  if (username.length < 3) {
    return { error: "Username must be at least 3 characters." };
  }
  if ((data.password || "").length < 6) {
    return { error: "Password must be at least 6 characters." };
  }

  // Check username not taken
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][UCOL["Username"]].toLowerCase() === username) {
      return { error: "Username already taken. Please choose another." };
    }
  }

  const group = data.group || "Driver";
  // Drivers and Guests activate immediately; Employees need approval
  const status = (group === "Employee") ? "Pending" : "Active";
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy HH:mm");

  sheet.appendRow([
    username,
    hashPassword(data.password),
    data.displayName || "",
    data.phone || "",
    data.carrier || "",
    group,
    status,
    now,
    ""
  ]);

  return { success: true, status, group };
}

function changePassword(data) {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const username = (data.username || "").toLowerCase().trim();
  const currentHash = hashPassword(data.currentPassword || "");
  const newHash = hashPassword(data.newPassword || "");

  if ((data.newPassword || "").length < 6) {
    return { error: "New password must be at least 6 characters." };
  }

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][UCOL["Username"]].toLowerCase() === username &&
        rows[i][UCOL["Password Hash"]] === currentHash) {
      sheet.getRange(i + 1, UCOL["Password Hash"] + 1).setValue(newHash);
      return { success: true };
    }
  }
  return { error: "Current password is incorrect." };
}

function getUsers(data) {
  // Only callable with valid admin session token passed from frontend
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const users = rows.slice(1).map((r, i) => ({
    rowIndex: i + 2,
    username:    r[UCOL["Username"]],
    displayName: r[UCOL["Display Name"]],
    phone:       r[UCOL["Phone"]],
    carrier:     r[UCOL["Carrier"]],
    group:       r[UCOL["Group"]],
    status:      r[UCOL["Status"]],
    created:     r[UCOL["Created Date"]],
    lastLogin:   r[UCOL["Last Login"]]
  }));
  return { users };
}

function updateUser(data) {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const username = (data.username || "").toLowerCase().trim();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][UCOL["Username"]].toLowerCase() === username) {
      const rowNum = i + 1;
      if (data.group)  sheet.getRange(rowNum, UCOL["Group"] + 1).setValue(data.group);
      if (data.status) sheet.getRange(rowNum, UCOL["Status"] + 1).setValue(data.status);
      if (data.displayName) sheet.getRange(rowNum, UCOL["Display Name"] + 1).setValue(data.displayName);
      // Admin can reset password
      if (data.newPassword) {
        if (data.newPassword.length < 6) return { error: "Password must be at least 6 characters." };
        sheet.getRange(rowNum, UCOL["Password Hash"] + 1).setValue(hashPassword(data.newPassword));
      }
      return { success: true };
    }
  }
  return { error: "User not found." };
}

function deleteUser(data) {
  const sheet = getUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const username = (data.username || "").toLowerCase().trim();

  // Protect master admin
  if (username === MASTER_ADMIN_USERNAME.toLowerCase()) {
    return { error: "The master admin account cannot be deleted." };
  }

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][UCOL["Username"]].toLowerCase() === username) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "User not found." };
}

// ============================================================
// TRANSPORTER LOG ACTIONS
// ============================================================
function getAllRecords() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { records: [], _v: 16 };
  const headers = data[0];
  const tz = Session.getScriptTimeZone();
  const timeColumns = ["Time In", "Time Out"];
  const dateColumns = ["Date"];
  const records = data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => {
      if (row[j] instanceof Date) {
        if (timeColumns.includes(h)) {
          obj[h] = Utilities.formatDate(row[j], tz, "hh:mm a");
        } else if (dateColumns.includes(h)) {
          obj[h] = Utilities.formatDate(row[j], tz, "MM/dd/yyyy");
        } else {
          obj[h] = Utilities.formatDate(row[j], tz, "MM/dd/yyyy hh:mm a");
        }
      } else {
        obj[h] = row[j];
      }
    });
    obj._rowIndex = i + 2;
    return obj;
  });
  return { records, _v: 16 };
}

function checkIn(data) {
  const sheet = getSheet();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(now, tz, "MM/dd/yyyy");
  const timeStr = Utilities.formatDate(now, tz, "hh:mm a");
  const allData = sheet.getDataRange().getValues();
  const activeRows = allData.slice(1).filter(r => r[COL["Status"]] === "Waiting" || r[COL["Status"]] === "In Progress");
  const queuePos = activeRows.length + 1;
  const estWait = (queuePos - 1) * 20;
  const rowId = "CR-" + now.getTime() + "-" + Math.random().toString(36).slice(2, 6).toUpperCase();

  const row = HEADERS.map(h => {
    switch(h) {
      case "Date":                return dateStr;
      case "Driver Name":         return data["Driver Name"] || "";
      case "Driver Phone":        return data["Driver Phone"] || "";
      case "Carrier":             return data["Carrier"] || "";
      case "Carrier Phone":       return data["Carrier Phone"] || "";
      case "Gate":                return data["Gate"] || "";
      case "Lane":                return data["Lane"] || "";
      case "Time In":             return timeStr;
      case "Time Out":            return "";
      case "Drop Off":            return data["Drop Off"] || 0;
      case "Pickup":              return data["Pickup"] || 0;
      case "Status":              return "Waiting";
      case "Vehicle Types":       return data["Vehicle Types"] || "";
      case "Comments":            return data["Comments"] || "";
      case "Queue Position":      return queuePos;
      case "Est. Wait (min)":     return estWait;
      case "Signed In By":        return data["Signed In By"] || "Self";
      case "Signed Out By":       return "";
      case "Row ID":              return rowId;
      case "Check-In Timestamp":  return now.getTime();
      default:                    return "";
    }
  });

  sheet.appendRow(row);
  appendAuditLog(
    "Check In",
    rowId,
    data["Driver Name"] || "",
    data["Carrier"] || "",
    data["Signed In By"] || "Self",
    "Status: Waiting | Queue: " + queuePos + " | Drop Off: " + (data["Drop Off"] || 0) + " | Pickup: " + (data["Pickup"] || 0),
    ""
  );
  return { success: true, rowId, queuePosition: queuePos, estWait, timeIn: timeStr };
}

function checkOut(data) {
  const sheet = getSheet();
  const allData = sheet.getDataRange().getValues();
  const rowId = data["rowId"] || data["Row ID"];
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][COL["Row ID"]] == rowId) {
      const now = new Date();
      const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
      const rowNum = i + 1;
      const driverName = allData[i][COL["Driver Name"]];
      const carrier = allData[i][COL["Carrier"]];
      const prevStatus = allData[i][COL["Status"]];
      sheet.getRange(rowNum, COL["Time Out"] + 1).setValue(timeStr);
      sheet.getRange(rowNum, COL["Status"] + 1).setValue("Completed");
      sheet.getRange(rowNum, COL["Signed Out By"] + 1).setValue(data["Signed Out By"] || "");
      sheet.getRange(rowNum, COL["Queue Position"] + 1).setValue("");
      sheet.getRange(rowNum, COL["Est. Wait (min)"] + 1).setValue("");
      resequenceQueue(sheet);
      appendAuditLog(
        "Check Out",
        rowId,
        driverName,
        carrier,
        data["Signed Out By"] || "",
        "Status: " + prevStatus + " → Completed | Time Out: " + timeStr,
        "Status: " + prevStatus
      );
      return { success: true, timeOut: timeStr };
    }
  }
  return { error: "Record not found" };
}

function updateStatus(data) {
  const sheet = getSheet();
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][COL["Row ID"]] == data["rowId"]) {
      const prevStatus = allData[i][COL["Status"]];
      const driverName = allData[i][COL["Driver Name"]];
      const carrier = allData[i][COL["Carrier"]];
      sheet.getRange(i + 1, COL["Status"] + 1).setValue(data["status"]);
      resequenceQueue(sheet);
      appendAuditLog(
        "Status Change",
        data["rowId"],
        driverName,
        carrier,
        data["changedBy"] || "",
        "Status: " + prevStatus + " → " + data["status"],
        "Status: " + prevStatus
      );
      return { success: true };
    }
  }
  return { error: "Record not found" };
}

function updateRecord(data) {
  const sheet = getSheet();
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][COL["Row ID"]] == data["rowId"]) {
      const rowNum = i + 1;
      const oldRow = allData[i];
      const updatable = [
        "Driver Name", "Driver Phone", "Carrier", "Carrier Phone",
        "Date", "Time Out", "Gate", "Lane", "Queue Position",
        "Status", "Drop Off", "Pickup", "Vehicle Types",
        "Comments", "Signed In By", "Signed Out By"
      ];
      const changes = [];
      const prevValues = [];
      updatable.forEach(field => {
        if (data[field] !== undefined) {
          const oldVal = String(oldRow[COL[field]] || "");
          const newVal = String(data[field] || "");
          if (oldVal !== newVal) {
            changes.push(field + ": " + (oldVal || "—") + " → " + (newVal || "—"));
            prevValues.push(field + ": " + (oldVal || "—"));
          }
          sheet.getRange(rowNum, COL[field] + 1).setValue(data[field]);
        }
      });
      appendAuditLog(
        "Edit Record",
        data["rowId"],
        oldRow[COL["Driver Name"]],
        oldRow[COL["Carrier"]],
        data["changedBy"] || "",
        changes.join(" | ") || "No changes",
        prevValues.join(" | ") || ""
      );
      return { success: true };
    }
  }
  return { error: "Record not found" };
}

function resequenceQueue(sheet) {
  const allData = sheet.getDataRange().getValues();
  const activeRows = [];
  for (let i = 1; i < allData.length; i++) {
    const status = allData[i][COL["Status"]];
    if (status === "Waiting" || status === "In Progress") {
      activeRows.push({ rowNum: i + 1, ts: Number(allData[i][COL["Check-In Timestamp"]]) || 0 });
    }
  }
  activeRows.sort((a, b) => a.ts - b.ts);
  activeRows.forEach((r, idx) => {
    sheet.getRange(r.rowNum, COL["Queue Position"] + 1).setValue(idx + 1);
    sheet.getRange(r.rowNum, COL["Est. Wait (min)"] + 1).setValue(idx * 20);
  });
}

function fixQueueNow() {
  resequenceQueue(getSheet());
  return { success: true, message: "Queue resequenced successfully" };
}

function getQueueInfo() {
  const allData = getSheet().getDataRange().getValues();
  const active = allData.slice(1).filter(r => r[COL["Status"]] === "Waiting" || r[COL["Status"]] === "In Progress");
  return { queueLength: active.length, nextPosition: active.length + 1, estWait: active.length * 20 };
}
