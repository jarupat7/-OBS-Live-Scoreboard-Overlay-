/**
 * =========================================================================
 * Badminton Scoreboard & Tournament Manager - Google Apps Script Backend API
 * =========================================================================
 * 
 * 🏸 เมนูพิเศษจะปรากฏขึ้นบน Google Sheets โดยอัตโนมัติ: "🏸 ระบบคะแนนแบดมินตัน"
 * 
 * วิธีการติดตั้ง:
 * 1. สร้าง Google Sheets ใหม่ขึ้นมา 1 ไฟล์
 * 2. ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมดออก แล้ววางโค้ดไฟล์นี้ลงไปทั้งหมด
 * 4. กดบันทึก (Save) 💾
 * 5. เลือกฟังก์ชัน "setupSheets" ที่แถบด้านบน แล้วกดปุ่ม "Run" (เรียกใช้) 1 ครั้ง
 * 6. กด Deploy (การทำให้ใช้งานได้) > New deployment (การทำให้ใช้งานได้รายการใหม่)
 * 7. เลือกประเภทเป็น "Web app" (เว็บแอป)
 * 8. ตั้งค่า:
 *    - Execute as: Me (ฉัน)
 *    - Who has access: Anyone (ทุกคน) **สำคัญมาก**
 * 9. กด Deploy และคัดลอก "Web app URL" ไปใส่ใน DEFAULT_API_URL ของไฟล์ html ต่างๆ
 */

// เมนูลัดบน Google Sheets เมื่อเปิดสเปรดชีต
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏸 ระบบคะแนนแบดมินตัน')
    .addItem('⚡ ติดตั้ง / รีเซ็ตโครงสร้างชีต (Setup Sheets)', 'setupSheets')
    .addItem('🔄 รีเซ็ตแมตช์ปัจจุบันเป็น 0-0', 'resetCurrentMatchFromMenu')
    .addItem('📋 ใส่ข้อมูลตัวอย่างตารางแข่ง (Add Sample Matches)', 'addSampleData')
    .addSeparator()
    .addItem('📁 ตั้งค่าโฟลเดอร์รูป Drive (Match ID + A,B)', 'promptSetDriveFolder')
    .addItem('🔄 ซิงค์รูปภาพ Drive ทันที (Refresh Photos)', 'refreshDrivePhotosFromMenu')
    .addItem('🔍 ตรวจสอบการเชื่อมต่อรูปภาพ Drive', 'testDrivePhotosMenu')
    .addToUi();
}

/**
 * ฟังก์ชันสร้างและจัดรูปแบบชีตทั้ง 3 แท็บอัตโนมัติ
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // -------------------------------------------------------------
  // 1. แท็บ LiveMatch
  // -------------------------------------------------------------
  var liveSheet = ss.getSheetByName("LiveMatch");
  if (!liveSheet) {
    liveSheet = ss.insertSheet("LiveMatch");
  } else {
    liveSheet.clear();
  }

  var liveHeaders = [
    "tournament_name", "tournament_logo", "match_id", "match_mode", "points_mode", "match_type", "court",
    "team_a_name", "team_b_name", "player_a1", "player_a2", "player_b1", "player_b2",
    "s1_a", "s1_b", "s2_a", "s2_b", "s3_a", "s3_b",
    "active_set", "score_a", "score_b", "status_banner", "serving_side", "history", "updated_at"
  ];
  liveSheet.appendRow(liveHeaders);
  liveSheet.appendRow([
    "BADMINTON CHAMPIONSHIP 2026", "", "M01", "singles", 21, "ประเภทเดี่ยว มือ S", "COURT 1",
    "TEAM A", "TEAM B", "PLAYER A1", "", "PLAYER B1", "",
    0, 0, 0, 0, 0, 0,
    1, 0, 0, "NONE", "L", "[]", new Date()
  ]);

  var liveHeaderRange = liveSheet.getRange(1, 1, 1, liveHeaders.length);
  liveHeaderRange.setBackground("#0f172a")
                 .setFontColor("#38bdf8")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");
  liveSheet.setFrozenRows(1);
  liveSheet.setTabColor("#2563eb");
  liveSheet.autoResizeColumns(1, liveHeaders.length);

  // -------------------------------------------------------------
  // 2. แท็บ Schedule
  // -------------------------------------------------------------
  var schedSheet = ss.getSheetByName("Schedule");
  if (!schedSheet) {
    schedSheet = ss.insertSheet("Schedule");
  } else {
    schedSheet.clear();
  }

  var schedHeaders = [
    "Match ID", "ประเภทการแข่งขัน", "กติกาแต้ม (15/21)", "คอร์ท", "ทีม A", "ผู้เล่น A1", "ผู้เล่น A2",
    "ทีม B", "ผู้เล่น B1", "ผู้เล่น B2", "สถานะ (Status)", "ผู้ชนะ (Winner)", "ผลคะแนนรวม (Final Score)"
  ];
  schedSheet.appendRow(schedHeaders);

  // ข้อมูลตัวอย่าง
  var sampleSchedule = [
    ["194", "P (RR1)", 15, "COURT 9", "Pipet", "กิตติคม", "ฐณวรรณ หมันดี", "Ron noun", "เพชรช่วง ชูรื่นวงศ์", "กัณตพงค์ ปันสำราญ", "Upcoming", "", ""],
    ["195", "P (RR1)", 15, "COURT 9", "พงษ์พาราน้อย", "ธนกฤต สันบุญเป็ง", "ชนาธิป เดโช", "MG", "ทรงพรรณ แมกตี่", "ภานุภัทร ก่ำ", "Upcoming", "", ""],
    ["198", "DU9 (RR1)", 15, "COURT 9", "แก้วละเมียดดับเจังหวัจป...", "โอยฤดา ทองถนอม", "พิชชา สัมฤทธิ์", "ญาณิสาสมอล", "อภิญญา วิงกาโจ", "ประภากร แก้วทอง", "Upcoming", "", ""]
  ];

  sampleSchedule.forEach(function(row) {
    schedSheet.appendRow(row);
  });

  var schedHeaderRange = schedSheet.getRange(1, 1, 1, schedHeaders.length);
  schedHeaderRange.setBackground("#1e293b")
                  .setFontColor("#facc15")
                  .setFontWeight("bold")
                  .setHorizontalAlignment("center");
  schedSheet.setFrozenRows(1);
  schedSheet.setTabColor("#f59e0b");

  // กฎการตรวจสอบสถานะ (ตั้งค่าแบบไม่อนุญาตให้ error ขวางการเขียน)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Upcoming', 'In Progress', 'Finished'], true)
    .setAllowInvalid(true)
    .build();
  schedSheet.getRange("K2:K1000").setDataValidation(statusRule);

  var ptsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([15, 21], true)
    .setAllowInvalid(true)
    .build();
  schedSheet.getRange("C2:C1000").setDataValidation(ptsRule);

  schedSheet.autoResizeColumns(1, schedHeaders.length);

  // -------------------------------------------------------------
  // 3. แท็บ MatchHistory
  // -------------------------------------------------------------
  var histSheet = ss.getSheetByName("MatchHistory");
  if (!histSheet) {
    histSheet = ss.insertSheet("MatchHistory");
  } else {
    histSheet.clear();
  }

  var histHeaders = [
    "วัน-เวลา (Timestamp)", "Match ID", "ประเภทการแข่งขัน", "กติกาแต้ม", "คอร์ท", "ทีม A", "รายชื่อทีม A",
    "ทีม B", "รายชื่อทีม B", "เซต 1", "เซต 2", "เซต 3", "ผู้ชนะ (Winner)", "ผลคะแนนรวม (Final Score)"
  ];
  histSheet.appendRow(histHeaders);

  var histHeaderRange = histSheet.getRange(1, 1, 1, histHeaders.length);
  histHeaderRange.setBackground("#064e3b")
                 .setFontColor("#34d399")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");
  histSheet.setFrozenRows(1);
  histSheet.setTabColor("#10b981");
  histSheet.autoResizeColumns(1, histHeaders.length);

  var defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheetByName("ชีต1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  Logger.log("✅ ติดตั้งและจัดรูปแบบ Google Sheets ทั้ง 3 แท็บเรียบร้อยแล้ว!");
}

function resetCurrentMatchFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var liveSheet = ss.getSheetByName("LiveMatch");
  if (liveSheet) {
    processResetMatch(liveSheet);
    SpreadsheetApp.getUi().alert("✅ รีเซ็ตคะแนนแมตช์ปัจจุบันเป็น 0-0 เรียบร้อยแล้ว");
  }
}

function addSampleData() {
  setupSheets();
  SpreadsheetApp.getUi().alert("✅ โหลดข้อมูลตัวอย่างและโครงสร้างตารางแข่งขันเรียบร้อยแล้ว");
}

// -------------------------------------------------------------
// Google Drive Match Photos Helper (จับคู่รูปภาพด้วย Match ID + A, B)
// -------------------------------------------------------------
var DEFAULT_DRIVE_FOLDER_ID = "1G2TfyeJhJy0eGiBSWmlMhehG4L2J_Bug"; // โฟลเดอร์ Google Drive รูปภาพนักกีฬา

function extractDriveFolderId(input) {
  if (!input) return "";
  input = input.toString().trim();
  var m = input.match(/folders\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  var m2 = input.match(/id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  if (input.indexOf('/') === -1 && input.length >= 20) return input;
  return input;
}

function getDriveFolderPhotosMap(folderId) {
  folderId = extractDriveFolderId(folderId);
  if (!folderId) return {};

  var cache = CacheService.getScriptCache();
  var cached = cache.get("drive_photos_map_" + folderId);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var map = {};
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext() && count < 500) {
      count++;
      var file = files.next();
      var name = file.getName();
      // ตัดนามสกุลไฟล์ออก เช่น 194_A.png -> 194_A
      var baseName = name.replace(/\.[^/.]+$/, "").toUpperCase().trim();
      var url = "https://lh3.googleusercontent.com/d/" + file.getId();
      map[baseName] = url;
      // แบบ normalized ไม่มีสัญลักษณ์ เช่น 194A
      var normalized = baseName.replace(/[^A-Z0-9]/g, "");
      map[normalized] = url;
      // เก็บแบบมีนามสกุลด้วยเผื่อตรง
      map[name.toUpperCase().trim()] = url;
    }
    // หากพบรูปภาพ ให้แคชไว้ 30 นาที (1800 วินาที)
    if (Object.keys(map).length > 0) {
      try {
        cache.put("drive_photos_map_" + folderId, JSON.stringify(map), 1800);
      } catch(ce) {}
    }
  } catch(err) {
    Logger.log("Error getDriveFolderPhotosMap: " + err);
  }
  return map;
}

function promptSetDriveFolder() {
  var ui = SpreadsheetApp.getUi();
  var curFolder = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
  var resp = ui.prompt(
    "📁 ตั้งค่าโฟลเดอร์รูปภาพนักกีฬา Google Drive",
    "กรุณาใส่ลิงก์โฟลเดอร์ Google Drive หรือ Folder ID ที่เก็บรูปภาพ (เช่น https://drive.google.com/drive/folders/...):\n(โฟลเดอร์ปัจจุบัน: " + (curFolder || "ยังไม่ได้ตั้งค่า") + ")",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() === ui.Button.OK) {
    var rawInput = resp.getResponseText();
    var folderId = extractDriveFolderId(rawInput);
    if (folderId) {
      PropertiesService.getScriptProperties().setProperty("drive_folder_id", folderId);
      CacheService.getScriptCache().remove("drive_photos_map_" + folderId);
      CacheService.getScriptCache().remove("live_match_data");
      var map = getDriveFolderPhotosMap(folderId);
      var total = Object.keys(map).length;
      ui.alert("✅ บันทึกโฟลเดอร์รูปภาพเรียบร้อยแล้ว!\nFolder ID: " + folderId + "\nพบไฟล์รูปภาพในโฟลเดอร์: " + total + " รูป");
    } else {
      ui.alert("⚠️ ไม่พบ ID โฟลเดอร์ที่ถูกต้อง กรุณาตรวจสอบลิงก์อีกครั้งครับ");
    }
  }
}

function refreshDrivePhotosFromMenu() {
  var folderId = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
  if (!folderId) {
    SpreadsheetApp.getUi().alert("⚠️ ยังไม่ได้ตั้งค่าโฟลเดอร์ Google Drive กรุณากดเมนู 'ตั้งค่าโฟลเดอร์รูป Drive' ก่อนครับ");
    return;
  }
  CacheService.getScriptCache().remove("drive_photos_map_" + folderId);
  CacheService.getScriptCache().remove("live_match_data");
  var map = getDriveFolderPhotosMap(folderId);
  var total = Object.keys(map).length;
  SpreadsheetApp.getUi().alert("✅ ซิงค์รูปภาพจาก Google Drive สำเร็จ!\nพบไฟล์รูปภาพทั้งหมด: " + total + " รูป");
}

function testDrivePhotosMenu() {
  var ui = SpreadsheetApp.getUi();
  var folderId = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
  if (!folderId) {
    ui.alert("⚠️ ยังไม่ได้ตั้งค่าโฟลเดอร์ Google Drive");
    return;
  }
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var names = [];
    while (files.hasNext()) {
      names.push(files.next().getName());
    }
    if (names.length > 0) {
      ui.alert("✅ เชื่อมต่อ Drive สำเร็จ 100%!\n📁 ชื่อโฟลเดอร์: " + folder.getName() + "\n🖼️ พบรูปภาพ: " + names.join(", "));
    } else {
      ui.alert("⚠️ เข้าถึงโฟลเดอร์ได้ แต่ยังไม่พบไฟล์รูปภาพข้างในครับ\nชื่อโฟลเดอร์: " + folder.getName());
    }
  } catch(e) {
    ui.alert("❌ เข้าถึง Google Drive ไม่สำเร็จ:\n" + e.toString() + "\n\n💡 คำแนะนำ: กรุณาคลิกขวาที่โฟลเดอร์ใน Drive แล้วเปิดแชร์เป็น 'ทุกคนที่มีลิงก์มีสิทธิ์ดู' และกด Review permissions หากมีแจ้งเตือนครับ");
  }
}

// ฟังก์ชันสำหรับกด Run ใน Apps Script Editor เพื่อทดสอบและยอมรับสิทธิ์ Drive ทันที
function testDrivePhotos() {
  var folderId = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
  Logger.log("Testing folder ID: " + folderId);
  try {
    var folder = DriveApp.getFolderById(folderId);
    Logger.log("✅ โฟลเดอร์ชื่อ: " + folder.getName());
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) {
      count++;
      var f = files.next();
      Logger.log("รูปที่ " + count + ": " + f.getName() + " (ID: " + f.getId() + ")");
    }
    Logger.log("รวมพบไฟล์: " + count + " ไฟล์");
  } catch(e) {
    Logger.log("❌ Error: " + e.toString());
  }
}

// -------------------------------------------------------------
// Request Handlers: รับคำขอ GET / POST จากภายนอก
// -------------------------------------------------------------

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var params = e ? e.parameter : {};
  var action = params.action || "getLive";
  var postData = null;

  if (e && e.postData && e.postData.contents) {
    try {
      postData = JSON.parse(e.postData.contents);
      if (postData.action) action = postData.action;
    } catch (err) {}
  }

  // 1. READ ACTIONS: อ่านจากชีตสดๆ โดยตรง เพื่อป้องกันการคืนค่าคะแนนเก่าจากแคช
  if (action === "getLive") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var liveSheet = ss.getSheetByName("LiveMatch");
      if (!liveSheet) {
        setupSheets();
        liveSheet = ss.getSheetByName("LiveMatch");
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: getLiveMatchData(liveSheet)
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === "testDrive") {
    var fId = extractDriveFolderId(params.folder_id || PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID);
    var res = { folder_id: fId };
    try {
      var folder = DriveApp.getFolderById(fId);
      res.folder_name = folder.getName();
      var files = folder.getFiles();
      var list = [];
      while (files.hasNext()) {
        var f = files.next();
        list.push({ name: f.getName(), id: f.getId() });
      }
      res.success = true;
      res.total = list.length;
      res.files = list;
    } catch(err) {
      res.success = false;
      res.error = err.toString();
    }
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getSchedule") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var schedSheet = ss.getSheetByName("Schedule");
      if (!schedSheet) {
        setupSheets();
        schedSheet = ss.getSheetByName("Schedule");
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: getScheduleData(schedSheet)
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === "getHistory") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var histSheet = ss.getSheetByName("MatchHistory");
      if (!histSheet) {
        setupSheets();
        histSheet = ss.getSheetByName("MatchHistory");
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: getHistoryData(histSheet)
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. WRITE/MUTATION ACTIONS: ใช้ LockService เพื่อป้องกัน Race Condition ในการเขียนชีต
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(5000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var liveSheet = ss.getSheetByName("LiveMatch");
    var schedSheet = ss.getSheetByName("Schedule");
    var histSheet = ss.getSheetByName("MatchHistory");

    if (!liveSheet || !schedSheet || !histSheet) {
      setupSheets();
      liveSheet = ss.getSheetByName("LiveMatch");
      schedSheet = ss.getSheetByName("Schedule");
      histSheet = ss.getSheetByName("MatchHistory");
    }

    var result = { success: true };

    if (action === "updateLive") {
      var dataToUpdate = postData ? postData.data : params;
      updateLiveMatchData(liveSheet, dataToUpdate);
      result.data = getLiveMatchData(liveSheet);
    }
    else if (action === "setPointsMode") {
      var pts = Number(params.points_mode || (postData ? postData.points_mode : 21));
      updateLiveMatchData(liveSheet, { points_mode: pts === 15 ? 15 : 21 });
      result.data = getLiveMatchData(liveSheet);
    }
    else if (action === "addPoint") {
      var side = params.side || (postData ? postData.side : "L");
      result.data = processAddPoint(liveSheet, side);
    }
    else if (action === "undoPoint") {
      result.data = processUndoPoint(liveSheet);
    }
    else if (action === "switchSide") {
      result.data = processSwitchSide(liveSheet);
    }
    else if (action === "setServe") {
      var serveSide = params.side || (postData ? postData.side : "NONE");
      result.data = processSetServe(liveSheet, serveSide);
    }
    else if (action === "resetMatch") {
      result.data = processResetMatch(liveSheet);
    }
    else if (action === "saveScheduleMatch") {
      var matchData = postData ? postData.data : params;
      saveScheduleMatchData(schedSheet, matchData);
      result.data = getScheduleData(schedSheet);
    }
    else if (action === "loadMatch") {
      var matchId = params.match_id || (postData ? postData.match_id : "");
      result.data = loadMatchToLive(schedSheet, liveSheet, matchId);
    }
    else if (action === "importExternalSchedule") {
      var sheetUrl = params.sheet_url || (postData ? postData.sheet_url : "");
      var tabName = params.tab_name || (postData ? postData.tab_name : "Data");
      var courtFilter = params.court_filter || (postData ? postData.court_filter : "");
      var defaultPts = Number(params.points_mode || (postData ? postData.points_mode : 15));
      var importMode = params.import_mode || (postData ? postData.import_mode : "append");
      var clearOld = (importMode === "overwrite" || params.clear_old === "true" || params.clear_old === true);
      
      result.data = importFromExternalSheet(schedSheet, sheetUrl, tabName, courtFilter, defaultPts, clearOld);
    }
    else if (action === "saveResult") {
      result.data = processSaveResult(liveSheet, schedSheet, histSheet);
    }
    else if (action === "setDriveFolder") {
      var fId = extractDriveFolderId(params.folder_id || (postData ? postData.folder_id : ""));
      if (fId) {
        PropertiesService.getScriptProperties().setProperty("drive_folder_id", fId);
        CacheService.getScriptCache().remove("drive_photos_map_" + fId);
        CacheService.getScriptCache().remove("live_match_data");
        var total = 0;
        try {
          var newMap = getDriveFolderPhotosMap(fId);
          total = Object.keys(newMap).length;
        } catch(de) {
          Logger.log("Drive map error: " + de);
        }
        result.folder_id = fId;
        result.total_photos = total;
      } else {
        result.success = false;
        result.error = "Invalid folder ID";
      }
    }
    else if (action === "refreshDrivePhotos") {
      var fId = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
      CacheService.getScriptCache().remove("drive_photos_map_" + fId);
      CacheService.getScriptCache().remove("live_match_data");
      var total = 0;
      try {
        var newMap = getDriveFolderPhotosMap(fId);
        total = Object.keys(newMap).length;
      } catch(de) {
        Logger.log("Drive map error: " + de);
      }
      result.folder_id = fId;
      result.total_photos = total;
    }
    else {
      result.success = false;
      result.error = "Unknown action: " + action;
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch(e) {}
    }
  }
}

// -------------------------------------------------------------
// Helper Functions: จัดการข้อมูลในแต่ละแท็บ
// -------------------------------------------------------------

function getLiveMatchData(sheet) {
  var data = sheet.getRange(2, 1, 1, 26).getValues()[0];
  var hist = [];
  try { hist = JSON.parse(data[24] || "[]"); } catch(e) { hist = []; }

  var pointsMode = Number(data[4]);
  if (pointsMode !== 15 && pointsMode !== 21) pointsMode = 21;

  var liveData = {
    tournament_name: data[0] || "",
    tournament_logo: data[1] || "",
    match_id: data[2] || "",
    match_mode: data[3] || "singles",
    points_mode: pointsMode,
    match_type: data[5] || "",
    court: data[6] || "",
    team_a_name: data[7] || "TEAM A",
    team_b_name: data[8] || "TEAM B",
    player_a1: data[9] || "",
    player_a2: data[10] || "",
    player_b1: data[11] || "",
    player_b2: data[12] || "",
    s1_a: Number(data[13]) || 0,
    s1_b: Number(data[14]) || 0,
    s2_a: Number(data[15]) || 0,
    s2_b: Number(data[16]) || 0,
    s3_a: Number(data[17]) || 0,
    s3_b: Number(data[18]) || 0,
    active_set: Number(data[19]) || 1,
    score_a: Number(data[20]) || 0,
    score_b: Number(data[21]) || 0,
    status_banner: data[22] || "NONE",
    serving_side: data[23] || "NONE",
    history: hist,
    updated_at: data[25],
    server_time: Date.now()
  };

  // ดึงรูปภาพนักกีฬาจาก Google Drive ตาม Match ID + A, B
  var folderId = PropertiesService.getScriptProperties().getProperty("drive_folder_id") || DEFAULT_DRIVE_FOLDER_ID || "";
  var photosMap = getDriveFolderPhotosMap(folderId);
  var mId = (liveData.match_id || "").toString().trim().toUpperCase();
  var mIdClean = mId.replace(/[^A-Z0-9]/g, "");

  liveData.photo_a = photosMap[mId + "_A"] || photosMap[mIdClean + "A"] || photosMap[mId + "_A.JPG"] || photosMap[mId + "_A.PNG"] || "";
  liveData.photo_b = photosMap[mId + "_B"] || photosMap[mIdClean + "B"] || photosMap[mId + "_B.JPG"] || photosMap[mId + "_B.PNG"] || "";
  liveData.photo_a1 = photosMap[mId + "_A1"] || photosMap[mIdClean + "A1"] || "";
  liveData.photo_a2 = photosMap[mId + "_A2"] || photosMap[mIdClean + "A2"] || "";
  liveData.photo_b1 = photosMap[mId + "_B1"] || photosMap[mIdClean + "B1"] || "";
  liveData.photo_b2 = photosMap[mId + "_B2"] || photosMap[mIdClean + "B2"] || "";
  liveData.drive_folder_id = folderId;

  // ลบแคช live_match_data เผื่อมีค้างอยู่ เพื่อให้อ่านคะแนนสดตรงจากชีตเสมอ
  try {
    CacheService.getScriptCache().remove("live_match_data");
  } catch(e) {}

  return liveData;
}

function updateLiveMatchData(sheet, d) {
  var cur = getLiveMatchData(sheet);
  var merged = Object.assign({}, cur, d);

  var pts = Number(merged.points_mode);
  if (pts !== 15 && pts !== 21) pts = 21;
  
  merged.points_mode = pts;
  merged.server_time = Date.now();
  merged.updated_at = new Date();

  // เคลียร์แคชเก่าทิ้งทันที
  try {
    CacheService.getScriptCache().remove("live_match_data");
  } catch(e) {}

  var rowData = [
    merged.tournament_name, merged.tournament_logo, merged.match_id, merged.match_mode, pts, merged.match_type, merged.court,
    merged.team_a_name, merged.team_b_name, merged.player_a1, merged.player_a2, merged.player_b1, merged.player_b2,
    merged.s1_a, merged.s1_b, merged.s2_a, merged.s2_b, merged.s3_a, merged.s3_b,
    merged.active_set, merged.score_a, merged.score_b, merged.status_banner, merged.serving_side,
    typeof merged.history === 'string' ? merged.history : JSON.stringify(merged.history),
    merged.updated_at
  ];
  sheet.getRange(2, 1, 1, 26).setValues([rowData]);
  SpreadsheetApp.flush(); // บังคับเขียนลงสเปรดชีตทันที ป้องกันอ่านค่าเก่า
}

function checkSetWon(a, b, pointsMode) {
  var target = (pointsMode === 15) ? 15 : 21;
  var maxCap = (pointsMode === 15) ? 21 : 30;

  if (a >= maxCap) return 'A';
  if (b >= maxCap) return 'B';
  if (a >= target && a - b >= 2) return 'A';
  if (b >= target && b - a >= 2) return 'B';
  return null;
}

function calculateBanner(scoreA, scoreB, setsWonA, setsWonB, isSetOver, isMatchOver, pointsMode) {
  if (isMatchOver) return "MATCH OVER";
  if (isSetOver) return "SET OVER";
  
  var target = (pointsMode === 15) ? 15 : 21;
  var deuceTrigger = target - 1;
  var maxCap = (pointsMode === 15) ? 21 : 30;

  var deuce = scoreA >= deuceTrigger && scoreB >= deuceTrigger && scoreA !== maxCap && scoreB !== maxCap;
  if (deuce) return "DEUCE";

  var isGP_A = (scoreA >= deuceTrigger && scoreA > scoreB);
  var isGP_B = (scoreB >= deuceTrigger && scoreB > scoreA);

  if (isGP_A) {
    return (setsWonA === 1) ? "MATCH POINT" : "GAME POINT";
  }
  if (isGP_B) {
    return (setsWonB === 1) ? "MATCH POINT" : "GAME POINT";
  }
  return "NONE";
}

function processAddPoint(sheet, side) {
  var live = getLiveMatchData(sheet);
  if (live.status_banner === "MATCH OVER") return live;

  var ptsMode = live.points_mode || 21;

  var snapshot = {
    s1_a: live.s1_a, s1_b: live.s1_b,
    s2_a: live.s2_a, s2_b: live.s2_b,
    s3_a: live.s3_a, s3_b: live.s3_b,
    active_set: live.active_set,
    score_a: live.score_a, score_b: live.score_b,
    status_banner: live.status_banner,
    serving_side: live.serving_side,
    points_mode: live.points_mode
  };
  live.history.push(snapshot);
  if (live.history.length > 50) live.history.shift();

  if (side === "L" || side === "A") {
    live.score_a++;
    live.serving_side = "L";
  } else {
    live.score_b++;
    live.serving_side = "R";
  }

  var setsWonA = 0;
  var setsWonB = 0;
  if (live.active_set >= 2) {
    if (checkSetWon(live.s1_a, live.s1_b, ptsMode) === 'A') setsWonA++;
    if (checkSetWon(live.s1_a, live.s1_b, ptsMode) === 'B') setsWonB++;
  }
  if (live.active_set >= 3) {
    if (checkSetWon(live.s2_a, live.s2_b, ptsMode) === 'A') setsWonA++;
    if (checkSetWon(live.s2_a, live.s2_b, ptsMode) === 'B') setsWonB++;
  }

  var winner = checkSetWon(live.score_a, live.score_b, ptsMode);
  var isMatchOver = false;

  if (winner) {
    if (winner === 'A') setsWonA++;
    if (winner === 'B') setsWonB++;

    if (live.active_set === 1) {
      live.s1_a = live.score_a;
      live.s1_b = live.score_b;
    } else if (live.active_set === 2) {
      live.s2_a = live.score_a;
      live.s2_b = live.score_b;
    } else if (live.active_set === 3) {
      live.s3_a = live.score_a;
      live.s3_b = live.score_b;
    }

    if (setsWonA >= 2 || setsWonB >= 2 || live.active_set === 3) {
      isMatchOver = true;
      live.status_banner = "MATCH OVER";
    } else {
      live.active_set++;
      live.score_a = 0;
      live.score_b = 0;
      live.status_banner = "NONE";
    }
  } else {
    live.status_banner = calculateBanner(live.score_a, live.score_b, setsWonA, setsWonB, false, false, ptsMode);
  }

  updateLiveMatchData(sheet, live);
  return live;
}

function processUndoPoint(sheet) {
  var live = getLiveMatchData(sheet);
  if (!live.history || live.history.length === 0) return live;

  var lastState = live.history.pop();
  Object.assign(live, lastState);

  updateLiveMatchData(sheet, live);
  return live;
}

function processSwitchSide(sheet) {
  var live = getLiveMatchData(sheet);
  
  var tempTeam = live.team_a_name; live.team_a_name = live.team_b_name; live.team_b_name = tempTeam;
  var tempP1 = live.player_a1; live.player_a1 = live.player_b1; live.player_b1 = tempP1;
  var tempP2 = live.player_a2; live.player_a2 = live.player_b2; live.player_b2 = tempP2;

  var t1 = live.s1_a; live.s1_a = live.s1_b; live.s1_b = t1;
  var t2 = live.s2_a; live.s2_a = live.s2_b; live.s2_b = t2;
  var t3 = live.s3_a; live.s3_a = live.s3_b; live.s3_b = t3;

  var curT = live.score_a; live.score_a = live.score_b; live.score_b = curT;

  if (live.serving_side === "L") live.serving_side = "R";
  else if (live.serving_side === "R") live.serving_side = "L";

  updateLiveMatchData(sheet, live);
  return live;
}

function processSetServe(sheet, side) {
  var live = getLiveMatchData(sheet);
  live.serving_side = side;
  updateLiveMatchData(sheet, live);
  return live;
}

function processResetMatch(sheet) {
  var live = getLiveMatchData(sheet);
  live.s1_a = 0; live.s1_b = 0;
  live.s2_a = 0; live.s2_b = 0;
  live.s3_a = 0; live.s3_b = 0;
  live.active_set = 1;
  live.score_a = 0;
  live.score_b = 0;
  live.status_banner = "NONE";
  live.history = [];
  updateLiveMatchData(sheet, live);
  return live;
}

function getScheduleData(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var pts = Number(r[2]);
    if (pts !== 15 && pts !== 21) pts = 21;

    list.push({
      row_index: i + 1,
      match_id: String(r[0]),
      match_type: String(r[1] || ""),
      points_mode: pts,
      court: String(r[3] || ""),
      team_a_name: String(r[4] || ""),
      player_a1: String(r[5] || ""),
      player_a2: String(r[6] || ""),
      team_b_name: String(r[7] || ""),
      player_b1: String(r[8] || ""),
      player_b2: String(r[9] || ""),
      status: String(r[10] || "Upcoming"),
      winner: String(r[11] || ""),
      final_score: String(r[12] || "")
    });
  }
  return list;
}

function saveScheduleMatchData(sheet, d) {
  var rows = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(d.match_id).trim()) {
      foundRow = i + 1;
      break;
    }
  }

  var pts = Number(d.points_mode);
  if (pts !== 15 && pts !== 21) pts = 21;

  var rowData = [
    String(d.match_id || ("M" + (rows.length < 10 ? "0" : "") + rows.length)),
    String(d.match_type || "ประเภททั่วไป"),
    pts,
    String(d.court || "COURT 1"),
    String(d.team_a_name || "TEAM A"),
    String(d.player_a1 || ""),
    String(d.player_a2 || ""),
    String(d.team_b_name || "TEAM B"),
    String(d.player_b1 || ""),
    String(d.player_b2 || ""),
    String(d.status || "Upcoming"),
    String(d.winner || ""),
    String(d.final_score || "")
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, 13).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function loadMatchToLive(schedSheet, liveSheet, matchId) {
  var schedule = getScheduleData(schedSheet);
  var target = null;
  for (var i = 0; i < schedule.length; i++) {
    if (String(schedule[i].match_id).trim() === String(matchId).trim()) {
      target = schedule[i];
      schedSheet.getRange(target.row_index, 11).setValue("In Progress");
      break;
    }
  }

  if (!target) throw new Error("Match ID not found: " + matchId);

  var live = getLiveMatchData(liveSheet);
  var isDoubles = (target.player_a2 || target.player_b2) ? "doubles" : "singles";

  live.match_id = target.match_id;
  live.match_type = target.match_type;
  live.points_mode = target.points_mode || 21;
  live.court = target.court;
  live.match_mode = isDoubles;
  live.team_a_name = target.team_a_name;
  live.player_a1 = target.player_a1;
  live.player_a2 = target.player_a2;
  live.team_b_name = target.team_b_name;
  live.player_b1 = target.player_b1;
  live.player_b2 = target.player_b2;

  live.s1_a = 0; live.s1_b = 0;
  live.s2_a = 0; live.s2_b = 0;
  live.s3_a = 0; live.s3_b = 0;
  live.active_set = 1;
  live.score_a = 0;
  live.score_b = 0;
  live.status_banner = "NONE";
  live.serving_side = "L";
  live.history = [];

  updateLiveMatchData(liveSheet, live);
  return live;
}

// -------------------------------------------------------------
// นำเข้าข้อมูลตารางแข่งและนักกีฬาจาก Google Sheet ภายนอก (แท็บ Data)
// -------------------------------------------------------------
function importFromExternalSheet(schedSheet, externalUrl, tabName, courtFilter, defaultPts, clearOld) {
  if (!externalUrl) throw new Error("กรุณาระบุ URL ของ Google Sheet ภายนอก");

  // สกัด Spreadsheet ID จาก URL
  var idMatch = externalUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  var externalId = idMatch ? idMatch[1] : externalUrl;

  var extSS = null;
  try {
    extSS = SpreadsheetApp.openById(externalId);
  } catch (err) {
    throw new Error("ไม่สามารถเปิด Google Sheet ได้ กรุณาตรวจสอบว่าได้แชร์สิทธิ์เป็น 'ทุกคนที่มีลิงก์มีสิทธิ์อ่าน (Anyone with link can view)'");
  }

  // หาแท็บข้อมูล
  var extSheet = null;
  if (tabName) {
    extSheet = extSS.getSheetByName(tabName);
  }
  if (!extSheet) {
    var gidMatch = externalUrl.match(/gid=([0-9]+)/);
    if (gidMatch) {
      var targetGid = Number(gidMatch[1]);
      var allSheets = extSS.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        if (allSheets[s].getSheetId() === targetGid) {
          extSheet = allSheets[s];
          break;
        }
      }
    }
  }
  if (!extSheet) {
    extSheet = extSS.getSheets()[0];
  }

  var values = extSheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { count: 0, message: "ไม่พบข้อมูลในแท็บ " + extSheet.getName() };
  }

  // 1. ค้นหาแถวที่เป็น Header (ตรวจสอบแถว 0, 1, 2)
  var headerRowIdx = 0;
  for (var r = 0; r < Math.min(3, values.length); r++) {
    var rowStr = values[r].map(function(c) { return String(c).trim().toLowerCase(); }).join(" ");
    if (rowStr.indexOf("match") >= 0 || rowStr.indexOf("court") >= 0 || rowStr.indexOf("player") >= 0 || rowStr.indexOf("นักกีฬา") >= 0) {
      headerRowIdx = r;
      break;
    }
  }

  var headers = values[headerRowIdx].map(function(h) { return String(h).trim(); });
  var headersLower = headers.map(function(h) { return h.toLowerCase(); });

  // 2. ระบุตำแหน่งคอลัมน์แบบไดนามิกและแม่นยำ
  // ค่าเริ่มต้นอิงตามโครงสร้างมาตรฐานตารางแบดมินตัน:
  // Col A=Match(0), Col B=Court(1), Col H=Type(7), Col I=Round(8), Col J=Group(9), 
  // Col K=Team1(10), Col L=Player1(11), Col M=Player2(12), Col R=Team2(17), Col S=Player1(18), Col T=Player2(19)
  var idxMatchId = 0;
  var idxCourt = 1;
  var idxType = 7;
  var idxRound = 8;
  var idxGroupA = 9;
  var idxTeamA = 10;
  var idxPlayerA1 = 11;
  var idxPlayerA2 = 12;
  var idxGroupB = 16;
  var idxTeamB = 17;
  var idxPlayerB1 = 18;
  var idxPlayerB2 = 19;

  // ค้นหาคอลัมน์ Team และ Player ที่มีซ้ำ 2 ชุด
  var teamCols = [];
  var player1Cols = [];
  var player2Cols = [];

  for (var c = 0; c < headersLower.length; c++) {
    var h = headersLower[c];
    if (h === "match" || h === "แมตช์" || h === "คู่ที่" || h === "ลำดับ" || h === "no." || h === "id") {
      idxMatchId = c;
    } else if (h === "court" || h === "คอร์ท" || h === "สนาม" || h === "คอร์ต") {
      idxCourt = c;
    } else if (h === "type" || h === "ประเภท" || h === "รุ่น") {
      idxType = c;
    } else if (h === "round" || h === "รอบ") {
      idxRound = c;
    }

    if (h === "team" || h === "ทีม" || h === "สังกัด" || h === "club" || h === "team a" || h === "team 1") {
      teamCols.push(c);
    }
    if (h === "player 1" || h === "player1" || h === "นักกีฬา 1" || h === "ผู้เล่น 1" || h === "player a1") {
      player1Cols.push(c);
    }
    if (h === "player 2" || h === "player2" || h === "นักกีฬา 2" || h === "ผู้เล่น 2" || h === "player a2") {
      player2Cols.push(c);
    }
  }

  if (teamCols.length >= 2) {
    idxTeamA = teamCols[0];
    idxTeamB = teamCols[1];
  }
  if (player1Cols.length >= 2) {
    idxPlayerA1 = player1Cols[0];
    idxPlayerB1 = player1Cols[1];
  }
  if (player2Cols.length >= 2) {
    idxPlayerA2 = player2Cols[0];
    idxPlayerB2 = player2Cols[1];
  }

  // 3. จัดการล้างตารางกรณี overwrite
  try {
    var maxRows = schedSheet.getMaxRows();
    if (maxRows > 1) {
      schedSheet.getRange(2, 1, maxRows - 1, 13).clearDataValidations();
    }
  } catch(e) {}

  if (clearOld) {
    var lastRow = schedSheet.getLastRow();
    if (lastRow > 1) {
      schedSheet.getRange(2, 1, lastRow - 1, 13).clearContent();
    }
  }

  var targetFilter = courtFilter ? String(courtFilter).trim().toLowerCase() : "";
  var filterNumbersOnly = targetFilter.replace(/[^0-9a-zA-Z]/g, '');
  var isFilterAll = (!targetFilter || targetFilter === "all" || targetFilter === "ทั้งหมด");

  var importedRows = [];
  var startRow = headerRowIdx + 1;

  for (var r = startRow; r < values.length; r++) {
    var row = values[r];
    if (!row || row.length === 0) continue;

    // ต้องมี Match ID
    var mIdVal = row[idxMatchId];
    if (mIdVal === undefined || mIdVal === null || String(mIdVal).trim() === "") continue;
    var mId = String(mIdVal).trim();

    // คอร์ท
    var rawCourt = (idxCourt >= 0 && row[idxCourt] !== undefined) ? String(row[idxCourt]).trim() : "";
    var courtLower = rawCourt.toLowerCase();
    var courtDigits = courtLower.replace(/[^0-9a-zA-Z]/g, '');

    // ตรวจสอบตัวกรองคอร์ท (Court Filter)
    if (!isFilterAll) {
      // ถ้าเลือกกรองคอร์ท แต่แถวนี้ไม่มีการระบุคอร์ท ให้ข้ามทันที
      if (!rawCourt || courtLower === "") {
        continue;
      }

      var matchFound = false;
      if (courtLower === targetFilter) matchFound = true;
      else if (courtDigits && filterNumbersOnly && courtDigits === filterNumbersOnly) matchFound = true;
      else if (courtLower.indexOf(targetFilter) >= 0 || targetFilter.indexOf(courtLower) >= 0) matchFound = true;

      if (!matchFound) {
        continue; // ไม่ตรงกับคอร์ทที่เลือก ข้ามไป
      }
    }

    // ประมวลผลชื่อและประเภท
    var mType = (idxType >= 0 && row[idxType]) ? String(row[idxType]).trim() : "ประเภททั่วไป";
    var mRound = (idxRound >= 0 && row[idxRound]) ? String(row[idxRound]).trim() : "";
    var displayType = mType + (mRound ? (" (" + mRound + ")") : "");

    // ดึงชื่อทีมและนักกีฬา
    var mTeamA = (idxTeamA >= 0 && row[idxTeamA]) ? String(row[idxTeamA]).trim() : "TEAM A";
    var mPA1 = (idxPlayerA1 >= 0 && row[idxPlayerA1]) ? String(row[idxPlayerA1]).trim() : "";
    var mPA2 = (idxPlayerA2 >= 0 && row[idxPlayerA2]) ? String(row[idxPlayerA2]).trim() : "";

    var mTeamB = (idxTeamB >= 0 && row[idxTeamB]) ? String(row[idxTeamB]).trim() : "TEAM B";
    var mPB1 = (idxPlayerB1 >= 0 && row[idxPlayerB1]) ? String(row[idxPlayerB1]).trim() : "";
    var mPB2 = (idxPlayerB2 >= 0 && row[idxPlayerB2]) ? String(row[idxPlayerB2]).trim() : "";

    // ข้ามกรณีที่ไม่มีชื่อนักกีฬาเลย
    if (!mPA1 && !mPB1 && mTeamA === "TEAM A" && mTeamB === "TEAM B") continue;

    var formattedCourt = rawCourt ? (rawCourt.toUpperCase().indexOf("COURT") >= 0 ? rawCourt.toUpperCase() : ("COURT " + rawCourt)) : (courtFilter ? ("COURT " + courtFilter) : "COURT 1");
    var mPts = (defaultPts === 15 || defaultPts === 21) ? defaultPts : 15;

    importedRows.push([
      mId,
      displayType,
      mPts,
      formattedCourt,
      mTeamA || "TEAM A",
      mPA1,
      mPA2,
      mTeamB || "TEAM B",
      mPB1,
      mPB2,
      "Upcoming",
      "",
      ""
    ]);
  }

  var addedCount = 0;
  var updatedCount = 0;

  // 4. เขียนแถวที่นำเข้าลงในชีต Schedule
  if (importedRows.length > 0) {
    if (clearOld) {
      // โหมดแทนที่ทั้งหมด (Overwrite): เขียนเริ่มจากแถว 2
      schedSheet.getRange(2, 1, importedRows.length, 13).setValues(importedRows);
      addedCount = importedRows.length;
    } else {
      // โหมดเพิ่มต่อท้ายและอัปเดต (Append & Merge): ไม่ลบของเก่า
      var existingData = schedSheet.getDataRange().getValues();
      var existingMap = {}; // match_id -> { rowNum, data }
      for (var ex = 1; ex < existingData.length; ex++) {
        var exId = String(existingData[ex][0]).trim();
        if (exId) {
          existingMap[exId] = {
            rowNum: ex + 1,
            data: existingData[ex]
          };
        }
      }

      var rowsToAppend = [];
      for (var imp = 0; imp < importedRows.length; imp++) {
        var item = importedRows[imp];
        var targetId = item[0];

        if (existingMap[targetId]) {
          // มี Match ID นี้อยู่แล้วในตารางเดิม: ให้อัปเดตข้อมูล แต่คงสถานะเดิม (In Progress / Finished) ไว้
          var exObj = existingMap[targetId];
          var exStatus = String(exObj.data[10] || "Upcoming").trim();
          var exWinner = String(exObj.data[11] || "").trim();
          var exFinalScore = String(exObj.data[12] || "").trim();

          // คงสถานะเดิมไว้หากมีการเริ่มแข่งหรือแข่งเสร็จแล้ว
          item[10] = exStatus || "Upcoming";
          item[11] = exWinner;
          item[12] = exFinalScore;

          schedSheet.getRange(exObj.rowNum, 1, 1, 13).setValues([item]);
          updatedCount++;
        } else {
          // เป็น Match ID ใหม่: เพิ่มต่อท้าย
          rowsToAppend.push(item);
          addedCount++;
        }
      }

      if (rowsToAppend.length > 0) {
        var currentLast = schedSheet.getLastRow();
        schedSheet.getRange(currentLast + 1, 1, rowsToAppend.length, 13).setValues(rowsToAppend);
      }
    }

    // ตั้งค่า Data Validation แบบปลอดภัย (allowInvalid = true)
    try {
      var totalRows = schedSheet.getLastRow();
      if (totalRows > 1) {
        var statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Upcoming', 'In Progress', 'Finished'], true)
          .setAllowInvalid(true)
          .build();
        schedSheet.getRange(2, 11, totalRows - 1, 1).setDataValidation(statusRule);
      }
    } catch(e) {}
  }

  return {
    success: true,
    imported_count: importedRows.length,
    added_count: addedCount,
    updated_count: updatedCount,
    mode: clearOld ? "overwrite" : "append",
    tab_used: extSheet.getName(),
    court_filter: courtFilter || "ทั้งหมด",
    schedule: getScheduleData(schedSheet)
  };
}

function processSaveResult(liveSheet, schedSheet, histSheet) {
  var live = getLiveMatchData(liveSheet);
  var ptsMode = live.points_mode || 21;
  
  var setResults = [];
  var winsA = 0;
  var winsB = 0;

  var s1W = checkSetWon(live.s1_a, live.s1_b, ptsMode);
  if (s1W) {
    setResults.push(live.s1_a + "-" + live.s1_b);
    if (s1W === 'A') winsA++; else winsB++;
  }
  var s2W = checkSetWon(live.s2_a, live.s2_b, ptsMode);
  if (s2W) {
    setResults.push(live.s2_a + "-" + live.s2_b);
    if (s2W === 'A') winsA++; else winsB++;
  }
  var s3W = checkSetWon(live.s3_a, live.s3_b, ptsMode);
  if (s3W) {
    setResults.push(live.s3_a + "-" + live.s3_b);
    if (s3W === 'A') winsA++; else winsB++;
  }

  if (setResults.length === 0) {
    setResults.push(live.score_a + "-" + live.score_b);
    if (live.score_a > live.score_b) winsA++; else if (live.score_b > live.score_a) winsB++;
  }

  var winnerName = winsA > winsB ? live.team_a_name : (winsB > winsA ? live.team_b_name : "เสมอ/ยังไม่จบ");
  var finalScoreStr = winsA + " - " + winsB + " เซต (" + setResults.join(", ") + ")";

  var playersA = live.player_a1 + (live.player_a2 ? (" / " + live.player_a2) : "");
  var playersB = live.player_b1 + (live.player_b2 ? (" / " + live.player_b2) : "");

  var histRow = [
    Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"),
    live.match_id,
    live.match_type,
    ptsMode + " แต้ม",
    live.court,
    live.team_a_name,
    playersA,
    live.team_b_name,
    playersB,
    live.s1_a + "-" + live.s1_b,
    live.s2_a + "-" + live.s2_b,
    live.s3_a + "-" + live.s3_b,
    winnerName,
    finalScoreStr
  ];
  histSheet.appendRow(histRow);

  if (live.match_id) {
    var schedRows = schedSheet.getDataRange().getValues();
    for (var i = 1; i < schedRows.length; i++) {
      if (String(schedRows[i][0]).trim() === String(live.match_id).trim()) {
        schedSheet.getRange(i + 1, 11).setValue("Finished");
        schedSheet.getRange(i + 1, 12).setValue(winnerName);
        schedSheet.getRange(i + 1, 13).setValue(finalScoreStr);
        break;
      }
    }
  }

  return {
    saved: true,
    winner: winnerName,
    final_score: finalScoreStr
  };
}

function getHistoryData(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  var list = [];
  for (var i = rows.length - 1; i >= 1; i--) {
    var r = rows[i];
    if (!r[0]) continue;
    list.push({
      timestamp: r[0],
      match_id: r[1],
      match_type: r[2],
      points_mode: r[3],
      court: r[4],
      team_a_name: r[5],
      players_a: r[6],
      team_b_name: r[7],
      players_b: r[8],
      set1: r[9],
      set2: r[10],
      set3: r[11],
      winner: r[12],
      final_score: r[13]
    });
  }
  return list;
}
