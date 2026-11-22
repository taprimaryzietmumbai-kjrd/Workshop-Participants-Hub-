/**
 * ZIET Mumbai – Participant Hub
 * OTP-protected personal profile portal. Deploy as a Google Apps Script Web App.
 * Identity is verified against Employee Code + already-registered email in the main sheet.
 */
const SPREADSHEET_ID = "1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE";
const MAIN_SHEET_NAME = "PRTM ADI";
const SUBMISSIONS_SHEET = "Correction Submissions";
const OTP_TTL_SECONDS = 300;
const SESSION_TTL_SECONDS = 900;
const MAX_OTP_ATTEMPTS = 5;


/**
 * Receive public-directory correction requests from the Participants' Hub.
 * Receives public-directory correction requests and updates the matched master participant row.
 */
function doPost(e) {
  let p = null;
  try {
    const rawBody = e && e.postData && e.postData.contents ? String(e.postData.contents) : "";
    const parameterPayload = e && e.parameter ? String(e.parameter.payload || "") : "";
    const payloadText = rawBody.trim().charAt(0) === "{" ? rawBody : parameterPayload;
    if (!payloadText) throw new Error("No request details were received.");
    p = JSON.parse(payloadText);
    if (p.action === "submitArrivalForm") {
      const result = saveArrivalFormSubmission_(p);
      CacheService.getScriptCache().put(
        "arrival-submission-status-" + String(p.requestId || ""),
        JSON.stringify({status:"success", message:result.message, row:result.row, sheetName:result.sheetName}),
        600
      );
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    if (p.action === "saveArrivalPdf") {
      const saved = saveArrivalFormPdf_(p);
      return ContentService.createTextOutput(JSON.stringify(saved)).setMimeType(ContentService.MimeType.JSON);
    }
    const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
    const {ss,sh} = mainSheet_();
    const c = columns_(sh);
    if (!c.name || !c.hindiName || !c.school) throw new Error("The master sheet must contain English name, Hindi name and school columns.");
    const rows = sh.getDataRange().getDisplayValues();
    const oldName = clean(p.previousName);
    const oldSchool = clean(p.previousSchool);
    const oldHindi = clean(p.previousHi);
    const stripEn = v => String(v || "").replace(/^(Dr\.?|Shri|Sh\.?|Smt\.?|Mrs\.?|Ms\.?|Mr\.?|Km\.?)[\s.]*/i, "");
    const stripHi = v => String(v || "").replace(/^(डॉ\.?|श्रीमती|श्री|सुश्री)[\s.]*/, "");
    const nameKey = v => normalized_(stripEn(v));
    const hindiKey = v => normalized_(stripHi(v));
    const nameMatches = [];
    for (let i=1; i<rows.length; i++) {
      const row=rows[i];
      if (nameKey(row[c.name-1])===nameKey(oldName) &&
          hindiKey(row[c.hindiName-1])===hindiKey(oldHindi)) nameMatches.push(i+1);
    }
    // The public directory uses shortened school labels; use school only to resolve duplicate name matches.
    let matches = nameMatches;
    if (matches.length > 1) {
      const schoolMatches = matches.filter(rn =>
        normalized_(rows[rn-1][c.school-1]).includes(normalized_(oldSchool)) ||
        normalized_(oldSchool).includes(normalized_(rows[rn-1][c.school-1]))
      );
      if (schoolMatches.length) matches = schoolMatches;
    }
    if (matches.length !== 1) throw new Error(matches.length
      ? "More than one participant matches the submitted name and Hindi name. No changes were made."
      : "No master-sheet row matched the submitted English and Hindi names. Check the name columns in PRTM ADI."); 
    const rowNumber=matches[0];
    const getOrAddCol=(current,header)=>{
      if(current) return current;
      const col=sh.getLastColumn()+1;
      sh.getRange(1,col).setValue(header);
      return col;
    };
    const salEnCol=getOrAddCol(c.salutationEn,"English Salutation");
    const salHiCol=getOrAddCol(c.salutationHi,"Hindi Salutation");
    const joiningCol=getOrAddCol(c.find(["Year of Joining","Year of Joining in KVS","KVS Joining Year","Sangathan Joining Year","संगठन में कार्यग्रहण का वर्ष"]), "Year of Joining in KVS");
    const year=clean(p.joiningYear);
    if (year && !/^\d{4}$/.test(year)) throw new Error("Year of Joining must be a four-digit year.");
    const updates=[
      [salEnCol,clean(p.salutationEn)],
      [c.name,clean(p.name)],
      [salHiCol,clean(p.salutationHi)],
      [c.hindiName,clean(p.hi)],
      [c.school,clean(p.school)]
    ];
    if(c.schoolHi) updates.push([c.schoolHi,clean(p.schoolHi)]);
    if(year) updates.push([joiningCol,year]);
    updates.forEach(([col,value])=>sh.getRange(rowNumber,col).setValue(value));
    const tabName="Correction Submissions";
    let log=ss.getSheetByName(tabName);
    const headers=["Timestamp","Participant ID","Previous English Salutation","Previous English Name","Previous Hindi Salutation","Previous Hindi Name","Previous School","Previous School Hindi","Previous Year of Joining","Updated English Salutation","Updated English Name","Updated Hindi Salutation","Updated Hindi Name","Updated School","Updated School Hindi","Updated Year of Joining","Status","Master Sheet Row"];
    if(!log){log=ss.insertSheet(tabName);log.appendRow(headers);log.setFrozenRows(1);}
    else if(log.getLastRow()===0){log.appendRow(headers);log.setFrozenRows(1);}
    log.appendRow([new Date(),clean(p.participantId),clean(p.previousSalutationEn),oldName,clean(p.previousSalutationHi),oldHindi,oldSchool,clean(p.previousSchoolHi),clean(p.previousJoiningYear),clean(p.salutationEn),clean(p.name),clean(p.salutationHi),clean(p.hi),clean(p.school),clean(p.schoolHi),year,"Automatically updated master sheet",rowNumber]);
    return ContentService.createTextOutput("Correction saved directly to the master sheet.").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    const message = String(err && err.message || err);
    if (p && p.requestId && p.action === "submitArrivalForm") {
      CacheService.getScriptCache().put(
        "arrival-submission-status-" + String(p.requestId),
        JSON.stringify({status:"error", message:message}),
        600
      );
    }
    if (p && p.requestId && p.action === "saveArrivalPdf") {
      CacheService.getScriptCache().put(
        "arrival-pdf-status-" + String(p.requestId),
        JSON.stringify({status:"error", message:message}),
        600
      );
    }
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e && e.parameter ? String(e.parameter.action || "") : "";
  if (action === "checkSubmission") return checkArrivalSubmissionStatus_(e);
  if (action === "checkDrivePdf") return checkArrivalPdfStatus_(e);
  if (action === "checkDriveAccess") return ContentService.createTextOutput(JSON.stringify(checkArrivalDriveAccess_())).setMimeType(ContentService.MimeType.JSON);
  const page = e && e.parameter ? String(e.parameter.page || "") : "";
  if (page === "participants") {
    return publicParticipantsJsonp_(e);
  }
  if (page === "" || page === "dashboard") {
    const template = HtmlService.createTemplateFromFile("Dashboard");
    template.webAppUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle("ZIET Mumbai — Participants' Dashboard")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  if (page === "profile") {
    return HtmlService.createHtmlOutputFromFile("Profile")
      .setTitle("ZIET Participant Personal Profile")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  if (page === "hostel") {
    const template = HtmlService.createTemplateFromFile("Hostel");
    template.region = String(e && e.parameter && e.parameter.region || "ahmedabad").toLowerCase() === "mumbai" ? "mumbai" : "ahmedabad";
    return template.evaluate()
      .setTitle("Hostel Room Preference / Allotment / Payment Confirmation — ZIET Mumbai")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "certificate") {
    return HtmlService.createHtmlOutputFromFile("Certificate")
      .setTitle("Download CPD Certificate — ZIET Mumbai")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "arrival") {
    const template = HtmlService.createTemplateFromFile("Arrival");
    template.region = String(e && e.parameter && e.parameter.region || "ahmedabad").toLowerCase() === "mumbai" ? "mumbai" : "ahmedabad";
    return template.evaluate()
      .setTitle("Arrival Facilitation & Departure Plan — ZIET Mumbai")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile("Profile")
    .setTitle("ZIET Participant Personal Profile")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}


/**
 * Public directory feed for the GitHub Participants' Hub.
 * Exposes only directory fields; never returns personal/private profile fields.
 * JSONP is used because the caller is a GitHub Pages site on another origin.
 * URL: ?page=participants&callback=participantDirectoryCallback
 */
function publicParticipantsJsonp_(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || "");
  if (!/^[A-Za-z_$][0-9A-Za-z_$\\.]*$/.test(callback)) {
    return ContentService.createTextOutput("Invalid callback.").setMimeType(ContentService.MimeType.TEXT);
  }
  const sh = mainSheet_().sh;
  const data = sh.getDataRange().getDisplayValues();
  if (data.length < 2) return ContentService.createTextOutput(callback + "([]);").setMimeType(ContentService.MimeType.JAVASCRIPT);
  const c = columns_(sh);
  if (!c.name || !c.hindiName || !c.school) {
    throw new Error("Public directory requires English name, Hindi name and Vidyalaya columns in PRTM ADI.");
  }
  // Resolve each value from the first matching header that actually has data.
  // This handles sheets where both Employee Code and Code, or YOJ and
  // Year of Joining in KVS, exist but only one column is populated.
  const headers = data[0];
  function valueFromAliases(row, aliases) {
    const wanted = aliases.map(normalized_);
    for (let col = 0; col < headers.length; col++) {
      if (wanted.includes(normalized_(headers[col]))) {
        const value = String(row[col] || "").trim();
        if (value) return value;
      }
    }
    return "";
  }
  function yearValue(row) {
    const raw = valueFromAliases(row, ["Year of Joining in KVS","Year of Joining","KVS Joining Year","Sangathan Joining Year","Year of Joining in Sangathan","Date of Joining","Joining Date","YOJ","संगठन में कार्यग्रहण का वर्ष"]);
    const match = raw.match(/(19|20)\\d{2}/);
    return match ? match[0] : raw;
  }
  const output = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[c.name - 1] || "").trim();
    const hindiName = String(row[c.hindiName - 1] || "").trim();
    const school = String(row[c.school - 1] || "").trim();
    if (!name && !hindiName && !school) continue;
    const designationCol = c.find(["Designation","Post","Post Held","पदनाम"]);
    output.push({
      name: name,
      hi: hindiName,
      school: school,
      schoolHi: valueFromAliases(row, ["Vidyalaya Name Hindi","Hindi Vidyalaya","School Name Hindi","Vidyalaya Hindi","विद्यालय का नाम","विद्यालय का नाम हिंदी में","विद्यालय का हिंदी नाम","विद्यालय का नाम (हिंदी)","विद्यालय"]),
      employeeCode: valueFromAliases(row, ["Employee Code","Employee ID","Emp Code","Employee No","Code","कर्मचारी कोड"]),
      joiningYear: yearValue(row),
      designation: designationCol ? String(row[designationCol - 1] || "").trim() : "PRT (Music)",
      designationHi: valueFromAliases(row, ["Hindi Designation","Designation Hindi","Post Hindi","पदनाम","हिंदी पदनाम","पदनाम हिंदी"]),
      qualification: valueFromAliases(row, ["Educational Qualification","Education Qualification","Educational Qlfcn.","Edu. Qlfcn.","Edu Qlfn.","EDU. QLFN.","Edu Qlfcn","Edu Qlfcn.","Edu. Qualification","Education","Qualification","Academic Qualification","शैक्षिक योग्यता","शैक्षणिक योग्यता"]),
      salutationEn: c.salutationEn ? String(row[c.salutationEn - 1] || "").trim() : "",
      salutationHi: c.salutationHi ? String(row[c.salutationHi - 1] || "").trim() : ""
    });
  }
  const json = JSON.stringify(output).replace(/</g, "\\u003c");
  return ContentService.createTextOutput(callback + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
}


/** Save arrival and departure planning independently of hostel room preferences. */
function submitArrivalFacilitation(data) {
  return saveArrivalFormSubmission_(data || {});
}


/** Save all Arrival.html fields to the region response sheet; every field is optional. */
function saveArrivalFormSubmission_(payload) {
  payload = payload || {};
  const requestId = String(payload.requestId || Utilities.getUuid()).replace(/[^a-zA-Z0-9_-]/g,"").slice(0,120);
  const isMumbai = String(payload.region || "").toLowerCase().indexOf("mumbai") !== -1;
  const sheetName = isMumbai ? "Arrival Facilitation - Mumbai" : "Arrival Facilitation";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const labels = {
    applicantName:"Full Name",name:"Full Name",employeeCode:"Employee Code",email:"Email",designation:"Designation",
    region:"Region",kvName:"Kendriya Vidyalaya",state:"State / UT",officeAddress:"KV Address",
    officePhone:"KV Office Phone",residencePhone:"Residence Phone",mobile:"Mobile Number",purpose:"Purpose of Visit",
    participant:"Hostel Allottee / Arriving Participant",block:"Hostel Block",room:"Room Number",
    arrivalDate:"Expected Date of Arrival",arrivalTime:"Expected Time of Arrival",actualArrival:"Actual Check-in",
    departureDate:"Expected Date of Departure",departureTime:"Expected Time of Departure",actualCheckout:"Actual Check-out",
    isLocalParticipant:"Local Participant",localHostelOccupied:"Local Participant Hostel Occupied",
    nonLocalHostelOccupied:"Non-local Participant Hostel Occupied",noPaymentReason:"Reason for No Payment",
    arrived:"Whether Arrived",checkedIn:"Checked In to Allotted Room",daysStayed:"Number of Days Stayed",
    roomOccupancy:"Room Occupancy",hostelRatePerDay:"Hostel Rate Per Day",finalHostelAmount:"Final Hostel Amount",
    course:"Course Name",coordinator:"Course Coordinator",facilitationBy:"Facilitation Ensured By",
    facilitationSlot:"Arrival Facilitation Time Slot",remarks:"Remarks",arrivalMode:"Arrival Mode",
    flightTerminal:"Airport Terminal",trainTerminal:"Railway Station",otherStation:"Other Railway Station",
    busArrival:"Bus Arrival Point",arrivalMumbaiTime:"Expected Arrival Time in Mumbai",
    arrivalHostelTime:"Expected Arrival Time at ZIET Hostel",arrivalHostelBlock:"Arrival Hostel Block",
    packedDinner:"Packed Dinner Parcel"
  };
  const values = {};
  Object.keys(payload).forEach(key => {
    if (["action","requestId","participantSignature","shared"].includes(key)) return;
    const value = payload[key];
    if (value === undefined || value === null || typeof value === "object") return;
    let label = labels[key] || key.replace(/([A-Z])/g," $1").replace(/^./,c=>c.toUpperCase()).trim();
    const family = key.match(/^family(Name|Sex|Relation|Age|KvsStatus)(\d+)$/);
    if (family) label = "Family Member " + family[2] + " — " + ({Name:"Name",Sex:"Sex",Relation:"Relationship",Age:"Age",KvsStatus:"KVS Status"}[family[1]]);
    values[label] = String(value).trim();
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let headers = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||"").trim()) : [];
    if (!headers.some(Boolean)) headers = [];
    ["Timestamp","Submission Request ID"].forEach(h=>{if(!headers.includes(h))headers.push(h);});
    Object.keys(values).forEach(h=>{if(!headers.includes(h))headers.push(h);});
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    const requestCol=headers.indexOf("Submission Request ID")+1;
    if(requestId && sh.getLastRow()>1){
      const prior=sh.getRange(2,requestCol,sh.getLastRow()-1,1).getDisplayValues();
      for(let i=0;i<prior.length;i++) if(String(prior[i][0])===requestId)
        return {ok:true,status:"success",message:"This submission was already saved; no duplicate row was created.",row:i+2,sheetName:sheetName};
    }
    const record={"Timestamp":new Date(),"Submission Request ID":requestId};
    Object.keys(values).forEach(h=>record[h]=values[h]);
    sh.appendRow(headers.map(h=>Object.prototype.hasOwnProperty.call(record,h)?record[h]:""));
    return {ok:true,status:"success",message:"Arrival form saved successfully to Google Sheets.",row:sh.getLastRow(),sheetName:sheetName};
  } finally { lock.releaseLock(); }
}
function checkArrivalSubmissionStatus_(e) {
  const callback=String(e&&e.parameter&&e.parameter.callback||"");
  if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) return ContentService.createTextOutput("Invalid callback").setMimeType(ContentService.MimeType.TEXT);
  const id=String(e&&e.parameter&&e.parameter.requestId||"").replace(/[^a-zA-Z0-9_-]/g,"").slice(0,120);
  const raw=id?CacheService.getScriptCache().get("arrival-submission-status-"+id):"";
  const data=raw?JSON.parse(raw):{status:"pending"};
  return ContentService.createTextOutput(callback+"("+JSON.stringify(data).replace(/</g,"\\u003c")+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function checkArrivalDriveAccess_() {
  try { const folder=DriveApp.getFolderById("11V4nKSttjV_knMP-AiO1dC8AIR_oBFnT"); return {status:"success",folder:folder.getName(),effectiveUser:Session.getEffectiveUser().getEmail()}; }
  catch(err) { return {status:"error",message:String(err&&err.message||err),effectiveUser:Session.getEffectiveUser().getEmail()}; }
}


/**
 * Run this function manually once from the Apps Script editor as the deployment owner.
 * It triggers the required Drive, Docs and Mail authorisation and verifies the destination folder.
 */
function authorizeArrivalDriveAccess() {
  const folder = DriveApp.getFolderById("11V4nKSttjV_knMP-AiO1dC8AIR_oBFnT");
  const folderName = folder.getName();
  const temp = DocumentApp.create("ZIET_Arrival_Authorization_Check");
  const docId = temp.getId();
  temp.getBody().appendParagraph("ZIET Mumbai Arrival PDF service authorization check.");
  temp.saveAndClose();
  DriveApp.getFileById(docId).setTrashed(true);
  const mailQuota = MailApp.getRemainingDailyQuota();
  return {
    authorized: true,
    destinationFolder: folderName,
    mailQuotaRemaining: mailQuota,
    message: "Drive, Docs and Mail services are accessible to this Apps Script account."
  };
}

function saveArrivalFormPdf_(payload) {
  const requestId=String(payload.requestId||"").replace(/[^a-zA-Z0-9_-]/g,"").slice(0,100);
  if(!requestId) throw new Error("Missing PDF request ID.");
  const cache=CacheService.getScriptCache(), key="arrival-pdf-status-"+requestId;
  cache.put(key,JSON.stringify({status:"processing"}),600);
  try {
    const fields=Array.isArray(payload.fields)?payload.fields:[];
    const participant=String(payload.participantName||"Participant").trim().slice(0,150);
    const employeeCode=String(payload.employeeCode||"").trim().slice(0,80);
    const workshop=String(payload.workshop||"").trim().slice(0,200);
    const safe=v=>String(v==null?"":v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"").slice(0,2000);
    const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd_HHmm");
    const docName="ZIET_Arrival_Form_"+(employeeCode||participant).replace(/[^a-zA-Z0-9_-]/g,"_")+"_"+stamp;
    const doc=DocumentApp.create(docName), body=doc.getBody();
    body.appendParagraph("KENDRIYA VIDYALAYA SANGATHAN").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("ZONAL INSTITUTE OF EDUCATION & TRAINING, MUMBAI").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("PARTICIPANT ARRIVAL & DEPARTURE FORM").setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph("Participant: "+participant);
    if(employeeCode)body.appendParagraph("Employee Code: "+employeeCode);
    if(workshop)body.appendParagraph("Workshop: "+workshop);
    body.appendParagraph("Generated: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd MMM yyyy, hh:mm a"));
    body.appendHorizontalRule();
    fields.forEach(item=>{if(!item||typeof item!=="object")return;const label=safe(item.label),value=safe(item.value);if(!label||!value)return;body.appendParagraph(label).setBold(true);body.appendParagraph(value);});
    doc.saveAndClose();
    const blob=DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(docName+".pdf");
    const folder=DriveApp.getFolderById("11V4nKSttjV_knMP-AiO1dC8AIR_oBFnT");
    const file=folder.createFile(blob);
    file.setDescription("Completed ZIET Mumbai Arrival & Departure Form for "+participant+(employeeCode?" ("+employeeCode+")":""));
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    const url=file.getUrl();
    MailApp.sendEmail({to:"taprimaryzietmumbai@gmail.com",subject:"ZIET Mumbai — Arrival Form PDF saved to Google Drive — "+participant,
      body:"A participant Arrival & Departure Form PDF has been saved to the designated Google Drive folder.\n\nParticipant: "+participant+"\nEmployee Code: "+employeeCode+"\nWorkshop: "+workshop+"\n\nPDF: "+url+"\n\nDrive folder: https://drive.google.com/drive/folders/11V4nKSttjV_knMP-AiO1dC8AIR_oBFnT"});
    const result={status:"saved",url:url,message:"PDF saved and coordinator notified."};
    cache.put(key,JSON.stringify(result),600);return result;
  } catch(err) { cache.put(key,JSON.stringify({status:"error",message:String(err&&err.message||err)}),600);throw err; }
}
function checkArrivalPdfStatus_(e) {
  const callback=String(e&&e.parameter&&e.parameter.callback||"");
  if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback))return ContentService.createTextOutput("Invalid callback").setMimeType(ContentService.MimeType.TEXT);
  const id=String(e&&e.parameter&&e.parameter.requestId||"").replace(/[^a-zA-Z0-9_-]/g,"").slice(0,100);
  const raw=id?CacheService.getScriptCache().get("arrival-pdf-status-"+id):"";
  const data=raw?JSON.parse(raw):{status:"pending"};
  return ContentService.createTextOutput(callback+"("+JSON.stringify(data).replace(/</g,"\\u003c")+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalized_(v) {
  return String(v == null ? "" : v).toLowerCase().replace(/[^a-z0-9\u0900-\u097f]/g, "");
}
function mainSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sh) throw new Error('Main sheet tab "' + MAIN_SHEET_NAME + '" was not found. Update MAIN_SHEET_NAME in Code.gs.');
  return {ss:ss, sh:sh};
}
function columns_(sh) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const find = aliases => {
    const wanted = aliases.map(normalized_);
    const i = headers.findIndex(h => wanted.includes(normalized_(h)));
    return i < 0 ? 0 : i + 1;
  };
  return {
    headers:headers,
    find:find,
    employee:find(["Employee Code","Employee ID","Emp Code","Employee No","Code","कर्मचारी कोड"]),
    email:find(["Email","email","Email Address","E-mail","Official Email","ईमेल","ईमेल पता"]),
    name:find(["Participant Name","Participant Name English","Name","Teacher Name","Name in English","English Name","Name of Participant","Name of the Participant","Name of Teacher","Name of the Teacher","Teacher's Name","Name (English)","Full Name"]),
    hindiName:find(["Name in Hindi","Hindi Name","Hindi Name of Participant","Participant Name Hindi","Name of Participant in Hindi","Name (Hindi)","Hindi Full Name","प्रतिभागी का नाम","प्रतिभागी का नाम ","नाम हिंदी में","शिक्षक का नाम","हिंदी में नाम","नाम हिंदी"]),
    school:find(["Vidyalaya","School","School Name","Vidyalaya Name","Name of Vidyalaya","Name of the Vidyalaya","Name of KV","Kendriya Vidyalaya","KV Name","School/Vidyalaya","Name of School","Present KV"]),
    schoolHi:find(["Vidyalaya Name Hindi","Hindi Vidyalaya","School Name Hindi","Vidyalaya Hindi","विद्यालय","विद्यालय ","विद्यालय का नाम","विद्यालय का नाम हिंदी में","विद्यालय का हिंदी नाम","विद्यालय का नाम (हिंदी)"]),
    gender:find(["Gender","Sex","लिंग"]),
    category:find(["Social Category","Category","Reservation Category","सामाजिक वर्ग"]),
    disability:find(["Differently Abled","Disability","Divyang","दिव्यांग","दिव्यांगता"]),
    phone:find(["Phone Number","Mobile Number","Phone","Mobile","Contact Number","Whatsapp","WhatsApp","WhatsApp Number","मोबाइल नंबर","फोन नंबर"]),
    salutationEn:find(["English Salutation","Salutation"]),
    salutationHi:find(["Hindi Salutation","हिंदी संबोधन"])
  };
}
function findEmployeeRow_(sh, cols, employeeCode, email) {
  if (!cols.employee || !cols.email) throw new Error("The main sheet must have Employee Code and registered Email columns before OTP access can be used. No personal information was returned.");
  const values = sh.getDataRange().getDisplayValues();
  const ec = String(employeeCode).trim().toLowerCase();
  const em = String(email).trim().toLowerCase();
  for (let r=1;r<values.length;r++) {
    if (String(values[r][cols.employee-1]).trim().toLowerCase()===ec &&
        String(values[r][cols.email-1]).trim().toLowerCase()===em) return r+1;
  }
  return 0;
}
/**
 * Public request form for Employee Code / registered email corrections.
 * It never edits the master participant row; the coordinator verifies and updates it manually.
 */
function submitIdentityCorrectionRequest(data) {
  data = data || {};
  const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
  const name = clean(data.name);
  const school = clean(data.school);
  const currentEmployeeCode = clean(data.currentEmployeeCode);
  const currentEmail = clean(data.currentEmail).toLowerCase();
  const newEmployeeCode = clean(data.newEmployeeCode);
  const newEmail = clean(data.newEmail).toLowerCase();
  const reason = clean(data.reason);
  if (clean(data.website)) return {ok:true,message:"Your request has been received."};
  if (!name || !school || (!newEmployeeCode && !newEmail)) {
    throw new Error("Please enter your full name, Vidyalaya, and the corrected Employee Code or new email.");
  }
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (currentEmail && !emailPattern.test(currentEmail)) throw new Error("Please enter a valid current email address.");
  if (newEmail && !emailPattern.test(newEmail)) throw new Error("Please enter a valid new email address.");
  if (newEmployeeCode && /\s/.test(newEmployeeCode)) throw new Error("Employee Code must not contain spaces.");

  const {ss,sh} = mainSheet_();
  const cols = columns_(sh);
  if (!cols.employee || !cols.email || !cols.name || !cols.school) {
    throw new Error("Required columns were not found in PRTM ADI. Check the header row for Employee Code, Email, Name and Vidyalaya. No changes were made.");
  }
  const rows = sh.getDataRange().getDisplayValues();
  const normalizeName = v => normalized_(String(v || "")
    .replace(/^(Dr\.?|Shri|Sh\.?|Smt\.?|Mrs\.?|Ms\.?|Mr\.?|Km\.?)[\s.]*/i, "")
    .replace(/^(डॉ\.?|श्रीमती|श्री|सुश्री)[\s.]*/, ""));
  const normalizeSchool = v => normalized_(v);
  const wantedName = normalizeName(name);
  const wantedSchool = normalizeSchool(school);
  if (!wantedName || !wantedSchool) throw new Error("Name and Vidyalaya must contain valid text.");

  const matches = [];
  for (let i=1; i<rows.length; i++) {
    const row = rows[i];
    const rowName = normalizeName(row[cols.name-1]);
    const rowSchool = normalizeSchool(row[cols.school-1]);
    const nameMatch = rowName === wantedName;
    const schoolMatch = rowSchool === wantedSchool ||
      (rowSchool.length >= 5 && wantedSchool.length >= 5 &&
       (rowSchool.includes(wantedSchool) || wantedSchool.includes(rowSchool)));
    const codeMatch = !currentEmployeeCode ||
      String(row[cols.employee-1]).trim().toLowerCase() === currentEmployeeCode.toLowerCase();
    const emailMatch = !currentEmail ||
      String(row[cols.email-1]).trim().toLowerCase() === currentEmail;
    if (nameMatch && schoolMatch && codeMatch && emailMatch) matches.push(i+1);
  }
  if (matches.length !== 1) {
    throw new Error(matches.length
      ? "More than one master-sheet record matched. No changes were made. Add the current Employee Code or registered email to identify the record."
      : "No unique record matched in PRTM ADI. Check the spelling of the name and Vidyalaya, and enter the current Employee Code or registered email if available. No changes were made.");
  }
  const rowNumber = matches[0];
  const oldCode = String(rows[rowNumber-1][cols.employee-1] || "").trim();
  const oldEmail = String(rows[rowNumber-1][cols.email-1] || "").trim().toLowerCase();
  const wantedCode = newEmployeeCode;
  const wantedEmail = newEmail;

  if (wantedCode) {
    const duplicate = rows.some((row,i) => i > 0 && i+1 !== rowNumber &&
      String(row[cols.employee-1] || "").trim().toLowerCase() === wantedCode.toLowerCase());
    if (duplicate) throw new Error("That Employee Code already belongs to another participant. No changes were made.");
  }
  if (wantedEmail) {
    const duplicate = rows.some((row,i) => i > 0 && i+1 !== rowNumber &&
      String(row[cols.email-1] || "").trim().toLowerCase() === wantedEmail);
    if (duplicate) throw new Error("That email address already belongs to another participant. No changes were made.");
  }

  const tabName = "Identity Correction Requests";
  let log = ss.getSheetByName(tabName);
  const headers = ["Timestamp","Full Name","Vidyalaya","Current Employee Code","Current Registered Email","Requested Employee Code","Requested New Email","Reason","Status","Master Sheet Row"];
  if (!log) {
    log = ss.insertSheet(tabName);
    log.getRange(1,1,1,headers.length).setValues([headers]);
    log.setFrozenRows(1);
  } else if (log.getLastRow() === 0) {
    log.getRange(1,1,1,headers.length).setValues([headers]);
    log.setFrozenRows(1);
  }

  // Record the attempt before editing, so every accepted request has an audit trail.
  const logRow = log.getLastRow() + 1;
  log.getRange(logRow,1,1,headers.length).setValues([[
    new Date(),name,school,currentEmployeeCode,currentEmail,newEmployeeCode,newEmail,reason,
    "Update in progress",rowNumber
  ]]);
  SpreadsheetApp.flush();

  try {
    if (wantedCode) sh.getRange(rowNumber,cols.employee).setNumberFormat("@").setValue(wantedCode);
    if (wantedEmail) sh.getRange(rowNumber,cols.email).setValue(wantedEmail);
    SpreadsheetApp.flush();
    const check = sh.getRange(rowNumber,1,1,sh.getLastColumn()).getDisplayValues()[0];
    if (wantedCode && String(check[cols.employee-1]).trim() !== wantedCode) {
      throw new Error("Google Sheets did not confirm the Employee Code change.");
    }
    if (wantedEmail && String(check[cols.email-1]).trim().toLowerCase() !== wantedEmail) {
      throw new Error("Google Sheets did not confirm the email change.");
    }
    log.getRange(logRow,9).setValue("Successfully updated master sheet");
    log.getRange(logRow,10).setValue(rowNumber);
    SpreadsheetApp.flush();
    return {ok:true,message:"Correction saved successfully. The master sheet and Identity Correction Requests log have both been updated."};
  } catch (err) {
    log.getRange(logRow,9).setValue("Update failed: " + clean(err.message || err));
    SpreadsheetApp.flush();
    throw new Error("The request was logged, but the master-sheet update could not be confirmed: " + String(err.message || err));
  }
}
function sendProfileOtp(employeeCode, registeredEmail) {
  employeeCode = String(employeeCode || "").trim();
  registeredEmail = String(registeredEmail || "").trim().toLowerCase();
  if (!employeeCode || !registeredEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(registeredEmail)) throw new Error("Enter your Employee Code and registered email address.");
  const cache = CacheService.getScriptCache();
  const {sh} = mainSheet_();
  const cols = columns_(sh);
  const row = findEmployeeRow_(sh, cols, employeeCode, registeredEmail);
  if (!row) throw new Error("The Employee Code and email did not match a registered record. Contact the workshop coordinator to verify your details.");
  const otp = String(Math.floor(100000 + Math.random()*900000));
  const key = "otp-" + Utilities.base64EncodeWebSafe(employeeCode.toLowerCase()+"|"+registeredEmail).slice(0,70);
  cache.put(key, JSON.stringify({hash:Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, otp).map(b=>(b+256).toString(16).slice(-2)).join(""), row:row, attempts:0}), OTP_TTL_SECONDS);
  MailApp.sendEmail({
    to:registeredEmail,
    subject:"Your ZIET Participant Hub verification code",
    body:"Your one-time verification code is: "+otp+"\n\nIt expires in 5 minutes and can be used only once. Do not share it with anyone. If you did not request this code, you can ignore this email."
  });
  return {ok:true,message:"OTP sent to your registered email. Check your inbox and spam folder. The code expires in 5 minutes."};
}
function verifyProfileOtp(employeeCode, registeredEmail, otp) {
  employeeCode=String(employeeCode||"").trim();
  registeredEmail=String(registeredEmail||"").trim().toLowerCase();
  otp=String(otp||"").trim();
  const cache=CacheService.getScriptCache();
  const key="otp-"+Utilities.base64EncodeWebSafe(employeeCode.toLowerCase()+"|"+registeredEmail).slice(0,70);
  const raw=cache.get(key);
  if(!raw) throw new Error("OTP expired or not requested. Request a new code.");
  const state=JSON.parse(raw);
  state.attempts++;
  const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,otp).map(b=>(b+256).toString(16).slice(-2)).join("");  if(hash!==state.hash) {
    if(state.attempts>=MAX_OTP_ATTEMPTS) cache.remove(key); else cache.put(key,JSON.stringify(state),OTP_TTL_SECONDS);
    throw new Error(state.attempts>=MAX_OTP_ATTEMPTS?"Too many incorrect attempts. Request a new OTP.":"Incorrect OTP. Please try again.");
  }
  cache.remove(key);
  const token=Utilities.getUuid()+Utilities.getUuid();
  cache.put("profile-session-"+token,JSON.stringify({row:state.row,employeeCode:employeeCode,email:registeredEmail}),SESSION_TTL_SECONDS);
  return {ok:true,token:token,message:"Verified. Your private profile is available for 15 minutes."};
}
function getPrivateProfile(token) {
  const session=CacheService.getScriptCache().get("profile-session-"+String(token||""));
  if(!session) throw new Error("Your session has expired. Please verify by OTP again.");
  const s=JSON.parse(session);
  const {sh}=mainSheet_(); const c=columns_(sh);
  const row=sh.getRange(s.row,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const val=col=>col?row[col-1]:"";
  return {
    employeeCode:val(c.employee), email:val(c.email), name:val(c.name), hindiName:val(c.hindiName),
    school:val(c.school), schoolHi:val(c.schoolHi), gender:val(c.gender), socialCategory:val(c.category),
    differentlyAbled:val(c.disability), phone:val(c.phone), salutationEn:val(c.salutationEn), salutationHi:val(c.salutationHi)
  };
}
function savePrivateProfile(token, data) {
  const session=CacheService.getScriptCache().get("profile-session-"+String(token||""));
  if(!session) throw new Error("Your session has expired. Please verify by OTP again.");
  const s=JSON.parse(session);
  data=data||{};
  const {ss,sh}=mainSheet_(); const c=columns_(sh);
  const updates=[
    ["name",c.name],["hindiName",c.hindiName],["school",c.school],["schoolHi",c.schoolHi],
    ["gender",c.gender],["socialCategory",c.category],["differentlyAbled",c.disability],["phone",c.phone]
  ];
  if(!c.name||!c.hindiName||!c.school) throw new Error("The main sheet needs English Name, Hindi Name and Vidyalaya columns. Contact the coordinator.");
  // Email and Employee Code are identity keys and are not self-editable here.
  updates.forEach(([key,col])=>{if(col && data[key]!==undefined) sh.getRange(s.row,col).setValue(String(data[key]).trim());});
  if(c.schoolHi && data.schoolHi!==undefined) sh.getRange(s.row,c.schoolHi).setValue(String(data.schoolHi).trim());
  if(c.gender && data.gender!==undefined) sh.getRange(s.row,c.gender).setValue(String(data.gender).trim());
  if(c.category && data.socialCategory!==undefined) sh.getRange(s.row,c.category).setValue(String(data.socialCategory).trim());
  if(c.disability && data.differentlyAbled!==undefined) sh.getRange(s.row,c.disability).setValue(String(data.differentlyAbled).trim());
  if(c.phone && data.phone!==undefined) sh.getRange(s.row,c.phone).setValue(String(data.phone).trim());
  // If the master sheet lacks optional private-data columns, add them with headers.
  const addIfMissing=(key, header, value)=>{
    if(!value) return;
    let col=key==="gender"?c.gender:key==="socialCategory"?c.category:key==="differentlyAbled"?c.disability:key==="phone"?c.phone:key==="schoolHi"?c.schoolHi:0;
    if(!col){col=sh.getLastColumn()+1;sh.getRange(1,col).setValue(header);}
    sh.getRange(s.row,col).setValue(String(value).trim());
  };
  addIfMissing("schoolHi","Vidyalaya Name (Hindi)",data.schoolHi);
  addIfMissing("gender","Gender",data.gender);
  addIfMissing("socialCategory","Social Category",data.socialCategory);
  addIfMissing("differentlyAbled","Differently Abled",data.differentlyAbled);
  addIfMissing("phone","Phone Number",data.phone);
  const log=ss.getSheetByName(SUBMISSIONS_SHEET)||ss.insertSheet(SUBMISSIONS_SHEET);
  if(log.getLastRow()===0) log.appendRow(["Timestamp","Employee Code","Action","Updated Fields"]);
  log.appendRow([new Date(),s.employeeCode,"OTP verified profile update",Object.keys(data).join(", ")]);
  return {ok:true,message:"Your profile corrections have been saved to the main sheet."};
}

/**
 * Share verified contact details with the course coordinator.
 * This does not modify the participant's master record.
 * The participant must explicitly click the sharing button after OTP verification.
 */
function shareContactWithCoordinator(token) {
  const cached = CacheService.getScriptCache().get("profile-session-" + String(token || ""));
  if (!cached) throw new Error("Your session has expired. Please verify by OTP again.");
  const session = JSON.parse(cached);
  const {ss, sh} = mainSheet_();
  const c = columns_(sh);
  const row = sh.getRange(session.row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const val = col => col ? row[col - 1] : "";
  const employeeCode = val(c.employee);
  const email = val(c.email);
  const phone = val(c.phone);
  const name = val(c.name);
  const school = val(c.school);
  if (!employeeCode || !email) throw new Error("Employee Code or registered email is missing from the main sheet. Please contact the coordinator.");
  const logName = "Coordinator Contact Sharing";
  let log = ss.getSheetByName(logName);
  if (!log) {
    log = ss.insertSheet(logName);
    log.appendRow(["Shared At", "Participant Name", "Employee Code", "Registered Email", "Phone Number", "Vidyalaya", "Consent"]);
    log.setFrozenRows(1);
  }
  log.appendRow([new Date(), name, employeeCode, email, phone, school, "Participant explicitly consented after OTP verification"]);
  MailApp.sendEmail({
    to: "taprimaryzietmumbai@gmail.com",
    subject: "Participant contact details shared — ZIET Music Teachers CPD",
    body: "A participant has explicitly chosen to share their contact details with the course coordinator after OTP verification.\n\nParticipant: " + name +
      "\nEmployee Code: " + employeeCode +
      "\nRegistered Email: " + email +
      "\nPhone Number: " + (phone || "Not provided") +
      "\nVidyalaya: " + (school || "Not provided") +
      "\n\nThe participant's main record was not changed by this sharing action. The sharing event is recorded in the '" + logName + "' tab."
  });
  return {ok:true, message:"Your contact details have been shared with the course coordinator and recorded in the Google Sheet. Your master record was not changed."};
}


/** Save hostel preferences and arrival-facilitation details; upload ground-floor evidence when required. */
function submitHostelPreference(data) {
  data = data || {};
  const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
  const region = String(data.region || "ahmedabad").toLowerCase() === "mumbai" ? "Mumbai" : "Ahmedabad";
  const name=clean(data.name), employeeCode=clean(data.employeeCode), email=clean(data.email).toLowerCase();
  const gender=clean(data.gender), block=clean(data.block), floor=clean(data.floor);
  let disability=clean(data.disability);
  const dob=clean(data.dob), groundReason=clean(data.groundReason);
  const arrivalMode=clean(data.arrivalMode), flightTerminal=clean(data.flightTerminal), trainTerminal=clean(data.trainTerminal);
  const otherStation=clean(data.otherStation), busArrival=clean(data.busArrival);
  const arrivalMumbaiTime=clean(data.arrivalMumbaiTime), arrivalHostelTime=clean(data.arrivalHostelTime), arrivalHostelBlock=clean(data.arrivalHostelBlock);
  if(region==="Mumbai") {
    const roster=getMumbaiParticipantOptions_();
    const match=roster.find(p=>String(p.employeeCode||"").trim()===employeeCode && String(p.name||"").trim().toLowerCase()===name.toLowerCase());
    if(!match) throw new Error("Participant record was not found in the Mumbai master roster. Please refresh the roster.");
    disability=formatHostelDisability_(match.disability);
    if(!["Yes","No","Prefer not to say"].includes(disability)) throw new Error("The Mumbai roster does not contain a valid Differently Abled value for this participant. Please contact the coordinator.");
  }
  if(!name || !employeeCode || !email || !gender || !block || !floor || !dob) throw new Error("Please complete all required fields, including date of birth.");
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Enter a valid email address.");
  if(block==="A1" && gender!=="Female") throw new Error("Block A1 is reserved for ladies. Please choose A2 or A3.");
  if(!["A1","A2","A3"].includes(block) || !["Ground Floor","1st Floor","2nd Floor","3rd Floor"].includes(floor)) throw new Error("Please choose valid room preferences.");
  if(!["Train","Flight","Bus"].includes(arrivalMode)) throw new Error("Please select your mode of arrival.");
  if(arrivalMode==="Flight" && !["CSMIA-T1","CSMIA-T2","NMIA"].includes(flightTerminal)) throw new Error("Please select your arrival airport terminal.");
  if(arrivalMode==="Train" && !["Bandra Terminus","Dadar","Mumbai Central","Borivali","CSMT","Thane","Any other"].includes(trainTerminal)) throw new Error("Please select your arriving railway station.");
  if(arrivalMode==="Train" && trainTerminal==="Any other" && !otherStation) throw new Error("Please specify the other railway station.");
  if(arrivalMode==="Bus" && !busArrival) throw new Error("Please enter your bus arrival point.");
  if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(arrivalMumbaiTime) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(arrivalHostelTime)) throw new Error("Please select both expected arrival times.");
  if(!["A1","A2","A3"].includes(arrivalHostelBlock)) throw new Error("Please select the hostel block you expect to arrive at.");
  const birth=new Date(dob+"T00:00:00");
  if(isNaN(birth.getTime()) || birth > new Date()) throw new Error("Please enter a valid date of birth.");
  const age=Math.floor((Date.now()-birth.getTime())/31557600000);

  let fileUrl="";
  if(floor==="Ground Floor") {
    if(!data.document || !data.document.base64 || !data.document.name) throw new Error("Please upload genuine supporting documentation for a ground-floor allotment request.");
    const bytes=Utilities.base64Decode(String(data.document.base64));
    if(bytes.length > 8*1024*1024) throw new Error("The document must be 8 MB or smaller.");
    const mime=clean(data.document.mimeType) || "application/octet-stream";
    const safeName=clean(data.document.name).replace(/[\\/:*?"<>|]/g,"_").slice(0,120);
    const blob=Utilities.newBlob(bytes,mime,safeName);
    const folder=DriveApp.getFolderById("1lYibl8qAWdiRG-fD5Fe2D8WX_Szw14vd");
    const file=folder.createFile(blob);
    file.setDescription("Ground-floor hostel allotment supporting document. Submitted by "+name+" ("+employeeCode+").");
    fileUrl=file.getUrl();
  }
  const ss=SpreadsheetApp.openById("1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE");
  const tab = region === "Mumbai" ? "Hostel Room Preferences - Mumbai" : "Hostel Room Preferences";
  let sh=ss.getSheetByName(tab);
  const headers=["Timestamp","Full Name","Employee Code","Email","Gender for A1 eligibility","Date of Birth","Age at submission","Differently Abled","Preferred Block","Preferred Floor","Preferred Toilet","Ground Floor Reason","Supporting Document URL","Review Status","Arrival Mode","Arrival Airport Terminal","Arrival Railway Station","Other Railway Station","Bus Arrival Point","Expected Arrival Time in Mumbai","Expected Arrival Time at ZIET Hostel","Arrival Hostel Block","Region"];
  if(!sh){sh=ss.insertSheet(tab);sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);}
  let currentHeaders=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getDisplayValues()[0].map(v=>String(v||"").trim());
  headers.forEach(h=>{if(!currentHeaders.includes(h)){currentHeaders.push(h);sh.getRange(1,currentHeaders.length).setValue(h);}});
  const record={
    "Timestamp":new Date(),"Full Name":name,"Employee Code":employeeCode,"Email":email,
    "Gender for A1 eligibility":gender,"Date of Birth":birth,"Age at submission":age,"Differently Abled":disability,
    "Preferred Block":block,"Preferred Floor":floor,"Preferred Toilet":"","Ground Floor Reason":groundReason,
    "Supporting Document URL":fileUrl,"Review Status":"Pending coordinator review","Arrival Mode":arrivalMode,
    "Arrival Airport Terminal":arrivalMode==="Flight"?flightTerminal:"",
    "Arrival Railway Station":arrivalMode==="Train"?trainTerminal:"",
    "Other Railway Station":arrivalMode==="Train"&&trainTerminal==="Any other"?otherStation:"",
    "Bus Arrival Point":arrivalMode==="Bus"?busArrival:"",
    "Expected Arrival Time in Mumbai":arrivalMumbaiTime,
    "Expected Arrival Time at ZIET Hostel":arrivalHostelTime,
    "Arrival Hostel Block":arrivalHostelBlock,"Region":region
  };
  sh.appendRow(currentHeaders.map(h=>Object.prototype.hasOwnProperty.call(record,h)?record[h]:""));
  const travelDetail="\n\nARRIVAL FACILITATION\nMode: "+arrivalMode+
    (arrivalMode==="Flight"?"\nAirport terminal: "+flightTerminal:"")+
    (arrivalMode==="Train"?"\nRailway station: "+trainTerminal+(trainTerminal==="Any other"?" — "+otherStation:""):"")+
    (arrivalMode==="Bus"?"\nBus arrival point: "+busArrival:"")+
    "\nExpected arrival in Mumbai: "+arrivalMumbaiTime+
    "\nExpected arrival at ZIET Hostel: "+arrivalHostelTime+
    "\nHostel block for arrival: "+arrivalHostelBlock;
  MailApp.sendEmail({to:"taprimaryzietmumbai@gmail.com",subject:"Hostel Preference & Arrival Facilitation — "+region+" Region — KV Music Teachers CPD",body:"A hostel preference and arrival facilitation response was submitted. Please verify eligibility and allotment before confirming.\n\nName: "+name+"\nEmployee Code: "+employeeCode+"\nEmail: "+email+"\nDOB: "+dob+" (age "+age+")\nDifferently Abled: "+(disability||"Not specified")+"\nBlock: "+block+"\nFloor: "+floor+"\nGround-floor reason: "+(groundReason||"Not provided")+"\nSupporting document: "+(fileUrl||"Not uploaded")+travelDetail+"\n\nFor further assistance: Anand Srivastava (Coordinator)."});
  return {ok:true,message:"Hostel preference and Arrival Facilitation details submitted successfully. Allotment is subject to coordinator verification. "+(fileUrl ? "Supporting document uploaded to the designated Drive folder." : "")};
}


/** Populate the hostel form from the participant master sheet. */
/** Populate hostel form with participant details from the master sheet. */
/** Return the selected region's roster. Mumbai public roster fields are supplemented by participants in the form. */
function getHostelParticipantOptions(region) {
  region = String(region || "ahmedabad").toLowerCase();
  if (region === "mumbai") {
    // Mumbai roster is stored in a separate, private Apps Script file (MumbaiRoster.gs).
    // Do not fetch participant DOBs, email addresses or phone numbers from the public GitHub directory.
    const roster = getMumbaiParticipantOptions_();
    if (!Array.isArray(roster) || roster.length !== 52) {
      throw new Error("The Mumbai participant roster did not pass the 52-record validation.");
    }
    return roster.map(p => ({
      name: String(p.name || "").trim(),
      employeeCode: String(p.employeeCode || "").trim(),
      email: String(p.email || "").trim(),
      gender: String(p.gender || "").trim(),
      disability: String(p.disability || "").trim(),
      dob: formatHostelDob_(p.dob),
      school: String(p.school || "").trim(),
      region: "Mumbai"
    })).filter(p => p.name);
  }

  // Ahmedabad continues to use the existing PRTM ADI master-sheet workflow.
  const {sh} = mainSheet_();
  const range = sh.getDataRange(), values = range.getValues(), display = range.getDisplayValues();
  const headers = display[0].map(h => String(h || "").toLowerCase().replace(/[^a-z0-9\\u0900-\\u097f]/g, ""));
  function findColumn(aliases) {
    const keys = aliases.map(a => String(a).toLowerCase().replace(/[^a-z0-9\\u0900-\\u097f]/g, ""));
    let index = headers.findIndex(h => keys.includes(h));
    if (index < 0) index = headers.findIndex(h => keys.some(k => h.includes(k)));
    return index;
  }
  const empCol=findColumn(["Employee Code","Employee ID","Emp Code","Employee No","Code"]);
  const nameCol=findColumn(["Participant Name","Name","Teacher Name","Full Name"]);
  const emailCol=findColumn(["Email","Email Address","E-mail","Official Email"]);
  const genderCol=findColumn(["Gender","Sex","लिंग"]);
  const dobCol=findColumn(["Date of Birth","DOB","Birth Date","जन्म तिथि","जन्म दिनांक"]);
  const disabilityCol=findColumn(["Differently Abled","Disability","Divyang","दिव्यांग","दिव्यांगता"]);
  if(empCol<0 || nameCol<0) throw new Error("Employee Code or Participant Name column not found.");
  const out=[];
  for(let i=1;i<values.length;i++) {
    const row=values[i], employeeCode=String(display[i][empCol]||"").trim(), name=String(display[i][nameCol]||"").trim();
    if(!employeeCode||!name) continue;
    out.push({
      employeeCode,
      name,
      email:emailCol>=0?String(display[i][emailCol]||"").trim():"",
      gender:genderCol>=0?String(display[i][genderCol]||"").trim():"",
      disability:disabilityCol>=0?formatHostelDisability_(display[i][disabilityCol]):"",
      dob:dobCol>=0?formatHostelDob_(row[dobCol],display[i][dobCol]):"",
      region:"Ahmedabad"
    });
  }
  return out;
}

/** Normalize roster DOB values to yyyy-MM-dd for the forms. Blank source DOBs stay blank. */
function formatHostelDob_(value, displayValue) {
  if (!value && !displayValue) return "";
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(displayValue || value || "").trim();
  let m = s.match(/^(\\d{1,2})[\\/. -](\\d{1,2})[\\/. -](\\d{4})$/);
  if (m) return m[3] + "-" + String(m[2]).padStart(2,"0") + "-" + String(m[1]).padStart(2,"0");
  m = s.match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);
  return m ? s : "";
}

/** Normalize disability values without inferring missing information. */
function formatHostelDisability_(value) {
  const s = String(value || "").trim();
  if (/^(yes|y|true|1|दिव्यांग)$/i.test(s)) return "Yes";
  if (/^(no|n|false|0|नहीं|नही)$/i.test(s)) return "No";
  if (/prefer not/i.test(s)) return "Prefer not to say";
  return s;
}

/** Refresh button handler used by both hostel and arrival forms. */
function refreshHostelParticipantOptions(region) {
  return getHostelParticipantOptions(region);
}

/* CPD CERTIFICATE ACCESS — email OTP only; no WhatsApp/SMS integration. */
const CPD_CERTIFICATE_FOLDER_ID = "1RmGYUGD1JyyAE1vMxbXMgSfZcB4JhzRL";
const CPD_CERTIFICATE_OTP_TTL = 300;
const CPD_CERTIFICATE_SESSION_TTL = 600;
const CPD_CERTIFICATE_MAX_ATTEMPTS = 5;

/**
 * Find the actual Date of Birth column for certificate identity checks.
 * Prefer explicit DOB headers and deliberately exclude joining-date fields.
 * This prevents a YOJ / Date of Joining column from being used as DOB.
 */
function findCpdDobColumn_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const key = value => normalized_(value);
  const exactAliases = [
    "Date of Birth", "DOB", "D.O.B.", "Birth Date", "Birthdate",
    "Date Of Birth (DD-MM-YYYY)", "जन्म तिथि", "जन्म दिनांक", "जन्म की तारीख"
  ].map(key);
  const excluded = /(joining|join|yoj|appointment|recruitment|नियुक्ति|कार्यग्रहण|संगठनमेंकार्यग्रहण)/i;

  // Exact DOB header match takes precedence. Do not accept generic "Date".
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || "").trim();
    const normalizedHeader = key(header);
    if (exactAliases.includes(normalizedHeader) && !excluded.test(header)) return i + 1;
  }

  // Accept a clearly-labelled DOB header only as a fallback.
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || "").trim();
    if (excluded.test(header)) continue;
    if (/date\s*of\s*birth|\bdob\b|birth\s*date|birthdate|जन्म\s*(तिथि|दिनांक|की\s*तारीख)/i.test(header)) return i + 1;
  }
  return 0;
}

/** Convert a DOB value to DDMMYYYY, preferring the sheet's displayed value. */
function normalizeCpdDob_(raw, display) {
  const s = String(display == null ? "" : display).trim();
  if (s) {
    let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (m) return m[1].padStart(2, "0") + m[2].padStart(2, "0") + m[3];
    m = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
    if (m) return m[3].padStart(2, "0") + m[2].padStart(2, "0") + m[1];
    // A sheet may display a date with a time suffix; retain only the date part.
    m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})\s+/);
    if (m) return m[1].padStart(2, "0") + m[2].padStart(2, "0") + m[3];
  }
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), "ddMMyyyy");
  }
  const fallback = String(raw == null ? "" : raw).trim();
  let m = fallback.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return m[1].padStart(2, "0") + m[2].padStart(2, "0") + m[3];
  m = fallback.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (m) return m[3].padStart(2, "0") + m[2].padStart(2, "0") + m[1];
  return "";
}

function normalizeCpdEmployeeCode_(value) {
  // Employee Codes are identifiers, not quantities. Keep only digits and
  // restore a leading zero if the sheet stores a five-digit code numerically.
  const digits = String(value == null ? "" : value).trim().replace(/\D/g, "");
  return digits && digits.length <= 5 ? digits.padStart(5, "0") : digits;
}

function cpdCertificateIdentity_(employeeCodeDob) {
  const input = String(employeeCodeDob || "").replace(/\s/g, "");
  if (!/^\d{13}$/.test(input)) {
    throw new Error("Enter the 13-digit EmployeeCodeDOB: 5-digit Employee Code followed by DOB in DDMMYYYY format.");
  }
  const employeeCode = input.slice(0, 5);
  const dobInput = input.slice(5);
  const {sh} = mainSheet_();
  const c = columns_(sh);
  const dobCol = findCpdDobColumn_(sh);
  if (!c.employee || !c.email || !c.name || !dobCol) {
    throw new Error("The master sheet must contain Employee Code, registered Email, participant Name and a clearly labelled Date of Birth / DOB column. A Year of Joining column is not accepted as DOB. No certificate was released.");
  }

  const range = sh.getDataRange();
  const data = range.getValues();
  const display = range.getDisplayValues();
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const rowCode = normalizeCpdEmployeeCode_(display[i][c.employee - 1] || data[i][c.employee - 1]);
    const rowDob = normalizeCpdDob_(data[i][dobCol - 1], display[i][dobCol - 1]);
    if (rowCode === employeeCode && rowDob === dobInput) {
      const email = String(display[i][c.email - 1] || "").trim().toLowerCase();
      const name = String(display[i][c.name - 1] || "").trim();
      if (email && name) matches.push({row: i + 1, employeeCode: rowCode, email: email, name: name});
    }
  }
  if (matches.length !== 1) {
    throw new Error(matches.length > 1
      ? "More than one participant record matches these details. Please contact the workshop coordinator."
      : "Employee Code and Date of Birth did not match a participant record. Please check the 13 digits and confirm that your DOB is recorded correctly in the participant master sheet.");
  }
  return matches[0];
}

function maskCpdEmail_(email) {
  const parts = String(email || "").split("@");
  if (parts.length !== 2) return "your registered email";
  const local = parts[0];
  const shown = local.length <= 2 ? local.charAt(0) + "*" : local.slice(0,2) + "****";
  return shown + "@" + parts[1];
}

function sendCpdCertificateOtp(employeeCodeDob) {
  const participant = cpdCertificateIdentity_(employeeCodeDob);
  const cache = CacheService.getScriptCache();
  const key = "cpd-cert-otp-" + Utilities.base64EncodeWebSafe(participant.employeeCode).slice(0,40);
  const throttleKey = key + "-sent";
  if (cache.get(throttleKey)) throw new Error("An OTP was sent recently. Please wait one minute before requesting another.");
  const otp = String(Math.floor(100000 + Math.random()*900000));
  const salt = Utilities.getUuid();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + otp)
    .map(b => (b+256).toString(16).slice(-2)).join("");
  cache.put(key, JSON.stringify({hash:hash,salt:salt,row:participant.row,employeeCode:participant.employeeCode,email:participant.email,attempts:0}), CPD_CERTIFICATE_OTP_TTL);
  cache.put(throttleKey, "1", 60);
  MailApp.sendEmail({
    to: participant.email,
    subject: "ZIET Mumbai — CPD Certificate Verification OTP",
    body: "Dear " + participant.name + ",\n\nYour one-time verification code for downloading your CPD certificate is: " + otp +
      "\n\nThis code expires in 5 minutes and can be used only once. Do not share it with anyone. If you did not request this code, please ignore this email.\n\nZIET Mumbai"
  });
  return {ok:true, maskedEmail:maskCpdEmail_(participant.email), message:"A 6-digit OTP has been sent to your registered email address (" + maskCpdEmail_(participant.email) + "). It expires in 5 minutes."};
}

function verifyCpdCertificateOtp(employeeCodeDob, otp) {
  const participant = cpdCertificateIdentity_(employeeCodeDob);
  otp = String(otp || "").trim();
  if (!/^\d{6}$/.test(otp)) throw new Error("Enter the 6-digit OTP sent to your registered email.");
  const cache = CacheService.getScriptCache();
  const key = "cpd-cert-otp-" + Utilities.base64EncodeWebSafe(participant.employeeCode).slice(0,40);
  const raw = cache.get(key);
  if (!raw) throw new Error("OTP expired or not requested. Please request a new OTP.");
  const state = JSON.parse(raw);
  if (state.row !== participant.row || state.email !== participant.email) {
    cache.remove(key);
    throw new Error("Participant verification state changed. Please request a new OTP.");
  }
  state.attempts++;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, state.salt + "|" + otp)
    .map(b => (b+256).toString(16).slice(-2)).join("");
  if (hash !== state.hash) {
    if (state.attempts >= CPD_CERTIFICATE_MAX_ATTEMPTS) cache.remove(key);
    else cache.put(key, JSON.stringify(state), CPD_CERTIFICATE_OTP_TTL);
    throw new Error(state.attempts >= CPD_CERTIFICATE_MAX_ATTEMPTS
      ? "Too many incorrect attempts. Request a new OTP."
      : "Incorrect OTP. Please try again.");
  }
  cache.remove(key);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put("cpd-cert-session-" + token, JSON.stringify({
    row:participant.row, employeeCode:participant.employeeCode, email:participant.email, name:participant.name
  }), CPD_CERTIFICATE_SESSION_TTL);
  return {ok:true, token:token, message:"Email verified successfully. You can now download your CPD certificate."};
}

function downloadCpdCertificate(token) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get("cpd-cert-session-" + String(token || ""));
  if (!raw) throw new Error("Your verification session has expired. Please verify again.");
  const session = JSON.parse(raw);
  const {sh} = mainSheet_();
  const c = columns_(sh);
  const currentName = c.name ? String(sh.getRange(session.row,c.name).getDisplayValue() || "").trim() : "";
  const currentCode = c.employee ? String(sh.getRange(session.row,c.employee).getDisplayValue() || "").trim() : "";
  const currentEmail = c.email ? String(sh.getRange(session.row,c.email).getDisplayValue() || "").trim().toLowerCase() : "";
  if (currentName !== session.name || currentCode !== session.employeeCode || currentEmail !== session.email) {
    cache.remove("cpd-cert-session-" + String(token || ""));
    throw new Error("Participant record changed after verification. Please verify again.");
  }
  const folder = DriveApp.getFolderById(CPD_CERTIFICATE_FOLDER_ID);
  const files = folder.getFiles();
  const allowed = {"application/pdf":true,"image/png":true,"image/jpeg":true};
  const matches = [];
  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType();
    if (!allowed[mime]) continue;
    const fileName = file.getName().replace(/\.(pdf|png|jpe?g)$/i,"").trim();
    if (fileName.toLowerCase() === session.name.toLowerCase()) matches.push(file);
  }
  if (matches.length === 0) throw new Error("Your certificate was not found. Please contact the workshop coordinator.");
  if (matches.length > 1) throw new Error("More than one certificate file matches your name. Please contact the workshop coordinator; no file was released.");
  const file = matches[0];
  if (file.getSize() > 10 * 1024 * 1024) throw new Error("The certificate file is larger than 10 MB. Please contact the coordinator.");
  const blob = file.getBlob();
  return {
    fileName:file.getName(),
    mimeType:blob.getContentType(),
    base64:Utilities.base64Encode(blob.getBytes()),
    message:"Certificate retrieved successfully."
  };
}
