/**
 * RAJ Agencies — Customer Experience & Testimonial System
 * Apps Script Web App + Google Sheets backend
 */
const CFG = {
  COMPANY: 'Raj Agencies',
  BRAND: 'RAJ GROUP',
  EMAIL: 'akkirajagencies@gmail.com',
  SHEET_NAME: 'Testimonials',
  DASHBOARD_NAME: 'Dashboard',
  SETTINGS_NAME: 'Settings',
  TZ: 'Asia/Kolkata',
  ID_PREFIX: 'RAJ-T-',
  // Upload RAJ_GROUP_Logo.png to Google Drive, make it viewable by link, and paste the file ID below.
  LOGO_DRIVE_FILE_ID: '1cO82ZCHe8P3a_JWtGfYnAZbUg7uPH8Dg',
  BACKGROUND_DRIVE_FILE_ID: '1t-yAl_9x6KH-8E4NGP4nS3KZa_ZUXds5',
  SEGMENTS: ['2 Wheeler','3 Wheeler','Car','LCV','HCV','Tractor','Earthmovers','Agriculture Parts','Hardware']
};

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.config = getPublicConfig_();
  return t.evaluate()
    .setTitle('RAJ Agencies | Customer Experience')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



/** GitHub Pages API endpoint for https://feedback.rajagencies.info */
function doPost(e) {
  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      return jsonResponse_({ok:false, error:'Missing payload'});
    }
    const data = JSON.parse(e.parameter.payload);
    const result = submitTestimonial(data);
    return jsonResponse_(result);
  } catch (error) {
    console.error('doPost error: ' + error);
    return jsonResponse_({ok:false, error:error && error.message ? error.message : 'Submission failed'});
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPublicConfig_() {
  let logoUrl = '';
  if (CFG.LOGO_DRIVE_FILE_ID) {
    try {
      const blob = DriveApp.getFileById(CFG.LOGO_DRIVE_FILE_ID).getBlob();
      logoUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
    } catch (e) {
      console.log('Logo load failed: ' + e.message);
    }
  }
  return {
    company: CFG.COMPANY,
    brand: CFG.BRAND,
    segments: CFG.SEGMENTS,
    logoUrl: logoUrl,
    backgroundUrl: 'https://drive.google.com/thumbnail?id=' + CFG.BACKGROUND_DRIVE_FILE_ID + '&sz=w2000'
  };
}

function setupRajSystem() {
  const ss = SpreadsheetApp.create('RAJ Agencies - Testimonial Management System');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone(CFG.TZ);
  const sh = ss.getSheets()[0];
  sh.setName(CFG.SHEET_NAME);
  const headers = ['Timestamp','Testimonial ID','Customer Name','Business / Shop Name','City','Mobile','Customer Type','Segments','Years Associated','Product Quality','Product Range','Parts Availability','Pricing','Delivery','Staff Support','Order Handling','Overall Experience','Recommend','Testimonial','Consent','Average Rating','Status','Featured','Admin Notes'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#075FCB').setFontColor('#FFFFFF');
  sh.getRange('A:X').setVerticalAlignment('middle');
  sh.autoResizeColumns(1, headers.length);
  sh.setColumnWidth(19, 420);
  sh.setColumnWidth(24, 320);
  sh.getRange('V2:V').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'], true).build());
  sh.getRange('W2:W').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['No','Yes'], true).build());
  sh.getRange(1,1,sh.getMaxRows(),headers.length).createFilter();

  const dash = ss.insertSheet(CFG.DASHBOARD_NAME);
  buildDashboard_(dash);
  const settings = ss.insertSheet(CFG.SETTINGS_NAME);
  settings.getRange('A1:B8').setValues([
    ['RAJ Agencies Testimonial System','Configuration'],
    ['Notification Email',CFG.EMAIL],
    ['ID Prefix',CFG.ID_PREFIX],
    ['Timezone',CFG.TZ],
    ['Brand Primary','#075FCB'],
    ['Brand Gold','#FFB000'],
    ['Status Flow','Pending → Approved / Rejected'],
    ['Spreadsheet ID',ss.getId()]
  ]);
  settings.getRange('A1:B1').setFontWeight('bold').setBackground('#FFB000');
  settings.autoResizeColumns(1,2);
  SpreadsheetApp.flush();
  return {spreadsheetUrl:ss.getUrl(), spreadsheetId:ss.getId(), message:'RAJ Testimonial System created successfully.'};
}

function buildDashboard_(sh) {
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1:H2').merge().setValue('RAJ AGENCIES — TESTIMONIAL DASHBOARD')
    .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#075FCB')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeights(1,2,34);
  const cards = [
    ['A4:B4','Total Testimonials',"=COUNTA(Testimonials!B2:B)"],
    ['C4:D4','Average Rating',"=IFERROR(ROUND(AVERAGE(Testimonials!U2:U),2),0)"],
    ['E4:F4','Approved',"=COUNTIF(Testimonials!V2:V,\"Approved\")"],
    ['G4:H4','Featured',"=COUNTIF(Testimonials!W2:W,\"Yes\")"]
  ];
  cards.forEach(([r,label,formula])=>{
    const range=sh.getRange(r); range.merge().setValue(label).setFontWeight('bold').setBackground('#FFF3D1').setHorizontalAlignment('center');
    const row=range.getRow()+1, col=range.getColumn();
    sh.getRange(row,col,1,2).merge().setFormula(formula).setFontSize(22).setFontWeight('bold').setFontColor('#075FCB').setHorizontalAlignment('center');
  });
  sh.getRange('A8:H8').merge().setValue('Management Notes').setFontWeight('bold').setBackground('#FFB000').setHorizontalAlignment('center');
  sh.getRange('A9:H12').merge().setValue('Use the Testimonials sheet to review submissions. Change Status to Approved/Rejected and Featured to Yes/No. The dashboard updates automatically.').setWrap(true).setVerticalAlignment('top');
  for(let c=1;c<=8;c++) sh.setColumnWidth(c,125);
}

function submitTestimonial(data) {
  validate_(data);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Another submission is being processed. Please try again.');
  try {
    const ss = getSS_();
    const sh = ss.getSheetByName(CFG.SHEET_NAME);
    const id = nextId_(sh);
    const ratings = ['productQuality','productRange','availability','pricing','delivery','support','orderHandling','overall'].map(k => Number(data[k] || 0));
    const avg = Math.round((ratings.reduce((a,b)=>a+b,0)/ratings.length)*100)/100;
    const rowData = [
      new Date(), id, clean_(data.name), clean_(data.business), clean_(data.city), clean_(data.mobile), clean_(data.customerType),
      (data.segments||[]).join(', '), clean_(data.years), ratings[0],ratings[1],ratings[2],ratings[3],ratings[4],ratings[5],ratings[6],ratings[7],
      clean_(data.recommend), clean_(data.testimonial), clean_(data.consent), avg, 'Pending', 'No', clean_(data.clientRef || '')
    ];
    const row = sh.getLastRow() + 1;
    sh.getRange(row,1,1,rowData.length).setValues([rowData]);
    sh.getRange(row,1,1,24).setVerticalAlignment('top');
    if (avg < 3.5) sh.getRange(row,1,1,24).setBackground('#FFF0F0');
    // Keep the customer-facing submit path fast. Email is intentionally not sent here.
    // The saved Sheet row is the source of truth; notification can be added separately if required.
    return {ok:true, id:id, average:avg};
  } finally { lock.releaseLock(); }
}

function validate_(d) {
  if (!d || !clean_(d.name) || !clean_(d.business) || !clean_(d.city)) throw new Error('Please complete all required customer details.');
  if (!Array.isArray(d.segments) || !d.segments.length) throw new Error('Please select at least one business segment.');
  if (!clean_(d.testimonial) || clean_(d.testimonial).length < 20) throw new Error('Please write a testimonial of at least 20 characters.');
  const keys=['productQuality','productRange','availability','pricing','delivery','support','orderHandling','overall'];
  if (keys.some(k => Number(d[k]) < 1 || Number(d[k]) > 5)) throw new Error('Please provide all ratings.');
}

function nextId_(sh) {
  const n = Math.max(1, sh.getLastRow());
  return CFG.ID_PREFIX + String(n).padStart(6,'0');
}
function clean_(v){ return String(v == null ? '' : v).trim().replace(/[<>]/g,''); }
function getSS_(){
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id) throw new Error('System is not initialized. Run setupRajSystem() once from Apps Script.');
  return SpreadsheetApp.openById(id);
}

function sendAdminEmail_(id,d,avg,url){
  const subject=`New RAJ Testimonial ${id} — ${d.business}`;
  const body=`A new customer testimonial has been received.\n\nID: ${id}\nCustomer: ${d.name}\nBusiness: ${d.business}\nCity: ${d.city}\nSegments: ${(d.segments||[]).join(', ')}\nAverage Rating: ${avg}/5\nRecommend: ${d.recommend}\n\nTestimonial:\n${d.testimonial}\n\nReview in dashboard: ${url}`;
  MailApp.sendEmail(CFG.EMAIL,subject,body);
}

function onOpen(){
  try { SpreadsheetApp.getUi().createMenu('RAJ Testimonial System')
    .addItem('Open Dashboard','openDashboard_')
    .addItem('Refresh Dashboard','refreshDashboard_')
    .addSeparator()
    .addItem('Approve Selected','approveSelected_')
    .addItem('Reject Selected','rejectSelected_')
    .addItem('Mark Featured','featureSelected_')
    .addToUi(); } catch(e) { console.log('onOpen skipped: ' + e); }
}
function openDashboard_(){ SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(CFG.DASHBOARD_NAME)); }
function refreshDashboard_(){ SpreadsheetApp.flush(); SpreadsheetApp.getActive().toast('Dashboard refreshed','RAJ GROUP',3); }
function setSelectedStatus_(status){ const sh=SpreadsheetApp.getActiveSheet(); if(sh.getName()!==CFG.SHEET_NAME) return; const r=sh.getActiveRange().getRow(); if(r>1) sh.getRange(r,22).setValue(status); }
function approveSelected_(){setSelectedStatus_('Approved');}
function rejectSelected_(){setSelectedStatus_('Rejected');}
function featureSelected_(){ const sh=SpreadsheetApp.getActiveSheet(); if(sh.getName()!==CFG.SHEET_NAME) return; const r=sh.getActiveRange().getRow(); if(r>1){sh.getRange(r,22).setValue('Approved');sh.getRange(r,23).setValue('Yes');} }
