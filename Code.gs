/**
 * RAJ Agencies Feedback API — media-enabled backend
 * Script Property required: SPREADSHEET_ID
 * Optional Script Property: TESTIMONIAL_MEDIA_FOLDER_ID
 * If folder ID is absent, a private Drive folder named "RAJ Agencies - Customer Testimonial Media" is created.
 */
const RAJ_SHEET_NAME = 'Testimonials';
const RAJ_MEDIA_FOLDER_NAME = 'RAJ Agencies - Customer Testimonial Media';

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ok:true, service:'RAJ Agencies Feedback API'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) throw new Error('System busy. Please retry.');
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('Missing payload');
    const d = JSON.parse(raw);
    validatePayload_(d);

    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!ssId) throw new Error('SPREADSHEET_ID Script Property is missing');
    const ss = SpreadsheetApp.openById(ssId);
    const sh = ss.getSheetByName(RAJ_SHEET_NAME);
    if (!sh) throw new Error('Testimonials sheet not found');

    const id = nextRajId_(sh);
    let media = {type:'', url:'', fileId:'', name:'', consent:''};
    if (d.media && d.media.data) media = saveMedia_(d.media, id, d.name || 'Customer');

    ensureMediaColumns_(sh);
    appendMappedRow_(sh, d, id, media);
    return json_({ok:true,id:id});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function validatePayload_(d) {
  if (!d || !String(d.name||'').trim() || !String(d.business||'').trim() || !String(d.city||'').trim()) throw new Error('Required customer details missing');
  if (!String(d.testimonial||'').trim() || String(d.testimonial).trim().length < 20) throw new Error('Testimonial is too short');
  if (d.media && d.media.data) {
    if (!d.media.consent) throw new Error('Media consent required');
    const type = String(d.media.type||'');
    const size = Number(d.media.size||0);
    if (!(type.indexOf('image/')===0 || type.indexOf('video/')===0)) throw new Error('Unsupported media type');
    if (type.indexOf('image/')===0 && size > 5*1024*1024) throw new Error('Photo exceeds 5 MB');
    if (type.indexOf('video/')===0 && size > 15*1024*1024) throw new Error('Video exceeds 15 MB');
  }
}

function getMediaFolder_() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('TESTIMONIAL_MEDIA_FOLDER_ID');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (_) {} }
  const it = DriveApp.getFoldersByName(RAJ_MEDIA_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(RAJ_MEDIA_FOLDER_NAME);
  props.setProperty('TESTIMONIAL_MEDIA_FOLDER_ID', folder.getId());
  return folder;
}

function saveMedia_(m, id, customer) {
  const type = String(m.type||'application/octet-stream');
  const ext = extensionFor_(type, m.name);
  const safeCustomer = String(customer||'Customer').replace(/[^a-zA-Z0-9 _-]/g,'').trim().replace(/\s+/g,'_').slice(0,35) || 'Customer';
  const filename = id + '_' + safeCustomer + '_' + (type.indexOf('image/')===0 ? 'Photo' : 'Video') + ext;
  const bytes = Utilities.base64Decode(String(m.data));
  const blob = Utilities.newBlob(bytes, type, filename);
  const file = getMediaFolder_().createFile(blob);
  // Intentionally private. Do not call setSharing(ANYONE...).
  return {type:type.indexOf('image/')===0?'Photo':'Video', url:file.getUrl(), fileId:file.getId(), name:filename, consent:'Yes'};
}

function extensionFor_(mime, original) {
  const map={'image/jpeg':'.jpg','image/png':'.png','image/webp':'.webp','image/heic':'.heic','video/mp4':'.mp4','video/webm':'.webm','video/quicktime':'.mov'};
  if (map[mime]) return map[mime];
  const m=String(original||'').match(/\.[a-zA-Z0-9]{1,5}$/); return m?m[0]:'';
}

function ensureMediaColumns_(sh) {
  const last = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1,1,1,last).getValues()[0].map(String);
  const wanted=['Media Type','Media Drive Link','Media File ID','Media File Name','Media Consent'];
  const missing=wanted.filter(x=>!headers.some(h=>norm_(h)===norm_(x)));
  if (missing.length) sh.getRange(1,last+1,1,missing.length).setValues([missing]);
}

function appendMappedRow_(sh,d,id,media) {
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const avg=average_([d.productQuality,d.productRange,d.availability,d.pricing,d.delivery,d.support,d.orderHandling,d.overall]);
  const values=headers.map(h=>valueForHeader_(h,d,id,media,avg));
  sh.getRange(sh.getLastRow()+1,1,1,values.length).setValues([values]);
}

function valueForHeader_(h,d,id,m,avg) {
  const k=norm_(h);
  const map={
    'timestamp':new Date(),'date':new Date(),'testimonialid':id,'id':id,'referenceid':id,
    'customername':d.name,'name':d.name,'businessshopname':d.business,'businessname':d.business,'businessshop':d.business,'business':d.business,
    'city':d.city,'mobile':d.mobile,'mobilenumber':d.mobile,'customertype':d.customerType,'associationduration':d.years,'association':d.years,'years':d.years,
    'segments':Array.isArray(d.segments)?d.segments.join(', '):d.segments,
    'productquality':d.productQuality,'productrange':d.productRange,'partsavailability':d.availability,'availability':d.availability,'pricing':d.pricing,'delivery':d.delivery,'staffsupport':d.support,'support':d.support,'orderhandling':d.orderHandling,'overallexperience':d.overall,'overall':d.overall,
    'averagerating':avg,'avgrating':avg,'recommendation':d.recommend,'recommend':d.recommend,'testimonial':d.testimonial,'feedback':d.testimonial,'consent':d.consent,
    'status':'Pending','featured':'No','adminnotes':d.clientRef?('Web ref: '+d.clientRef):'',
    'mediatype':m.type,'mediadrivelink':m.url,'mediaurl':m.url,'mediafileid':m.fileId,'mediafilename':m.name,'mediaconsent':m.consent
  };
  return Object.prototype.hasOwnProperty.call(map,k) ? map[k] : '';
}

function nextRajId_(sh) {
  const headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(norm_);
  let col=headers.findIndex(x=>x==='testimonialid'||x==='id'||x==='referenceid')+1;
  if (!col) col=2;
  const last=sh.getLastRow(); let max=0;
  if(last>1){sh.getRange(2,col,last-1,1).getDisplayValues().flat().forEach(v=>{const m=String(v).match(/RAJ-T-(\d+)/i);if(m)max=Math.max(max,Number(m[1]))})}
  return 'RAJ-T-'+String(max+1).padStart(6,'0');
}
function average_(a){const n=a.map(Number).filter(x=>x>=1&&x<=5);return n.length?Math.round((n.reduce((x,y)=>x+y,0)/n.length)*100)/100:''}
function norm_(s){return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'')}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function healthCheck(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');const sh=id&&SpreadsheetApp.openById(id).getSheetByName(RAJ_SHEET_NAME);const folder=getMediaFolder_();return {spreadsheet:!!id,sheet:!!sh,mediaFolderId:folder.getId(),mediaFolderName:folder.getName()}}
