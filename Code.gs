/**
 * The Origin PH — order backend (Google Apps Script)
 * Setup: open your Google Sheet > Extensions > Apps Script > paste this file >
 * Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * Copy the /exec URL into ORDER_ENDPOINT in index.html.
 */
const SHEET_ID     = '';   // paste your Google Sheet ID here if the script is NOT created from inside the sheet
const SHEET_NAME   = 'Orders';
const FOLDER_NAME  = 'Origin PH Payment Proofs';
const OWNER_EMAIL  = 'emmanuellevillaranda@gmail.com';   // new-order notifications
const PRICES = {t15:1700,t30:2500,t60:4000,ghk50:1300,ghk100:1600,kpv:1600,nad:1600,mots:3000,fuan:1800,
  rt20:2400,cagri10:3000,lcarn600:1800,fat526:2300,ss31:2000,blend50:2300,blend100:2700};
const RATES = {luzon:[95,115,180],visayas:[155,175,250],mindanao:[185,205,350]};
const _UNUSED_HEADERS = ['Timestamp','Reference','Status','Name','Age','Phone','Delivery','Island group','Province','City','Barangay',
  'Street','ZIP','Landmark','Items','Total kits','Product subtotal','Extra syringes','Shipping','Payment method','Payment fee',
  'Total (customer)','Total (server check)','Total OK?','Payment ref #','Proof link','Gifts'];

function doGet(){ return json_({ok:true,service:'The Origin PH orders'}); }

function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const data = JSON.parse(e.postData.contents);
    const o = data.order;
    if (!o || !o.items || !o.items.length) throw new Error('Empty order');
    if (!data.proofBase64) throw new Error('Missing payment proof');

    // Recompute the total on the server so edited prices cannot slip through.
    let qty = 0, sub = 0;
    o.items.forEach(i => { if (!(i.id in PRICES)) throw new Error('Unknown product'); qty += i.quantity; sub += PRICES[i.id]*i.quantity; });
    if (qty < 1 || qty > 6) throw new Error('Invalid quantity');
    const extra = Math.max(0, Math.floor(o.extraSyringes||0));
    let ship = 0;
    if (o.deliveryMethod === 'jnt') {
      const r = RATES[o.address.islandGroup]; if (!r) throw new Error('Invalid region');
      ship = r[qty===1?0:qty<=3?1:2];
    }
    const before = sub + extra*15 + ship;
    const fee = o.paymentMethod === 'QR Ph' ? Math.round(before*0.015*100)/100 : 0;
    const serverTotal = Math.round((before+fee)*100)/100;
    const totalOk = Math.abs(serverTotal - Number(o.total)) < 0.02;

    // Save proof image to Drive.
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

    // Sequential reference per day: REF + MMDDYYYY + NN
    const props = PropertiesService.getScriptProperties();
    const now = new Date();
    const day = Utilities.formatDate(now,'Asia/Manila','MMddyyyy');
    const n = (Number(props.getProperty('seq_'+day))||0)+1;
    props.setProperty('seq_'+day,String(n));
    const ref = /^REF\d{8}[A-Z0-9]{4}$/.test(o.clientRef||'') ? o.clientRef : 'REF'+day+('0'+n).slice(-2);

    const blob = Utilities.newBlob(Utilities.base64Decode(data.proofBase64), data.proofMime||'image/jpeg', ref+'_'+(o.lastName||'proof')+'.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No spreadsheet: set SHEET_ID at the top of the script');
    const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const a = o.address||{};
    const V = {
      timestamp: now, ref: ref, status: 'Pending verification', name: o.name, firstName: o.firstName, lastName: o.lastName,
      age: o.age, phone: "'"+o.phone, delivery: o.deliveryMethod==='jnt'?'J&T Express EZ':'Lalamove',
      islandGroup: a.islandGroup, province: a.province, city: a.city, barangay: a.barangay, street: a.street,
      zip: a.zip, landmark: a.landmark, items: o.items.map(i=>i.quantity+' x '+i.name).join('; '),
      freeSyr: qty*6, extra: extra, payMethod: o.paymentMethod, subtotal: sub, extraCost: extra*15, shipping: ship,
      fee: fee, total: serverTotal, check: totalOk?'OK':'MISMATCH (customer sent '+o.total+')',
      payRef: o.paymentReference||'', proof: file.getUrl(), gifts: (o.complimentaryGifts||[]).join(', ')
    };
    if (sh.getLastRow() === 0) sh.appendRow(['Timestamp / Date','First Name','Last Name']);
    let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    // Add any columns your sheet is missing (reference, status, proof link...).
    const have = headers.map(colKey_);
    [['ref','Reference'],['status','Status'],['payRef','Payment Ref #'],['proof','Proof Link'],['check','Total Check']].forEach(p=>{
      if (have.indexOf(p[0]) < 0) { headers.push(p[1]); sh.getRange(1, headers.length).setValue(p[1]); }
    });
    sh.appendRow(headers.map(h => { const k = colKey_(h); return k && V[k] !== undefined ? V[k] : ''; }));

    if (OWNER_EMAIL) MailApp.sendEmail(OWNER_EMAIL, 'New order '+ref+' — ₱'+serverTotal,
      o.name+' | '+o.items.map(i=>i.quantity+' x '+i.name).join(', ')+'\nPayment proof: '+file.getUrl()+(totalOk?'':'\n⚠ TOTAL MISMATCH'));

    return json_({ok:true, referenceNumber:ref});
  } catch (err) {
    return json_({ok:false, error:String(err.message||err)});
  } finally { try{lock.releaseLock();}catch(_){} }
}

// Matches your sheet's header names to order fields (case-insensitive, order matters).
const FIELDS = [
  ['payRef',/payment ref|transaction/i],['ref',/reference|order (no|id|num)|^ref/i],['status',/status/i],['proof',/proof|screenshot/i],
  ['check',/total (ok|check)|^check/i],['gifts',/gift/i],['timestamp',/time|date/i],['firstName',/first/i],['lastName',/last/i],
  ['age',/^age/i],['phone',/phone|mobile|contact/i],['delivery',/deliver|courier/i],['islandGroup',/group|island/i],
  ['province',/province|region/i],['city',/city|municip/i],['barangay',/barangay/i],['street',/street|address/i],
  ['zip',/postal|zip/i],['landmark',/landmark/i],['freeSyr',/complimentary|free/i],['extraCost',/extra.*(cost|total|amount)/i],
  ['extra',/extra/i],['items',/item|product/i],['payMethod',/payment method/i],['subtotal',/subtotal/i],['shipping',/shipping/i],
  ['fee',/processing|fee/i],['total',/grand|total/i],['name',/^(full )?name/i]
];
function colKey_(h){ for (const f of FIELDS) if (f[1].test(h)) return f[0]; return null; }

function json_(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/** Run this from the editor (select testOrder > Run) to test saving + email without the website. */
function testOrder(){
  const fake = {postData:{contents:JSON.stringify({
    proofBase64: Utilities.base64Encode(Utilities.newBlob('test').getBytes()),
    proofMime:'image/jpeg',
    order:{name:'TEST ORDER',firstName:'Test',lastName:'Order',age:30,phone:'09000000000',deliveryMethod:'jnt',
      address:{islandGroup:'luzon',province:'Metro Manila (NCR)',city:'Manila City',barangay:'Test',street:'Test St',zip:'1000',landmark:''},
      items:[{id:'t15',name:'Tirzepatide 15mg',quantity:1}],extraSyringes:0,paymentMethod:'InstaPay',total:1795,paymentReference:'TEST',complimentaryGifts:[]}})}};
  Logger.log(doPost(fake).getContent());
}