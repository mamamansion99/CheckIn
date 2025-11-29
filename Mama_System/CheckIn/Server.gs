// ---------------- CONFIG (single object) ----------------
var CONFIG = {
  SHEET_ID: '1KsimOBXcP2PhZ3Y16DXo7KKcTO9sMNksKJbxc5VEHEQ',
  CHECKIN_LOG_SHEET: 'Checkin_Log',
  ROOMS_SHEET: 'Rooms',
  DEFAULT_FOLDER: '1O6KDNnOWrFoUI54GBKbKhoL-ESjlfwfY',

  // headers on Rooms sheet
  ROOM_HEADER: 'RoomID',          // e.g. A101
  ROOM_HDR_ID: 'RoomFolderId',    // main room folder id
  CHECKIN_HDR_ID: 'CheckInFolderId',
  CHECKOUT_HDR_ID: 'CheckOutFolderId', // (not used in this code)
  ROOM_STATUS_HEADER: 'Status',   // column storing current occupancy state
  TEMPLATE_ID: '1OWoJ0GTSnh43QXslYZ42pmQ0nKV7k4_fZsdTVSvb9OE',

  OPENCHAT_LINKS: {
    A: 'https://line.me/ti/g2/dONR8eAdCqgxzVm_5R_rT0OHcVthoguInw74LQ?utm_source=invitation&utm_medium=link_copy&utm_campaign=default',
    B: 'https://line.me/ti/g2/7vuG4-RhTbE6zknNYX_JR6buwMWF89Cbxmw7Og?utm_source=invitation&utm_medium=link_copy&utm_campaign=default'
  }

};

const WELCOME_URL  = "https://drive.google.com/file/d/1wWcvDnXa5oJnAFxfdYAhc1f35Nr8Kfhx/view?usp=drive_link";
const OPENCHAT_LINKS = CONFIG.OPENCHAT_LINKS || {};
const N8N_WEBHOOK_URL = 'https://n8n.srv1112305.hstgr.cloud/webhook-test/37fbebfb-2a9a-410c-b913-728550b8fbd9';

/* ---------------- Web App ---------------- */
function doGet() {
  var t = HtmlService.createTemplateFromFile('Index'); // require Index.html in project
  t.webAppUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Room Inspector')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- helpers: area normalization & headers ---------- */
/** Normalize area names so headers are consistent.
 *  - trims, uppercases
 *  - PHOTO[BED]  -> BED
 *  - BED[]       -> BED
 *  - "curtain "  -> CURTAIN
 */
function normalizeArea(area) {
  let s = String(area || '').trim().toUpperCase();
  // If pattern like XXX[YYY] keep inside; otherwise use as-is
  const m = s.match(/^[A-Z_]+(?:\[(.+?)\])?$/);
  if (m && m[1]) s = m[1].trim().toUpperCase();
  // drop trailing [] and collapse spaces to underscores
  s = s.replace(/\[\]$/, '').replace(/\s+/g, '_');
  return s || 'GEN';
}

function parseHeader(header) {
  const raw = String(header || '').trim();
  // split by last underscore -> LEFT = area-ish, RIGHT = suffix-ish
  const us = raw.lastIndexOf('_');
  if (us === -1) return { area: normalizeArea(raw), suffix: '' };

  const left = raw.slice(0, us);
  const right = raw.slice(us + 1);

  const area = normalizeArea(left);
  // normalize suffix: NOTE/NOTES, PHOTO/PHOTOS, STATUS
  let suffix = String(right || '').trim().toUpperCase();
  if (suffix === 'NOTE') suffix = 'NOTES';
  if (suffix === 'PHOTO') suffix = 'PHOTOS';
  return { area, suffix };
}

/** 3) Find an existing column by (area, suffix) in a tolerant way. */
function findExistingCol(sh, wantArea, wantSuffix /* 'STATUS'|'NOTES'|'PHOTOS' */) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return -1;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || ''));
  const A = normalizeArea(wantArea);
  const S = String(wantSuffix || '').trim().toUpperCase().replace(/^NOTE$/, 'NOTES').replace(/^PHOTO$/, 'PHOTOS');

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    const { area, suffix } = parseHeader(h);
    if (area === A && suffix === S) {
      return c + 1; // 1-based
    }
  }
  return -1;
}

/** 4) Get or create the correct column. Only create if truly missing. */
function getOrCreateCol(sh, area, suffix /* 'STATUS'|'NOTES'|'PHOTOS' */) {
  const idx = findExistingCol(sh, area, suffix);
  if (idx !== -1) return idx;

  // create with canonical header: AREA_Status/Notes/Photos
  const A = normalizeArea(area);
  let S = String(suffix || '').toUpperCase();
  if (S === 'NOTE') S = 'NOTES';
  if (S === 'PHOTO') S = 'PHOTOS';

  const header = `${A}_${S.charAt(0)}${S.slice(1).toLowerCase()}`; // e.g., BED_Photos
  const lastCol = sh.getLastColumn();
  const headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  headers.push(header);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers.length; // new 1-based index
}

/** Ensure the 3 columns for a given area exist */
/** 5) Convenience helpers mirroring your old names */
function ensureColumns(sh, rawArea) {
  const a = normalizeArea(rawArea);
  getOrCreateCol(sh, a, 'STATUS');
  getOrCreateCol(sh, a, 'NOTES');
  getOrCreateCol(sh, a, 'PHOTOS');
}

function getColIndex(sh, headerName) {
  // keep for backward compatibility if something else still calls it
  // but prefer getColByAreaSuffix below
  const lastCol = sh.getLastColumn();
  if (!lastCol) throw new Error('No headers');
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const idx = headers.indexOf(headerName);
  if (idx === -1) throw new Error('Header not found: ' + headerName);
  return idx + 1;
}
function getColByAreaSuffix(sh, area, suffix) {
  return getOrCreateCol(sh, area, suffix);
}

function notifyN8N(roomId, building, floor) {
  const payload = {
    source: 'room_inspect',
    roomId,
    building,
    floor,
    timestamp: new Date().toISOString()
  };

  try {
    UrlFetchApp.fetch(N8N_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('⚠️ notifyN8N failed: ' + err);
  }
}

function doPost(e) {
  try {
    var ct = (e && e.postData && e.postData.type || '').toLowerCase();
    if (!ct.includes('application/json')) {
      return ContentService.createTextOutput(JSON.stringify({ ok:false, error:'Use JSON' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var payload    = JSON.parse(e.postData.contents || '{}');
    var fields     = payload.fields     || {};
    var metaByArea = payload.metaByArea || {}; // { AREA:{status, note}, ... }
    var filesArr   = payload.files      || [];

    var building    = fields.building  || '';
    var floor       = fields.floor     || '';
    var roomId      = String(fields.roomId || '').trim();
    var inspector   = fields.inspector || '';
    var globalNotes = fields.globalNotes || '';

    // === resolve Drive folder ===
    var resolved = { via:'checkin', id:getFolderIdForRoom(roomId) };
    if (!resolved.id) resolved = { via:'room',    id:getRoomFolderId(roomId) };
    if (!resolved.id) resolved = { via:'default', id:CONFIG.DEFAULT_FOLDER };
    var folder = DriveApp.getFolderById(resolved.id);

    // === timestamp ===
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var ts  = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');

    // === save uploaded photos ===
    var perArea = {}; // { AREA: [urls...] }
    filesArr.forEach(function(f, i){
      var area = normalizeArea(f.area || 'GEN');
      var mime = f.mime || 'application/octet-stream';
      var name = f.name || ('upload_'+i+'.bin');
      var blob = Utilities.newBlob(Utilities.base64Decode(f.base64 || ''), mime, name);

      var ext  = guessExt(mime, name);
      blob.setName((roomId || 'ROOM') + '_' + ts + '_' + area + (filesArr.length > 1 ? '_' + (i+1) : '') + ext);

      var driveFile = folder.createFile(blob);
      var fileUrl   = driveFile.getUrl();

      if (!perArea[area]) perArea[area] = [];
      perArea[area].push(fileUrl);
    });

    // === save signature (base64 → Drive) ===
    var signatureUrl = '';
    var sigData = '';

    if (fields.tenantSignature) {
      sigData = fields.tenantSignature;
    } else if (e.parameter.tenantSignature) {
      sigData = e.parameter.tenantSignature;
    }

    if (sigData && sigData.startsWith('data:image/png;base64,')) {
      var base64 = sigData.split(',')[1];
      var bytes = Utilities.base64Decode(base64);
      var sigBlob = Utilities.newBlob(bytes, 'image/png',
        (roomId || 'ROOM') + '_' + ts + '_SIGNATURE.png');
      var sigFile = folder.createFile(sigBlob);
      signatureUrl = sigFile.getUrl();
    }


    // Fallback: if no base64, use first uploaded file under area "SIGNATURE"
    if (!signatureUrl && perArea['SIGNATURE'] && perArea['SIGNATURE'].length) {
      signatureUrl = perArea['SIGNATURE'][0];
    }

    // === append row to Checkin_Log ===
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sh = ss.getSheetByName(CONFIG.CHECKIN_LOG_SHEET);
    var base = [now, building, floor, roomId, inspector, globalNotes, folder.getUrl()];
    while (base.length < sh.getLastColumn()) base.push('');
    sh.appendRow(base);
    var r = sh.getLastRow();

    updateRoomStatus(ss, roomId, 'Checked In');

    // === prepare rowObj for PDF ===
    var areas = new Set([
      ...Object.keys(metaByArea || {}),
      ...Object.keys(perArea   || {})
    ]);
    const rowObj = {
      timestamp: now,
      building,
      floor,
      roomId,
      inspector,
      globalNotes,
      folderUrl: folder.getUrl(),
      signatureUrl
    };
    areas.forEach(a => {
      const meta = metaByArea[a] || {};
      const st = meta.status || 'ok';    // always at least "ok"
      const nt = meta.note   || '';
      const ph = (perArea[a] || []);     // keep as array of URLs

      rowObj[`${a}_STATUS`] = st;
      rowObj[`${a}_NOTES`]  = nt;
      rowObj[`${a}_PHOTOS`] = ph;  // <-- store array, not joined text
    });

    // === generate PDF ===
    var pdfUrl = createInspectionPdf(rowObj);

    // make PDF shareable
    try {
      var pdfFile = DriveApp.getFileById(pdfUrl.match(/[-\w]{25,}/)[0]);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = pdfFile.getUrl();
    } catch (err) {
      Logger.log("⚠️ Cannot set sharing for PDF: " + err);
    }

    // === update sheet (PDF + Signature URL) ===
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var colPdf = headers.indexOf("Inspection PDF") + 1;
    if (colPdf > 0) sh.getRange(r, colPdf).setValue(pdfUrl);

    var colSig = headers.indexOf("Tenant Signature") + 1;
    if (signatureUrl && colSig > 0) sh.getRange(r, colSig).setValue(signatureUrl);

    // update photo/status/notes
    areas.forEach(function(aRaw){
      var a = normalizeArea(aRaw);
      ensureColumns(sh, a);

      var urls = (perArea[a] || []).join('\n');
      if (urls) sh.getRange(r, getColByAreaSuffix(sh, a, 'PHOTOS')).setValue(urls);

      var meta = metaByArea[aRaw] || metaByArea[a] || {};
      if (meta.status) sh.getRange(r, getColByAreaSuffix(sh, a, 'STATUS')).setValue(meta.status);
      if (meta.note)   sh.getRange(r, getColByAreaSuffix(sh, a, 'NOTES')).setValue(meta.note);
    });

    // fire-and-forget webhook for downstream automations
    notifyN8N(roomId, building, floor);


    // === find tenant's LINE ID & send PDFs + welcome pack ===
    var lineId = getLineIdForRoom(roomId);
    var bld = getBuildingLetter_(building, roomId);
    var inviteText = '';
    var availableBlds = [];
    if (OPENCHAT_LINKS) {
      Object.keys(OPENCHAT_LINKS).forEach(function (key) {
        var url = OPENCHAT_LINKS[key];
        if (!url) return;
        availableBlds.push({
          key: key,
          letter: String(key).toUpperCase(),
          url: url
        });
      });
    }

    var chosen = availableBlds.find(function (entry) { return entry.letter === bld; }) || null;
    if (chosen) {
      inviteText =
        `\n\n👥 ชุมชนผู้เช่า – อาคาร ${chosen.letter}\n` +
        `${chosen.url}\n` +
        `หมายเหตุ: กลุ่มนี้เพื่อคุยกันเอง/เตือนกัน • งานซ่อม/การเงินให้ติดต่อ OA เท่านั้นนะคะ`;
    } else if (availableBlds.length) {
      if (availableBlds.length === 1) {
        var single = availableBlds[0];
        inviteText =
          `\n\n👥 ชุมชนผู้เช่า – อาคาร ${single.letter}\n` +
          `${single.url}\n` +
          `หมายเหตุ: กลุ่มนี้เพื่อคุยกันเอง/เตือนกัน • งานซ่อม/การเงินให้ติดต่อ OA เท่านั้นนะคะ`;
      } else {
        var lines = availableBlds.map(function (entry) {
          return `อาคาร ${entry.letter}: ${entry.url}`;
        }).join('\n');
        inviteText =
          "\n\n👥 ชุมชนผู้เช่า\n" +
          lines + "\n" +
          "หมายเหตุ: กลุ่มนี้เพื่อคุยกันเอง/เตือนกัน • งานซ่อม/การเงินให้ติดต่อ OA เท่านั้นนะคะ";
      }
    }

    var tenantMsg = [
      `ยินดีต้อนรับสู่ Mama Mansion ห้อง ${roomId} 🎉`,
      `ขอบคุณที่ตรวจรับห้องและลงชื่อเรียบร้อยแล้วนะคะ`,
      '',
      '✅ ใบตรวจรับ (PDF)',
      pdfUrl,
      '',
      '📘 คู่มือการอยู่เบื้องต้น (PDF)',
      WELCOME_URL
    ];

    if (inviteText) tenantMsg.push('', inviteText.trim());

    tenantMsg.push(
      '',
      'ต้องการความช่วยเหลือ ทักแชทได้ตลอดค่ะ 💬',
      'โทรติดต่อ: 082-082-9484 ☎️'
    );
    tenantMsg = tenantMsg.join('\n');

    var ownerMsg = `มีการตรวจรับห้องเรียบร้อยแล้ว ✅
อาคาร: ${building || '-'}
ชั้น: ${floor || '-'}
ห้อง: ${roomId || '-'}
ผู้ตรวจ: ${inspector || '-'}

ไฟล์ตรวจรับ: ${pdfUrl}
โฟลเดอร์รูป: ${folder.getUrl()}`;

    if (inviteText) ownerMsg += '\n\n' + inviteText.trim();

    Logger.log("Invite debug → room: %s, building field: %s, detected: %s, invites: %s", roomId, building, bld || '-', JSON.stringify(availableBlds.map(function (e) { return e.letter + ':' + e.url; })));
    Logger.log("Invite text preview → %s", inviteText || '(empty)');

    if (lineId) {
      Logger.log("LINE notify → tenant %s", lineId);
      // 1) Send the normal text first
      sendLineMessage(tenantMsg, lineId);

      // 2) Then send the Flex invite
      if (chosen) {
        sendLineFlexInvite(chosen.letter, chosen.url, lineId);
      } else if (availableBlds.length >= 2) {
        var linkA = availableBlds.find(function (entry) { return entry.letter === 'A'; });
        var linkB = availableBlds.find(function (entry) { return entry.letter === 'B'; });
        // Fallback: unknown building -> show both buttons (A/B)
        if (linkA && linkB) {
          sendLineFlexTwoInvites(linkA.url, linkB.url, lineId);
        } else {
          // If we do not have both A/B, send all available as single buttons sequentially.
          availableBlds.forEach(function (entry) {
            sendLineFlexInvite(entry.letter, entry.url, lineId);
          });
        }
      } else if (availableBlds.length === 1) {
        sendLineFlexInvite(availableBlds[0].letter, availableBlds[0].url, lineId);
      }
    } else {
      Logger.log("⚠️ No LINE ID found for room " + roomId);
    }

    var managerId = PropertiesService.getScriptProperties().getProperty('LINE_MANAGER_ID') ||
                    PropertiesService.getScriptProperties().getProperty('LINE_OWNER_ID');
    if (managerId) {
      Logger.log("LINE notify → manager %s", managerId);
      sendLineMessage(ownerMsg, managerId);
      if (chosen) {
        sendLineFlexInvite(chosen.letter, chosen.url, managerId);
      } else if (availableBlds.length >= 2) {
        var linkA = availableBlds.find(function (entry) { return entry.letter === 'A'; });
        var linkB = availableBlds.find(function (entry) { return entry.letter === 'B'; });
        if (linkA && linkB) {
          sendLineFlexTwoInvites(linkA.url, linkB.url, managerId);
        } else if (availableBlds.length === 1) {
          sendLineFlexInvite(availableBlds[0].letter, availableBlds[0].url, managerId);
        }
      } else if (availableBlds.length === 1) {
        sendLineFlexInvite(availableBlds[0].letter, availableBlds[0].url, managerId);
      }
    } else {
      Logger.log("⚠️ No LINE_MANAGER_ID/LINE_OWNER_ID property set; manager not notified.");
    }


    // === final JSON response ===
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      roomId: roomId,
      wroteAreas: Array.from(areas),
      pdfUrl: pdfUrl,
      signatureUrl: signatureUrl
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/* ---------------- Room status updates ---------------- */
function updateRoomStatus(ssOrNull, roomId, statusValue) {
  if (!CONFIG.ROOM_STATUS_HEADER || !roomId || !statusValue) return;

  var ss = ssOrNull || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.ROOMS_SHEET);
  if (!sh) {
    Logger.log("⚠️ updateRoomStatus: missing sheet %s", CONFIG.ROOMS_SHEET);
    return;
  }

  var lastCol = sh.getLastColumn();
  if (!lastCol) return;

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxRoom = headers.indexOf(CONFIG.ROOM_HEADER);
  var idxStatus = headers.indexOf(CONFIG.ROOM_STATUS_HEADER);
  if (idxRoom < 0 || idxStatus < 0) {
    Logger.log("⚠️ updateRoomStatus: missing headers room=%s status=%s", CONFIG.ROOM_HEADER, CONFIG.ROOM_STATUS_HEADER);
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  var roomVals = sh.getRange(2, idxRoom + 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < roomVals.length; r++) {
    var rid = String(roomVals[r][0]).trim();
    if (!rid) continue;
    if (rid.toUpperCase() === roomId.toUpperCase()) {
      sh.getRange(r + 2, idxStatus + 1).setValue(statusValue);
      return;
    }
  }

  Logger.log("⚠️ updateRoomStatus: room %s not found in sheet %s", roomId, CONFIG.ROOMS_SHEET);
}

/* ---------------- Folder lookups ---------------- */
function getFolderIdForRoom(roomId) {
  // dedicated Check-in folder (if any)
  return lookupFolderId(roomId, [CONFIG.CHECKIN_HDR_ID]);
}

function getRoomFolderId(roomId) {
  // main room folder (fallback)
  return lookupFolderId(roomId, [CONFIG.ROOM_HDR_ID]);
}

function lookupFolderId(roomId, headerCandidates) {
  if (!roomId) return null;

  var cache = CacheService.getScriptCache();
  var key = 'ROOMFOLDER:' + headerCandidates.join(',') + ':' + roomId;
  var hit = cache.get(key);
  if (hit) return hit;

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sh = ss.getSheetByName(CONFIG.ROOMS_SHEET);
  if (!sh) return null;

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxOf = function (h) { return headers.indexOf(h); };

  var cRoom = idxOf(CONFIG.ROOM_HEADER);
  if (cRoom < 0) return null;

  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var rid = String(row[cRoom]).trim();
    if (!rid || rid.toUpperCase() !== roomId.toUpperCase()) continue;

    // try each candidate ID column
    for (var k = 0; k < headerCandidates.length; k++) {
      var h = headerCandidates[k];
      var c = idxOf(h);
      if (c >= 0) {
        var raw = String(row[c]).trim();
        var id  = extractFolderId(raw);
        if (id) { cache.put(key, id, 600); return id; }
      }
    }
    break;
  }
  return null;
}

/* ---------------- Utils ---------------- */
function extractFolderId(val) {
  if (!val) return null;
  if (/^[A-Za-z0-9_-]{20,}$/.test(val)) return val;                  // raw Id
  var m = String(val).match(/\/folders\/([A-Za-z0-9_-]{20,})/);      // from URL
  return m ? m[1] : null;
}

function getBuildingLetter_(building, roomId) {
  var available = [];
  if (OPENCHAT_LINKS) {
    Object.keys(OPENCHAT_LINKS).forEach(function (k) {
      if (OPENCHAT_LINKS[k]) available.push(String(k).toUpperCase());
    });
  }
  if (!available.length) return '';

  var scan = function (text) {
    if (!text) return '';
    var upper = String(text).toUpperCase();
    for (var i = 0; i < upper.length; i++) {
      var ch = upper.charAt(i);
      if (available.indexOf(ch) !== -1) return ch;
    }
    return '';
  };

  var fromRoom = scan(roomId);
  if (fromRoom) return fromRoom;
  return scan(building);
}

function guessExt(mime, name) {
  var m = (mime || '').toLowerCase();
  if (m.indexOf('png')  !== -1) return '.png';
  if (m.indexOf('jpeg') !== -1 || m.indexOf('jpg') !== -1) return '.jpg';
  if (m.indexOf('heic') !== -1) return '.heic';
  var i = name.lastIndexOf('.');
  return i > -1 ? name.slice(i) : '.bin';
}

/* ---------------- Quick tests (optional) ---------------- */
function testAuth() {
  try {
    // === Check Spreadsheet ===
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    Logger.log('✅ Spreadsheet: %s', ss.getName());

    var shRooms = ss.getSheetByName(CONFIG.ROOMS_SHEET);
    if (!shRooms) throw new Error('Rooms sheet not found');

    var headers = shRooms.getRange(1, 1, 1, shRooms.getLastColumn()).getValues()[0];
    Logger.log('Rooms headers: %s', JSON.stringify(headers));

    var need = [CONFIG.ROOM_HEADER, CONFIG.ROOM_HDR_ID, CONFIG.CHECKIN_HDR_ID, CONFIG.CHECKOUT_HDR_ID];
    need.forEach(function (h) {
      Logger.log('%s %s', headers.indexOf(h) >= 0 ? '✅' : '⚠️ MISSING', h);
    });

    // === Check Default Folder ===
    var f = DriveApp.getFolderById(CONFIG.DEFAULT_FOLDER);
    Logger.log('✅ Folder OK: %s (%s)', f.getName(), f.getUrl());

    // === Check Template Document ===
    var doc = DocumentApp.openById(CONFIG.TEMPLATE_ID);
    Logger.log('✅ Template Doc: %s (%s)', doc.getName(), 'https://docs.google.com/document/d/' + CONFIG.TEMPLATE_ID);

    Logger.log('🎉 testAuth passed');
  } catch (err) {
    Logger.log('❌ testAuth failed: %s', err && err.message ? err.message : String(err));
    throw err;
  }
}

function createInspectionPdf(row) {
  const docId = CONFIG.TEMPLATE_ID;
  const folder = DriveApp.getFolderById(CONFIG.DEFAULT_FOLDER);

  const ts = Utilities.formatDate(
    new Date(row.timestamp),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm'
  );

  const copy = DriveApp.getFileById(docId).makeCopy(
    `ตรวจห้อง_${row.roomId}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}`,
    folder
  );

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // === General placeholders ===
  body.replaceText('{{DATE}}', ts);
  body.replaceText('{{BUILDING}}', row.building || '');
  body.replaceText('{{FLOOR}}', row.floor || '');
  body.replaceText('{{ROOM_ID}}', row.roomId || '');
  body.replaceText('{{INSPECTOR}}', row.inspector || '');
  body.replaceText('{{GLOBAL_NOTES}}', row.globalNotes || '');

  // === Areas ===
  const areas = [
    'DOOR',
    'CURTAIN',
    'BED',
    'CHAIR_TABLE',
    'WARDROBE',
    'AC',
    'TOILET_SINK',
    'SHOWER_HEATER',
    'WALL_FLOOR_CEILING'
  ];

  areas.forEach(a => {
    body.replaceText(`{{${a}_STATUS}}`, row[`${a}_STATUS`] || '');
    body.replaceText(`{{${a}_NOTES}}`,  row[`${a}_NOTES`]  || '');

    const urls = row[`${a}_PHOTOS`] || [];
    const found = body.findText(`{{${a}_PHOTOS}}`);
    if (found) {
      const el = found.getElement();
      el.asText().setText(''); // clear placeholder

  urls.forEach(u => {
    try {
      const id = u.match(/[-\w]{25,}/)[0];
      const blob = DriveApp.getFileById(id).getBlob();   // <-- you need this!

      const img = el.getParent().insertInlineImage(
        el.getParent().getChildIndex(el),
        blob
      );

      // scale proportionally
      const maxWidth = 150;   // px
      const maxHeight = 150;  // px
      const w = img.getWidth();
      const h = img.getHeight();

      if (w > h) {
        img.setWidth(maxWidth);
        img.setHeight(h * (maxWidth / w));
      } else {
        img.setHeight(maxHeight);
        img.setWidth(w * (maxHeight / h));
      }

    } catch (err) {
      Logger.log(`⚠️ Could not insert image for ${a}: ${err}`);
    }
  });
    }
  });

  // === Signature ===
// === Signature ===
try {
  var sigBlob = null;

  if (row.signatureUrl) {
    var m1 = String(row.signatureUrl).match(/[-\w]{25,}/);
    if (m1) sigBlob = DriveApp.getFileById(m1[0]).getBlob();
  } else if (row.SIGNATURE_PHOTOS && row.SIGNATURE_PHOTOS.length) {
    var m2 = String(row.SIGNATURE_PHOTOS[0]).match(/[-\w]{25,}/);
    if (m2) sigBlob = DriveApp.getFileById(m2[0]).getBlob();
  }

  var found = body.findText('{{SIGNATURE}}');
  if (found) {
    var el = found.getElement();
    el.asText().setText('');

    if (sigBlob) {
      var img = el.getParent().insertInlineImage(el.getParent().getChildIndex(el), sigBlob);
      var maxW = 150, maxH = 80;
      var w = img.getWidth(), h = img.getHeight();
      if (w / h >= maxW / maxH) {
        img.setWidth(maxW);
        img.setHeight(h * (maxW / w));
      } else {
        img.setHeight(maxH);
        img.setWidth(w * (maxH / h));
      }
    } else {
      Logger.log("⚠️ No signature image available.");
    }
  } else {
    Logger.log("⚠️ {{SIGNATURE}} placeholder not found in template.");
  }
} catch (err) {
  Logger.log("⚠️ Could not insert signature: " + err);
}

  doc.saveAndClose();

  // === Export PDF ===
  const pdfBlob = copy.getAs(MimeType.PDF);
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return pdfFile.getUrl();
}


function sendLineMessage(msg, toUserId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) {
    Logger.log("⚠️ Missing LINE_TOKEN in script properties");
    return;
  }

  const payload = {
    to: toUserId,
    messages: [{ type: "text", text: msg }]
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    payload: JSON.stringify(payload)
  });
}

function sendLineFlexInvite(bld, link, toUserId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token || !toUserId || !link) return;

  const title = `เข้าร่วมกลุ่มผู้เช่า – อาคาร ${bld}`;
  const payload = {
    to: toUserId,
    messages: [{
      type: "flex",
      altText: title,
      contents: {
        type: "bubble",
        hero: {
          type: "image",
          url: "https://drive.google.com/thumbnail?id=1NTax9HfzpOIuio7v--YPNooKZiGl5Ueu&sz=w1200",
          size: "full",
          aspectRatio: "20:9",
          aspectMode: "cover"
        },
        body: {
          type: "box",
          layout: "vertical",
          spacing: "sm",
          contents: [
            { type: "text", text: title, weight: "bold", size: "md" },
            { type: "text", text: "พูดคุย/เตือนกันเอง • งานซ่อมหรือการเงินให้ทัก OA", size: "sm", color: "#666666", wrap: true }
          ]
        },
        footer: {
          type: "box",
          layout: "vertical",
          contents: [{
            type: "button",
            style: "primary",
            action: { type: "uri", label: "เข้าร่วมกลุ่ม", uri: link }
          }]
        }
      }
    }]
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function sendLineFlexTwoInvites(linkA, linkB, toUserId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token || !toUserId) return;

  const payload = {
    to: toUserId,
    messages: [{
      type: "flex",
      altText: "เข้าร่วมกลุ่มผู้เช่า – เลือกอาคาร",
      contents: {
        type: "bubble",
        body: {
          type: "box",
          layout: "vertical",
          contents: [
            { type: "text", text: "เข้าร่วมกลุ่มผู้เช่า", weight: "bold", size: "md" },
            { type: "text", text: "เลือกอาคารของคุณเพื่อเข้าร่วม", size: "sm", color: "#666666" }
          ]
        },
        footer: {
          type: "box",
          layout: "vertical",
          spacing: "sm",
          contents: [
            { type: "button", style: "primary", action: { type: "uri", label: "อาคาร A", uri: linkA } },
            { type: "button", style: "secondary", action: { type: "uri", label: "อาคาร B", uri: linkB } }
          ]
        }
      }
    }]
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function getLineIdForRoom(roomId) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const rooms = ss.getSheetByName("Rooms");
  const log   = ss.getSheetByName("Sheet1"); // Reservation log

  if (!rooms || !log) return null;

  // === Step 1: หา Reservation Code จาก Rooms ===
  const roomData = rooms.getDataRange().getValues(); // all rows
  const headersR = roomData[0];
  const cRoomId  = headersR.indexOf("RoomID");
  const cResCode = headersR.indexOf("Hg Code"); // ต้องแก้เป็นชื่อจริงใน sheet ของคุณ

  if (cRoomId < 0 || cResCode < 0) return null;

  let resCode = "";
  for (let i = 1; i < roomData.length; i++) {
    if (String(roomData[i][cRoomId]).trim().toUpperCase() === roomId.toUpperCase()) {
      resCode = String(roomData[i][cResCode]).trim();
      break;
    }
  }
  if (!resCode) return null;

  // === Step 2: หา LineID จาก Reservation Log ===
  const logData = log.getDataRange().getValues();
  const headersL = logData[0];
  const cCode    = headersL.indexOf("รหัสการจอง");
  const cLineId  = headersL.indexOf("Line User ID");

  if (cCode < 0 || cLineId < 0) return null;

  for (let i = 1; i < logData.length; i++) {
    if (String(logData[i][cCode]).trim() === resCode) {
      return String(logData[i][cLineId]).trim();
    }
  }
  return null;
}

function testCopyTemplate() {
  const docId = CONFIG.TEMPLATE_ID;
  const folder = DriveApp.getFolderById(CONFIG.DEFAULT_FOLDER);
  const copy = DriveApp.getFileById(docId).makeCopy("TEST_COPY", folder);
  Logger.log(copy.getUrl());
}
