// ─── BEWIJS UPLOADEN ─────────────────────────────────────────────────────────

function uploadBewijs(base64Data, mimeType, bestandsnaam, ritId, categorie, gedeeldeMap, gedeeldePersonenMap, gedeeldeEmailDomeinMap) {
  try {
    var email  = getEmail_();
    var folder = gedeeldeMap || DriveApp.getFolderById(CONFIG.DRIVE_MAP_ID);
    var blob   = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, bestandsnaam);
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var bestandKol = categorie === 'fiets' ? 9 : 11;
    var ext = '';
    var dotIdx = bestandsnaam.lastIndexOf('.');
    if (dotIdx !== -1) ext = bestandsnaam.slice(dotIdx).toLowerCase();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0]||'').toString() === ritId &&
          (data[i][1]||'').toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        var cel = sheet.getRange(i + 1, bestandKol);
        if (categorie === 'woonwerk') {
          var bestaandeNote = (cel.getNote()||'').trim();
          var bestaandeUrls = bestaandeNote ? bestaandeNote.split('\n').filter(Boolean) : [];
          if (bestaandeUrls.length === 0) {
            // Eerste bestand: URL direct in kolom K, naam zonder getal (WW)
            try { file.setName(bouwBestandsnaam_(data, i, email, categorie, 0) + ext); } catch(_) {}
            cel.setValue(fileUrl);
            cel.setNote(fileUrl);
          } else {
            // Extra bestand: volgnummer = aantal bestaande + 1
            var fileNumInRit = bestaandeUrls.length + 1;
            try { file.setName(bouwBestandsnaam_(data, i, email, categorie, fileNumInRit) + ext); } catch(_) {}
            // Bij het tweede bestand: hernoem het eerste van 'WW' naar 'WW-1'
            if (bestaandeUrls.length === 1) {
              try {
                var eersteIdMatch = bestaandeUrls[0].match(/\/d\/([^\/]+)\//);
                if (eersteIdMatch && eersteIdMatch[1]) {
                  var eersteBestand = DriveApp.getFileById(eersteIdMatch[1]);
                  var eersteNaam = eersteBestand.getName();
                  var eersteDot = eersteNaam.lastIndexOf('.');
                  var eersteExt = eersteDot !== -1 ? eersteNaam.slice(eersteDot).toLowerCase() : '';
                  eersteBestand.setName(bouwBestandsnaam_(data, i, email, categorie, 1) + eersteExt);
                }
              } catch(_) {}
            }
            // Voeg vervolgrij in direct na de laatste rij van deze rit
            var volgNr = 2;
            var insertNaRij = i + 1; // 1-based: standaard direct na de hoofdrij
            for (var r = 1; r < data.length; r++) {
              var rId = (data[r][0]||'').toString();
              if (rId === '_' + ritId + '_' + volgNr) volgNr++;
              // Laatste rij die bij deze rit hoort (hoofdrij of vervolgrij)
              if (rId === ritId || (rId.charAt(0) === '_' && rId.slice(1, rId.lastIndexOf('_')) === ritId)) {
                insertNaRij = r + 1;
              }
            }
            var vervolRij = new Array(12).fill('');
            vervolRij[0]  = '_' + ritId + '_' + volgNr;
            vervolRij[1]  = (data[i][1]||'').toString(); // email
            vervolRij[2]  = (data[i][2]||'').toString(); // naam — nodig voor kleurBeeldCheckKol_
            vervolRij[3]  = data[i][3]; // jaar — nodig voor bulk goedkeuring
            vervolRij[4]  = data[i][4]; // kwartaal — nodig voor bulk goedkeuring
            vervolRij[9]  = 'Bijlage ' + volgNr; // kolom J: verwijzing naar hoofdrij
            vervolRij[10] = fileUrl;
            vervolRij[11] = data[i][11]; // invoerdatum
            sheet.insertRowAfter(insertNaRij);
            var vervolRijNr = insertNaRij + 1;
            sheet.getRange(vervolRijNr, 1, 1, vervolRij.length).setValues([vervolRij]);
            try {
              var vervolNaam      = (data[i][2]||'').toString().trim();
              var vervolEmail     = (data[i][1]||'').toString().toLowerCase().trim();
              var _personenMap    = gedeeldePersonenMap    || bouwPersonenMap_(ss);
              var _emailDomeinMap = gedeeldeEmailDomeinMap || bouwEmailDomeinMap_(ss);
              var vervolDomein = ((_personenMap[vervolNaam]||{}).domein||'').trim();
              if (!vervolDomein && vervolEmail) vervolDomein = (_emailDomeinMap[vervolEmail]||'').trim();
              var isBeeldVervol = isGroeneCheckKol_(vervolNaam, vervolDomein);
              zetGoedkeuringInOpRij_(sheet, vervolRijNr, 'woonwerk', isBeeldVervol);
              var _ckkV = checkKolVoorSheet_(sheet.getName());
              if (_ckkV) sheet.getRange(vervolRijNr, _ckkV).insertCheckboxes().setBackground(isBeeldVervol ? '#c9f080' : null);
            } catch(_) {}
            cel.setNote(bestaandeNote + '\n' + fileUrl);
          }
        } else {
          try { file.setName(bouwBestandsnaam_(data, i, email, categorie, 1) + ext); } catch(_) {}
          cel.setValue(fileUrl);
        }
        break;
      }
    }
    return { ok: true, fileId: file.getId(), url: fileUrl };
  } catch(e) { Logger.log('❌ uploadBewijs: ' + e); return { ok: false, error: e.toString() }; }
}


// ─── BATCH UPLOAD (nieuwe methode: meerdere bestanden in één aanroep) ────────

function uploadBewijzenBatch(data, ritId, categorie) {
  try {
    var ss              = getSS_();
    var gedeeldeMap     = DriveApp.getFolderById(CONFIG.DRIVE_MAP_ID);
    var personenMap     = bouwPersonenMap_(ss);
    var emailDomeinMap  = bouwEmailDomeinMap_(ss);
    for (var i = 0; i < data.length; i++) {
      var res = uploadBewijs(data[i].base64, data[i].mimeType, data[i].naam, ritId, categorie, gedeeldeMap, personenMap, emailDomeinMap);
      if (!res.ok) return res;
    }
    return { ok: true };
  } catch(e) { Logger.log('❌ uploadBewijzenBatch: ' + e); return { ok: false, error: e.toString() }; }
}

function getDriveUploadConfig() {
  try { return { token: ScriptApp.getOAuthToken(), folderId: CONFIG.DRIVE_MAP_ID }; }
  catch(e) { return { token: '', folderId: CONFIG.DRIVE_MAP_ID }; }
}


// ─── DIRECTE DRIVE UPLOAD (snellere methode, met automatische fallback) ───────

function initDriveUpload(bestandsnaam, mimeType) {
  try {
    var token    = ScriptApp.getOAuthToken();
    var metadata = { name: bestandsnaam, parents: [CONFIG.DRIVE_MAP_ID] };
    var response = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json; charset=UTF-8',
          'X-Upload-Content-Type': mimeType
        },
        payload: JSON.stringify(metadata),
        muteHttpExceptions: true
      }
    );
    if (response.getResponseCode() !== 200) return { ok: false };
    var uploadUrl = response.getHeaders()['Location'];
    if (!uploadUrl) return { ok: false };
    return { ok: true, uploadUrl: uploadUrl };
  } catch(e) { Logger.log('❌ initDriveUpload: ' + e); return { ok: false, error: e.toString() }; }
}

function initDriveUploadsArray(bestandsInfo) {
  try {
    var uploadUrls = [];
    for (var i = 0; i < bestandsInfo.length; i++) {
      var res = initDriveUpload(bestandsInfo[i].naam, bestandsInfo[i].mimeType);
      if (!res.ok) return { ok: false };
      uploadUrls.push(res.uploadUrl);
    }
    return { ok: true, uploadUrls: uploadUrls };
  } catch(e) { Logger.log('❌ initDriveUploadsArray: ' + e); return { ok: false, error: e.toString() }; }
}

function registreerBewijsFileId(driveFileId, bestandsnaam, ritId, categorie) {
  try {
    var email   = getEmail_();
    var file    = DriveApp.getFileById(driveFileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = 'https://drive.google.com/file/d/' + driveFileId + '/view';

    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var bestandKol = categorie === 'fiets' ? 9 : 11;
    var ext = '';
    var dotIdx = bestandsnaam.lastIndexOf('.');
    if (dotIdx !== -1) ext = bestandsnaam.slice(dotIdx).toLowerCase();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0]||'').toString() === ritId &&
          (data[i][1]||'').toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        var cel = sheet.getRange(i + 1, bestandKol);
        if (categorie === 'woonwerk') {
          var bestaandeNote = (cel.getNote()||'').trim();
          var bestaandeUrls = bestaandeNote ? bestaandeNote.split('\n').filter(Boolean) : [];
          if (bestaandeUrls.length === 0) {
            try { file.setName(bouwBestandsnaam_(data, i, email, categorie, 0) + ext); } catch(_) {}
            cel.setValue(fileUrl);
            cel.setNote(fileUrl);
          } else {
            var fileNumInRit = bestaandeUrls.length + 1;
            try { file.setName(bouwBestandsnaam_(data, i, email, categorie, fileNumInRit) + ext); } catch(_) {}
            if (bestaandeUrls.length === 1) {
              try {
                var eersteIdMatch = bestaandeUrls[0].match(/\/d\/([^\/]+)\//);
                if (eersteIdMatch && eersteIdMatch[1]) {
                  var eersteBestand = DriveApp.getFileById(eersteIdMatch[1]);
                  var eersteNaam    = eersteBestand.getName();
                  var eersteDot     = eersteNaam.lastIndexOf('.');
                  var eersteExt     = eersteDot !== -1 ? eersteNaam.slice(eersteDot).toLowerCase() : '';
                  eersteBestand.setName(bouwBestandsnaam_(data, i, email, categorie, 1) + eersteExt);
                }
              } catch(_) {}
            }
            var volgNr      = 2;
            var insertNaRij = i + 1;
            for (var r = 1; r < data.length; r++) {
              var rId = (data[r][0]||'').toString();
              if (rId === '_' + ritId + '_' + volgNr) volgNr++;
              if (rId === ritId || (rId.charAt(0) === '_' && rId.slice(1, rId.lastIndexOf('_')) === ritId)) {
                insertNaRij = r + 1;
              }
            }
            var vervolRij = new Array(12).fill('');
            vervolRij[0]  = '_' + ritId + '_' + volgNr;
            vervolRij[1]  = (data[i][1]||'').toString();
            vervolRij[2]  = (data[i][2]||'').toString();
            vervolRij[3]  = data[i][3];
            vervolRij[4]  = data[i][4];
            vervolRij[9]  = 'Bijlage ' + volgNr;
            vervolRij[10] = fileUrl;
            vervolRij[11] = data[i][11];
            sheet.insertRowAfter(insertNaRij);
            var vervolRijNr = insertNaRij + 1;
            sheet.getRange(vervolRijNr, 1, 1, vervolRij.length).setValues([vervolRij]);
            try {
              var vervolNaam   = (data[i][2]||'').toString().trim();
              var vervolEmail  = (data[i][1]||'').toString().toLowerCase().trim();
              var vervolDomein = ((bouwPersonenMap_(ss)[vervolNaam]||{}).domein||'').trim();
              if (!vervolDomein && vervolEmail) vervolDomein = (bouwEmailDomeinMap_(ss)[vervolEmail]||'').trim();
              var isBeeldVervol = isGroeneCheckKol_(vervolNaam, vervolDomein);
              zetGoedkeuringInOpRij_(sheet, vervolRijNr, 'woonwerk', isBeeldVervol);
              var _ckkV = checkKolVoorSheet_(sheet.getName());
              if (_ckkV) sheet.getRange(vervolRijNr, _ckkV).insertCheckboxes().setBackground(isBeeldVervol ? '#c9f080' : null);
            } catch(_) {}
            cel.setNote(bestaandeNote + '\n' + fileUrl);
          }
        } else {
          try { file.setName(bouwBestandsnaam_(data, i, email, categorie, 1) + ext); } catch(_) {}
          cel.setValue(fileUrl);
        }
        break;
      }
    }
    return { ok: true, fileId: driveFileId, url: fileUrl };
  } catch(e) { Logger.log('❌ registreerBewijsFileId: ' + e); return { ok: false, error: e.toString() }; }
}

function registreerBewijsFilesArray(fileIds, bestandsnamen, ritId, categorie) {
  try {
    for (var i = 0; i < fileIds.length; i++) {
      var res = registreerBewijsFileId(fileIds[i], bestandsnamen[i], ritId, categorie);
      if (!res.ok) return res;
    }
    return { ok: true };
  } catch(e) { Logger.log('❌ registreerBewijsFilesArray: ' + e); return { ok: false, error: e.toString() }; }
}


// ─── BESTAAND WW-BEWIJS VERWIJDEREN ──────────────────────────────────────────

function verwijderWwBewijs(ritId, teVerwijderenUrl) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server is bezet, probeer het opnieuw.' };
  try {
    var email = getEmail_();
    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_('woonwerk'));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var data = sheet.getDataRange().getValues();

    // ① Zoek de hoofdrij van de rit
    var hoofdIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0]||'').toString() === ritId &&
          (data[i][1]||'').toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        hoofdIdx = i;
        break;
      }
    }
    if (hoofdIdx < 0) return { ok: false, error: 'Rit niet gevonden.' };

    // ② Lees alle URLs uit de notitie van de K-cel (kolom 11, 1-gebaseerd)
    var kCel     = sheet.getRange(hoofdIdx + 1, 11);
    var notitie  = (kCel.getNote() || '').trim();
    var alleUrls = notitie
      ? notitie.split('\n').map(function(u) { return u.trim(); }).filter(Boolean)
      : [];
    if (alleUrls.length === 0) {
      var celWrd = (kCel.getValue() || '').toString().trim();
      if (celWrd) alleUrls = [celWrd];
    }

    var delIdx = alleUrls.indexOf(teVerwijderenUrl);
    if (delIdx < 0) return { ok: false, error: 'Bewijsstuk niet gevonden in deze rit.' };

    // ③ Verwijder het Drive-bestand — stop direct als dit mislukt (sheet nog onaangepast)
    var fileId = fileIdUitUrl_(teVerwijderenUrl);
    if (!fileId) return { ok: false, error: 'Kan het Drive-bestand-ID niet bepalen.' };
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch(e) {
      return { ok: false, error: 'Drive-bestand verwijderen mislukt: ' + e.toString() };
    }

    // ④ Bouw de nieuwe URL-lijst (zonder het verwijderde bestand)
    var nieuweUrls = alleUrls.filter(function(_, j) { return j !== delIdx; });
    var n      = alleUrls.length;
    var nNieuw = nieuweUrls.length;

    // ⑤ Zoek vervolgrijen voor deze rit (ID-patroon: _ritId_N, N = 2..4)
    var vervolRijen = []; // [{ volgNr, sheetRij (1-gebaseerd) }]
    for (var r = 1; r < data.length; r++) {
      var rId    = (data[r][0]||'').toString();
      var rEmail = (data[r][1]||'').toString().toLowerCase().trim();
      if (rEmail !== email.toLowerCase().trim()) continue;
      var prefix = '_' + ritId + '_';
      if (rId.indexOf(prefix) !== 0) continue;
      var nr = parseInt(rId.slice(prefix.length));
      if (!isNaN(nr) && nr >= 2) {
        vervolRijen.push({ volgNr: nr, sheetRij: r + 1 });
      }
    }
    vervolRijen.sort(function(a, b) { return a.volgNr - b.volgNr; });

    // ⑥ Bepaal welke vervolgrij verdwijnt en welke hernummerd worden
    // delIdx === 0 → vervolgrij volgNr=2 verdwijnt (diens URL schuift naar de hoofdrij);
    //               volgNr=3,4 schuiven naar 2,3
    // delIdx >= 1 → vervolgrij volgNr=delIdx+1 verdwijnt;
    //               hogere volgNrs schuiven één omlaag
    var teVerwijderenVolgNr = (delIdx === 0) ? (n >= 2 ? 2 : -1) : (delIdx + 1);

    var rowsToDelete   = [];
    var rowsToRenumber = []; // [{ sheetRij, newVolgNr }]
    vervolRijen.forEach(function(vr) {
      if (teVerwijderenVolgNr > 0 && vr.volgNr === teVerwijderenVolgNr) {
        rowsToDelete.push(vr.sheetRij);
      } else if (teVerwijderenVolgNr > 0 && vr.volgNr > teVerwijderenVolgNr) {
        rowsToRenumber.push({ sheetRij: vr.sheetRij, newVolgNr: vr.volgNr - 1 });
      }
    });

    // ⑦ Werk de K-cel van de hoofdrij bij (waarde + notitie)
    if (nNieuw === 0) {
      kCel.setValue('');
      kCel.setNote('');
    } else {
      kCel.setValue(nieuweUrls[0]);
      kCel.setNote(nieuweUrls.join('\n'));
    }

    // ⑧ Hernummer de vervolgrijen (kolom A = nieuw ID, kolom J = "Bijlage N")
    //    Doe dit VOOR het verwijderen zodat rijcijfers nog correct zijn
    rowsToRenumber.forEach(function(u) {
      sheet.getRange(u.sheetRij, 1).setValue('_' + ritId + '_' + u.newVolgNr);
      sheet.getRange(u.sheetRij, 10).setValue('Bijlage ' + u.newVolgNr);
    });

    // ⑨ Verwijder de betrokken vervolgrij (van onder naar boven — is max. 1 rij)
    rowsToDelete.sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(rij) { sheet.deleteRow(rij); });

    // ⑩ Hernoem overblijvende Drive-bestanden (niet-kritisch — fouten negeren)
    if (nNieuw > 0) {
      var freshData    = sheet.getDataRange().getValues();
      var hoofdIdxNieuw = -1;
      for (var fi = 1; fi < freshData.length; fi++) {
        if ((freshData[fi][0]||'').toString() === ritId &&
            (freshData[fi][1]||'').toString().toLowerCase().trim() === email.toLowerCase().trim()) {
          hoofdIdxNieuw = fi;
          break;
        }
      }
      if (hoofdIdxNieuw >= 0) {
        for (var j = 0; j < nieuweUrls.length; j++) {
          var rFid = fileIdUitUrl_(nieuweUrls[j]);
          if (!rFid) continue;
          // fileNumInRit: 0 = enkel bestand ("WW"), 1+ = volgnummer ("WW-1", "WW-2", …)
          var newFileNum = nNieuw === 1 ? 0 : j + 1;
          try {
            var rFile     = DriveApp.getFileById(rFid);
            var oudeNaam  = rFile.getName();
            var dot       = oudeNaam.lastIndexOf('.');
            var ext       = dot !== -1 ? oudeNaam.slice(dot).toLowerCase() : '';
            var nieuweNaam = bouwBestandsnaam_(freshData, hoofdIdxNieuw, email, 'woonwerk', newFileNum) + ext;
            if (oudeNaam !== nieuweNaam) rFile.setName(nieuweNaam);
          } catch(_) {}
        }
      }
    }

    // ⑪ Invalideer de rittenCache voor dit kwartaal
    try {
      var jaar     = parseInt(data[hoofdIdx][3]) || 0;
      var kwartaal = parseInt(data[hoofdIdx][4]) || 0;
      if (jaar && kwartaal) invalideerRittenCache_(email, 'woonwerk', jaar, kwartaal);
    } catch(_) {}

    return { ok: true };
  } catch(e) {
    Logger.log('❌ verwijderWwBewijs: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
