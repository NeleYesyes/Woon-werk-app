// ─── BEWIJS UPLOADEN ─────────────────────────────────────────────────────────

function uploadBewijs(base64Data, mimeType, bestandsnaam, ritId, categorie) {
  try {
    var email  = getEmail_();
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_MAP_ID);
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
    return { ok: true, fileId: file.getId(), url: fileUrl };
  } catch(e) { Logger.log('❌ uploadBewijs: ' + e); return { ok: false, error: e.toString() }; }
}


// ─── BATCH UPLOAD (nieuwe methode: meerdere bestanden in één aanroep) ────────

function uploadBewijzenBatch(data, ritId, categorie) {
  try {
    for (var i = 0; i < data.length; i++) {
      var res = uploadBewijs(data[i].base64, data[i].mimeType, data[i].naam, ritId, categorie);
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
