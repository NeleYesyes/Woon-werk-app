// ─── RITTEN ──────────────────────────────────────────────────────────────────

function extractUrlUitFormule_(formule) {
  if (!formule) return '';
  var m = formule.match(/=HYPERLINK\("([^"]+)"/i);
  return m ? m[1] : '';
}

function invalideerRittenCache_(email, categorie, jaar, kwartaal) {
  try { CacheService.getUserCache().remove('ritten|' + categorie + '|' + email + '|' + jaar + '|' + kwartaal); } catch(_) {}
}

function getRitten(categorie, jaar, kwartaal) {
  try {
    var email    = getEmail_();
    var cacheKey = 'ritten|' + categorie + '|' + email + '|' + jaar + '|' + kwartaal;
    var cache    = CacheService.getUserCache();
    var cached   = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data      = sheet.getDataRange().getValues();
    var formules  = sheet.getDataRange().getFormulas();
    var notities  = (categorie === 'woonwerk' && sheet.getLastRow() >= 2)
      ? sheet.getRange(2, 11, data.length - 1, 1).getNotes() : null;
    var result   = [];

    var jaarKol = 3;
    var kwKol   = 4;

    for (var i = 1; i < data.length; i++) {
      if ((data[i][0]||'').toString().charAt(0) === '_') continue; // vervolgrij, overslaan
      if ((data[i][1]||'').toString().toLowerCase().trim() !== email.toLowerCase().trim()) continue;
      if (parseInt(data[i][jaarKol]) !== parseInt(jaar))     continue;
      if (parseInt(data[i][kwKol])   !== parseInt(kwartaal)) continue;

      var rawDatum;
      if (categorie === 'fiets') {
        rawDatum = data[i][9] || data[i][5]; // index 9 (nieuw), fallback index 5 (oud)
      } else if (categorie === 'woonwerk') {
        // Index 11 = invoerdatum (ISO "yyyy-MM-dd"), index 10 = bestand-URL (niet gebruiken!)
        var inv = data[i][11];
        if (inv) {
          rawDatum = inv;
        } else {
          // Fallback: parse maandStr "juni 2026" van index 5
          var ms = (data[i][5]||'').toString().trim().toLowerCase().split(' ');
          var mi = MAANDNAMEN_SERVER_.indexOf(ms[0]);
          var jj = parseInt(ms[1]) || parseInt(data[i][3]) || 0;
          rawDatum = (mi >= 0 && jj) ? new Date(jj, mi, 1) : null;
        }
      } else {
        rawDatum = data[i][5]; // dienst: index 5 = datum
      }
      var rit = { id: (data[i][0]||'').toString(), naam: (data[i][2]||'').toString(),
                  jaar: parseInt(data[i][jaarKol]), kwartaal: parseInt(data[i][kwKol]),
                  datum: formatDatum_(rawDatum), rijnummer: i + 1 };

      if (categorie === 'fiets') {
        rit.bestemming = (data[i][6]||'').toString();
        rit.km         = parseFloat(data[i][7]||0);
        var bu = extractUrlUitFormule_((formules[i]||[])[8]) || (data[i][8]||'').toString();
        rit.bestandUrl = bu; rit.bestandId = fileIdUitUrl_(bu);
      } else if (categorie === 'woonwerk') {
        rit.van     = (data[i][6]||'').toString();
        rit.naar    = (data[i][7]||'').toString();
        rit.vervoer = (data[i][8]||'').toString();
        rit.bedrag  = parseFloat(data[i][9]||0) || 0;
        var notitie = notities ? ((notities[i-1] || [''])[0] || '') : '';
        var bestandUrls;
        if (notitie.trim()) {
          bestandUrls = notitie.split('\n').map(function(u){ return u.trim(); }).filter(Boolean);
        } else {
          var buRaw = extractUrlUitFormule_((formules[i]||[])[10]) || (data[i][10]||'').toString().trim();
          bestandUrls = buRaw ? [buRaw] : [];
        }
        rit.bestandUrl  = bestandUrls[0] || '';
        rit.bestandId   = fileIdUitUrl_(bestandUrls[0] || '');
        rit.bestandUrls = bestandUrls;
        rit.bestandIds  = bestandUrls.map(fileIdUitUrl_).filter(Boolean);
      } else {
        rit.bestemming    = (data[i][6]||'').toString(); rit.doel = (data[i][7]||'').toString();
        rit.km            = parseFloat(data[i][8]||0);
        rit.parkeerBedrag = parseFloat(data[i][9]||0) || 0;
        var bu = extractUrlUitFormule_((formules[i]||[])[10]) || (data[i][10]||'').toString();
        rit.bestandUrl = bu; rit.bestandId = fileIdUitUrl_(bu);
      }
      result.push(rit);
    }
    result.sort(function(a,b){ return a.datum.localeCompare(b.datum); });
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(_) {}
    return result;
  } catch(e) { Logger.log('❌ getRitten: ' + e); return []; }
}

function getRittenAlle(jaar, kwartaal) {
  return {
    fiets:    getRitten('fiets',    jaar, kwartaal),
    woonwerk: getRitten('woonwerk', jaar, kwartaal),
    dienst:   getRitten('dienst',   jaar, kwartaal),
  };
}

function saveRit(rit, categorie) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server is bezet, probeer het opnieuw.' };
  try {
    var validatie = validateEntryDate(rit.datum);
    if (!validatie.valid) return { ok: false, error: validatie.error };

    var email   = getEmail_();
    var ss      = getSS_();
    var sheet   = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var kwartaal = validatie.quarter;
    var jaar     = validatie.year;
    var persoon  = getPersonData_(email, ss);
    var isBeeld  = isGroeneCheckKol_(persoon.naam || rit.naam || '', persoon.domein || '');
    var id       = rit.id || generateId_();
    var rij;

    var bestandIdx;
    if (categorie === 'fiets') {
      var maandStr = formatMaandJaarServer_(rit.datum);
      rij        = [id, email, persoon.naam || rit.naam || '', jaar, kwartaal, maandStr, rit.bestemming || '', parseFloat(rit.km||0), '', rit.datum];
      bestandIdx = 8;
      // indices: 0   1    2                             3     4         5          6                   7                      8   9
    } else if (categorie === 'woonwerk') {
      var maandStrWw = formatMaandJaarServer_(rit.datum);
      rij        = [id, email, persoon.naam || rit.naam || '', jaar, kwartaal, maandStrWw, rit.van||'', rit.naar||'', rit.vervoer||'', parseFloat(rit.bedrag||0)||'', '', rit.datum];
      bestandIdx = 10;
      // indices: 0   1    2                             3     4         5            6            7             8               9                 10  11
    } else {
      var vandaag = Utilities.formatDate(getEffectiveDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      rij        = [id, email, persoon.naam || rit.naam || '', jaar, kwartaal, rit.datum, rit.bestemming||'', rit.doel||'', parseFloat(rit.km||0), parseFloat(rit.parkeerBedrag||0)||'', '', vandaag];
      bestandIdx = 10;
      // indices: 0   1    2                             3     4         5          6                   7             8                      9                           10  11
    }

    if (rit.id) {
      var alleData = sheet.getDataRange().getValues();
      var gevRij   = -1; var bestaandLink = '';
      for (var r = 1; r < alleData.length; r++) {
        if ((alleData[r][0]||'').toString() === rit.id &&
            (alleData[r][1]||'').toString().toLowerCase().trim() === email.toLowerCase().trim()) {
          gevRij       = r + 1;
          bestaandLink = (alleData[r][bestandIdx]||'').toString();
          if (categorie === 'dienst' && (alleData[r][11]||'').toString().trim() !== '') {
            rij[11] = (alleData[r][11]||'').toString();  // preserveer originele invoerdatum
          }
          break;
        }
      }
      if (gevRij > 0) {
        var bCelEdit = sheet.getRange(gevRij, bestandIdx + 1);
        if (categorie === 'woonwerk') {
          // Bewaar note (bevat alle URLs) en plain URL-waarde; geen rich text nodig
          var notitieEdit = bCelEdit.getNote();
          rij[bestandIdx] = (bCelEdit.getValue()||'').toString().trim()
                            || extractUrlUitFormule_(bCelEdit.getFormula());
          sheet.getRange(gevRij, 1, 1, rij.length).setValues([rij]);
          if (notitieEdit) bCelEdit.setNote(notitieEdit);
        } else {
          var savedFormEdit = bCelEdit.getFormula();
          rij[bestandIdx] = bestaandLink;
          sheet.getRange(gevRij, 1, 1, rij.length).setValues([rij]);
          if (savedFormEdit) bCelEdit.setFormula(savedFormEdit);
        }
        zetGoedkeuringInOpRij_(sheet, gevRij, categorie, isBeeld);
        if (categorie === 'dienst') sheet.getRange(gevRij, 10).setNumberFormat('0.00');
      } else {
        var nieuweRij1 = appendRijNaData_(sheet, rij);
        zetGoedkeuringInOpRij_(sheet, nieuweRij1, categorie, isBeeld);
        if (categorie === 'dienst') sheet.getRange(nieuweRij1, 10).setNumberFormat('0.00');
      }
    } else {
      var nieuweRij2 = appendRijNaData_(sheet, rij);
      zetGoedkeuringInOpRij_(sheet, nieuweRij2, categorie, isBeeld);
      if (categorie === 'dienst') sheet.getRange(nieuweRij2, 10).setNumberFormat('0.00');
    }

    invalideerRittenCache_(email, categorie, jaar, kwartaal);
    try { sorteerSheet_(sheet, bouwEmailNaamMap_(ss)); } catch(_) {}
    var _ckk = checkKolVoorSheet_(sheet.getName()); if (_ckk) { try { kleurBeeldCheckKol_(ss, sheet, _ckk); } catch(_e) {} }
    return { ok: true, id: id };
  } catch(e) { Logger.log('❌ saveRit: ' + e); return { ok: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteRit(id, categorie) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Je bent net van hetzelfde idee als je collega, dus probeer je binnen een minuutje opnieuw? Dankjewel!' };
  try {
    var email = getEmail_();
    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var data    = sheet.getDataRange().getValues();
    var jaar = 0, kwartaal = 0;

    // Verzamel bestand-IDs vóór het wissen (daarna zijn de celwaarden weg)
    var bestandKolIdx = (categorie === 'fiets') ? 8 : 10;
    var teTrashIds = [];
    for (var j = 1; j < data.length; j++) {
      var jId    = (data[j][0]||'').toString();
      var jEmail = (data[j][1]||'').toString().toLowerCase().trim();
      if (jEmail !== email.toLowerCase().trim()) continue;
      var _jPfx = '_' + id + '_'; var _jNr = parseInt(jId.slice(_jPfx.length));
      if (jId !== id && !(jId.indexOf(_jPfx) === 0 && !isNaN(_jNr) && _jNr >= 2)) continue;
      var fid = fileIdUitUrl_((data[j][bestandKolIdx]||'').toString().trim());
      if (fid) teTrashIds.push(fid);
    }

    // Verwijder vervolgrijen (woonwerk multi-bewijs) én hoofdrij, van onder naar boven
    for (var i = data.length - 1; i >= 1; i--) {
      var rijId    = (data[i][0]||'').toString();
      var rijEmail = (data[i][1]||'').toString().toLowerCase().trim();
      if (rijEmail !== email.toLowerCase().trim()) continue;
      var isHoofd  = rijId === id;
      var _vPfx = '_' + id + '_'; var _vNr = parseInt(rijId.slice(_vPfx.length));
      var isVervol = rijId.indexOf(_vPfx) === 0 && !isNaN(_vNr) && _vNr >= 2;
      if (!isHoofd && !isVervol) continue;
      if (isHoofd) { jaar = parseInt(data[i][3]); kwartaal = parseInt(data[i][4]); }
      sheet.deleteRow(i + 1);
    }

    // Stuur bijbehorende Drive-bestanden naar de prullenbak (defensief per bestand)
    for (var t = 0; t < teTrashIds.length; t++) {
      try { DriveApp.getFileById(teTrashIds[t]).setTrashed(true); } catch(_) {}
    }
    if (jaar) invalideerRittenCache_(email, categorie, jaar, kwartaal);
    try { sorteerSheet_(sheet, bouwEmailNaamMap_(ss)); } catch(_) {}
    var _ckk2 = checkKolVoorSheet_(sheet.getName()); if (_ckk2) { try { kleurBeeldCheckKol_(ss, sheet, _ckk2); } catch(_e) {} }
    return { ok: true };
  } catch(e) { Logger.log('❌ deleteRit: ' + e); return { ok: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}
