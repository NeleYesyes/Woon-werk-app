// ─── TRIGGERS ────────────────────────────────────────────────────────────────

function verwerkKwartaalSlot_(kwSheet, hdrRij, lockWaarde) {
  var isLocked = (lockWaarde === true || lockWaarde === 'TRUE' || lockWaarde === 'true');
  kwSheet.getRange(hdrRij, 11).setBackground(isLocked ? '#bbf7d0' : '#dbeafe');
  var jaar = parseInt(kwSheet.getRange('B2').getValue());
  var kwLabel = kwSheet.getRange(hdrRij, 1).getValue().toString();
  var kwMatch = kwLabel.match(/Kwartaal\s+(\d)/);
  if (!kwMatch || isNaN(jaar)) return;
  var kw = parseInt(kwMatch[1]);
  var props = PropertiesService.getScriptProperties();
  if (isLocked) props.setProperty('lock_' + jaar + '_' + kw, 'true');
  else props.deleteProperty('lock_' + jaar + '_' + kw);
  var ss = getSS_();
  if (isLocked) verbergRitRijenVoorKwartaal_(ss, jaar, kw);
  else toonRitRijenVoorKwartaal_(ss, jaar, kw);
}

function verbergRitRijenVoorKwartaal_(ss, jaar, kw) {
  [CONFIG.SHEETS.FIETSVERGOEDING, CONFIG.SHEETS.WOON_WERK, CONFIG.SHEETS.DIENSTVERPLAATSING].forEach(function(naam) {
    var sheet = ss.getSheetByName(naam);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

    // Eerste pass: verzamel IDs van hoofdrijen die overeenkomen met jaar/kw
    var teVerbIds = {};
    for (var r = 0; r < data.length; r++) {
      var idStr = (data[r][0]||'').toString();
      if (idStr.charAt(0) === '_') continue;
      if (parseInt(data[r][3]) === jaar && parseInt(data[r][4]) === kw) teVerbIds[idStr] = true;
    }

    // Tweede pass: verberg hoofdrijen én hun vervolrijen (multi-bewijs)
    for (var r2 = 0; r2 < data.length; r2++) {
      var rid = (data[r2][0]||'').toString();
      var verberg = teVerbIds[rid] === true;
      if (!verberg && rid.charAt(0) === '_') {
        var m = rid.match(/^_(.+)_\d+$/);
        verberg = !!(m && teVerbIds[m[1]]);
      }
      if (verberg) sheet.hideRows(r2 + 2);
    }
  });
}

function toonRitRijenVoorKwartaal_(ss, jaar, kw) {
  [CONFIG.SHEETS.FIETSVERGOEDING, CONFIG.SHEETS.WOON_WERK, CONFIG.SHEETS.DIENSTVERPLAATSING].forEach(function(naam) {
    var sheet = ss.getSheetByName(naam);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

    // Eerste pass: verzamel IDs van hoofdrijen die overeenkomen met jaar/kw
    var teTonenIds = {};
    for (var r = 0; r < data.length; r++) {
      var idStr = (data[r][0]||'').toString();
      if (idStr.charAt(0) === '_') continue;
      if (parseInt(data[r][3]) === jaar && parseInt(data[r][4]) === kw) teTonenIds[idStr] = true;
    }

    // Tweede pass: toon hoofdrijen én hun vervolrijen (multi-bewijs)
    for (var r2 = 0; r2 < data.length; r2++) {
      var rid2 = (data[r2][0]||'').toString();
      var toon = teTonenIds[rid2] === true;
      if (!toon && rid2.charAt(0) === '_') {
        var m2 = rid2.match(/^_(.+)_\d+$/);
        toon = !!(m2 && teTonenIds[m2[1]]);
      }
      if (toon) sheet.showRows(r2 + 2);
    }
  });
}

function installeerJaarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onEditJaar' || fn === 'controleerKwartaaloverzicht_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditJaar').forSpreadsheet(CONFIG.SPREADSHEET_ID).onEdit().create();
  ScriptApp.newTrigger('controleerKwartaaloverzicht_').timeBased().everyHours(1).create();
  Logger.log('✅ Triggers geïnstalleerd.');
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Triggers geïnstalleerd.\n\n' +
      '• Bewerkingen in de tabbladen worden automatisch verwerkt.\n' +
      '• Het Kwartaaloverzicht wordt elk uur automatisch gecontroleerd en hersteld indien nodig.'
    );
  } catch(_) {}
}

function onEditJaar(e) {
  try {
    var sheet     = e.range.getSheet();
    var sheetNaam = sheet.getName();

    // Goedkeuring in rit-tabbladen: kleur onmiddellijk aanpassen bij klik
    var goedkKolMap = {};
    goedkKolMap[CONFIG.SHEETS.FIETSVERGOEDING]    = 11;
    goedkKolMap[CONFIG.SHEETS.WOON_WERK]          = 13;
    goedkKolMap[CONFIG.SHEETS.DIENSTVERPLAATSING] = 13;
    var beeldKolMap = {};
    beeldKolMap[CONFIG.SHEETS.FIETSVERGOEDING]    = 13;
    beeldKolMap[CONFIG.SHEETS.WOON_WERK]          = 15;
    beeldKolMap[CONFIG.SHEETS.DIENSTVERPLAATSING] = 15;
    var checkKolMap = {};
    checkKolMap[CONFIG.SHEETS.FIETSVERGOEDING]    = 12;
    checkKolMap[CONFIG.SHEETS.WOON_WERK]          = 14;
    checkKolMap[CONFIG.SHEETS.DIENSTVERPLAATSING] = 14;
    var bestandKolMap = {};
    bestandKolMap[CONFIG.SHEETS.FIETSVERGOEDING]    = 9;   // kolom I (index 8)
    bestandKolMap[CONFIG.SHEETS.WOON_WERK]          = 11;  // kolom K (index 10)
    bestandKolMap[CONFIG.SHEETS.DIENSTVERPLAATSING] = 11;  // kolom K (index 10)

    // Rij 1: Nico bulk-checkbox (niet-Beeld)
    if (goedkKolMap[sheetNaam] !== undefined &&
        e.range.getColumn() === goedkKolMap[sheetNaam] &&
        e.range.getRow() === 1 &&
        e.range.getValue() === true) {
      var n = bulkGoedkeurenVoorTabblad_(sheet, goedkKolMap[sheetNaam]);
      e.range.setValue(false);
      var kwLabelN = bepaalActieveKwartalen_().map(function(k){return 'K'+k.kwartaal+' '+k.jaar;}).join(' + ');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        n + ' rijen goedgekeurd voor ' + kwLabelN + '.',
        'Bulk goedkeuring', 5);
      return;
    }

    // Rij 1: Inse/Beeld bulk-checkbox
    if (beeldKolMap[sheetNaam] !== undefined &&
        e.range.getColumn() === beeldKolMap[sheetNaam] &&
        e.range.getRow() === 1 &&
        e.range.getValue() === true) {
      var nb = bulkGoedkeurenBeeldVoorTabblad_(sheet, beeldKolMap[sheetNaam]);
      e.range.setValue(false);
      var kwLabelNb = bepaalActieveKwartalen_().map(function(k){return 'K'+k.kwartaal+' '+k.jaar;}).join(' + ');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        nb + ' rijen goedgekeurd (Beeld) voor ' + kwLabelNb + '.',
        'Bulk goedkeuring Beeld', 5);
      return;
    }

    // Individuele klik in Nico-goedkeuringskolom (rij > 1)
    if (goedkKolMap[sheetNaam] !== undefined &&
        e.range.getColumn() === goedkKolMap[sheetNaam] &&
        e.range.getRow() > 1) {
      var huidig  = (e.range.getValue()||'').toString().trim();
      var eerste  = huidig.toLowerCase().charAt(0);
      var ritVal  = (eerste === 'j') ? 'Ja, goedgekeurd'
                  : (eerste === 'n') ? 'Nee, afgekeurd'
                  : (eerste === 'i') ? 'Ingediend'
                  : huidig;
      if (ritVal !== huidig) e.range.setValue(ritVal);
      var ritKleur = goedkeuringStatusKleur_(ritVal);
      e.range.setBackground(ritKleur.bg).setFontColor(ritKleur.txt);
      var checkCel = sheet.getRange(e.range.getRow(), checkKolMap[sheetNaam]);
      if (ritVal === 'Nee, afgekeurd') {
        checkCel.insertCheckboxes().setValue(true).setBackground(null).setFontColor(null);
      } else if (ritVal === 'Ja, goedgekeurd') {
        checkCel.insertCheckboxes().setBackground(null).setFontColor(null).setValue(false);
      } else {
        checkCel.setBackground(null).setFontColor(null);
      }
      verversKwartaaloverzichtAlsBestaat_();
      return;
    }

    // Individuele klik in Beeld-goedkeuringskolom (rij > 1)
    if (beeldKolMap[sheetNaam] !== undefined &&
        e.range.getColumn() === beeldKolMap[sheetNaam] &&
        e.range.getRow() > 1) {
      var huidigB = (e.range.getValue()||'').toString().trim();
      var eersteB = huidigB.toLowerCase().charAt(0);
      var ritValB = (eersteB === 'j') ? 'Ja, goedgekeurd'
                  : (eersteB === 'n') ? 'Nee, afgekeurd'
                  : (eersteB === 'i') ? 'Ingediend'
                  : huidigB;
      if (ritValB !== huidigB) e.range.setValue(ritValB);
      var ritKleurB = goedkeuringStatusKleur_(ritValB);
      e.range.setBackground(ritKleurB.bg).setFontColor(ritKleurB.txt);
      var checkCelB = sheet.getRange(e.range.getRow(), checkKolMap[sheetNaam]);
      if (ritValB === 'Nee, afgekeurd') {
        checkCelB.insertCheckboxes().setValue(true).setBackground(null).setFontColor(null);
      } else if (ritValB === 'Ja, goedgekeurd') {
        checkCelB.insertCheckboxes().setBackground(null).setFontColor(null).setValue(false);
      } else {
        checkCelB.setBackground(null).setFontColor(null);
      }
      verversKwartaaloverzichtAlsBestaat_();
      return;
    }

    // Manuele invoer/correctie in rit-tabbladen
    if (goedkKolMap[sheetNaam] !== undefined && e.range.getRow() > 1 &&
        e.range.getColumn() !== goedkKolMap[sheetNaam] &&
        e.range.getColumn() !== checkKolMap[sheetNaam] &&
        e.range.getColumn() !== beeldKolMap[sheetNaam]) {
      // Maandkolom (kolom 6) in Fiets en Woon-Werk: zet tekstopmaak zodat Sheets de waarde
      // niet als datum herkent en rechts uitlijnt
      var isMaandBlad = (sheetNaam === CONFIG.SHEETS.FIETSVERGOEDING || sheetNaam === CONFIG.SHEETS.WOON_WERK);
      if (isMaandBlad && e.range.getColumn() === 6) {
        e.range.setNumberFormat('@').setHorizontalAlignment('left');
      }

      var manCateg     = (sheetNaam === CONFIG.SHEETS.FIETSVERGOEDING) ? 'fiets'
                       : (sheetNaam === CONFIG.SHEETS.WOON_WERK)       ? 'woonwerk' : 'dienst';
      var bestandKol   = bestandKolMap[sheetNaam]; // rechterrand van de inhoudskolommen (I voor Fiets, K voor Woon-Werk/Dienst)
      // Verplichte rechtergrens voor de "rij is volledig"-controle, per tabblad instelbaar:
      // bij Dienstverplaatsing zijn J (parkeerkost) en K (bestand) optioneel, dus daar ligt
      // de grens op I in plaats van op bestandKol (K).
      var verplichtKol = (sheetNaam === CONFIG.SHEETS.DIENSTVERPLAATSING) ? 9 : bestandKol;
      var ssManueel    = getSS_();
      var manRijStart  = e.range.getRow();
      var manRijEinde  = manRijStart + e.range.getNumRows() - 1;
      var manMoetSorteren  = false;
      var manMoetVerversen = false;

      // Bij plakken kan het edit-event meerdere rijen beslaan: elke geraakte rij apart
      // controleren (typen en plakken dus gelijk behandelen) en celwaarden vers uitlezen.
      for (var manRij = manRijStart; manRij <= manRijEinde; manRij++) {
        var manData     = sheet.getRange(manRij, 1, 1, 5).getValues()[0];
        var manId       = (manData[0]||'').toString().trim();
        var manNaam     = (manData[2]||'').toString().trim();
        var manJaar     = manData[3];
        var manKw       = manData[4];
        var manEmail    = (manData[1]||'').toString().trim();

        // Handmatig ingetikte rij zonder verborgen kolommen A (id) en B (e-mail): aanvullen
        if (manNaam && !manId && !manEmail) {
          sheet.getRange(manRij, 1).setValue(generateId_());
          var manGevonden = vindEmailVoorNaam_(getSS_(), manNaam);
          if (manGevonden.email) {
            sheet.getRange(manRij, 2).setValue(manGevonden.email);
            manEmail = manGevonden.email;
          }
          // Match gevonden (ook bij omgekeerde volgorde, bv. "Marjan Depuydt"): kolom C
          // rechtzetten naar de officiële schrijfwijze uit Personeelsgegevens kolom B.
          if (manGevonden.naam && manGevonden.naam !== manNaam) {
            sheet.getRange(manRij, 3).setValue(manGevonden.naam);
            manNaam = manGevonden.naam;
          }
        }

        if (!(manNaam && manJaar && manKw)) continue;
        manMoetVerversen = true;

        // Goedkeuringswaarde: controleer beide kolommen
        var manGoedkVal = (sheet.getRange(manRij, goedkKolMap[sheetNaam]).getValue()||'').toString().trim()
                       || (sheet.getRange(manRij, beeldKolMap[sheetNaam]).getValue()||'').toString().trim();

        // Rij "klaar om af te werken": inhoudskolommen C t/m verplichtKol (I voor Fiets en
        // Dienst, K voor Woon-Werk) staan allemaal vol — niet langer afhankelijk van welke
        // kolom precies bewerkt werd, zodat handmatige invoer/plakken zonder de optionele
        // kolommen ook sorteert.
        var manCompleet = true;
        if (verplichtKol) {
          var manInhoud = sheet.getRange(manRij, 3, 1, verplichtKol - 2).getValues()[0];
          for (var mi = 0; mi < manInhoud.length; mi++) {
            if ((manInhoud[mi]||'').toString().trim() === '') { manCompleet = false; break; }
          }
        }

        if (manCompleet && !manGoedkVal) {
          // Rij is volledig en nog niet afgewerkt: goedkeuring instellen + nadien sorteren
          var manDomein = ((bouwPersonenMap_(ssManueel)[manNaam]||{}).domein||'').trim();
          if (!manDomein && manEmail) manDomein = (bouwEmailDomeinMap_(ssManueel)[manEmail.toLowerCase()]||'').trim();
          zetGoedkeuringInOpRij_(sheet, manRij, manCateg, isGroeneCheckKol_(manNaam, manDomein));
          manMoetSorteren = true;
        } else {
          // Nog niet volledig, of al afgewerkt: enkel cache wissen
          if (manEmail) invalideerRittenCache_(manEmail, manCateg, parseInt(manJaar), parseInt(manKw));
        }
      }

      if (manMoetSorteren) {
        // Eén keer sorteren ná de hele lus: sorteerSheet_ herschikt alle rijen op het tabblad,
        // dus tussentijds sorteren per rij zou de overige rij-indices laten verschuiven.
        sorteerSheet_(sheet, bouwEmailNaamMap_(ssManueel));
        var _ckkMan = checkKolVoorSheet_(sheetNaam); if (_ckkMan) { try { kleurBeeldCheckKol_(ssManueel, sheet, _ckkMan); } catch(_e) {} }
        SpreadsheetApp.getActiveSpreadsheet().toast('Rijen hergesorteerd op naam — controleer de positie van de nieuwe rij.', 'Alfabetisch gesorteerd ↕', 4);
      }
      if (manMoetVerversen) verversKwartaaloverzichtAlsBestaat_();
    }

    // Aanpak: klik op gekleurde knop (col 1) om sectie te tonen
    if (sheetNaam === 'Aanpak' && e.range.getColumn() === 1 && e.range.getValue() === true) {
      var note = sheet.getRange(1, 1).getNote();
      if (!note) return;
      var meta; try { meta = JSON.parse(note); } catch(_) { return; }
      var keuze = null;
      meta.forEach(function(tab) { if (tab.buttonRij === e.range.getRow()) keuze = tab.naam; });
      if (keuze) { toonAanpakSectie_(sheet, keuze, meta); return; }
    }

    if (sheetNaam !== 'Kwartaaloverzicht') return;

    if (e.range.getRow() === 2 && e.range.getColumn() === 2) { maakKwartaaloverzicht(); return; }

    if (e.range.getColumn() === 8 && e.range.getRow() > 3) {
      var hdrRij    = e.range.getRow();
      var dataStart = parseInt(sheet.getRange(hdrRij, 9).getValue());
      var dataEinde = parseInt(sheet.getRange(hdrRij, 10).getValue());
      if (!dataStart || !dataEinde || dataStart < 1 || dataEinde < dataStart) {
        var naamCelH = (sheet.getRange(hdrRij, 2).getValue()||'').toString().trim();
        if (naamCelH === 'Logghe Nico') {
          PropertiesService.getScriptProperties().setProperty('tariefFiets_Logghe Nico', (parseFloat(e.range.getValue())||0).toString());
          return;
        }
        var titelCelA = (sheet.getRange(hdrRij - 1, 1).getValue()||'').toString();
        var tariefVal = parseFloat(e.range.getValue()) || 0;
        if (titelCelA.indexOf('Fietsvergoeding') > -1) {
          PropertiesService.getScriptProperties().setProperty('tariefFiets', tariefVal.toString());
          maakKwartaaloverzicht();
          return;
        }
        if (titelCelA.indexOf('Dienstverplaatsing') > -1) {
          PropertiesService.getScriptProperties().setProperty('tariefDienst', tariefVal.toString());
          maakKwartaaloverzicht();
          return;
        }
        return;
      }
      if (e.range.getValue() === true) {
        sheet.showRows(dataStart, dataEinde - dataStart + 1);
        // Geopend kwartaal: oranje header + instructietekst kolom 12
        sheet.getRange(hdrRij, 1, 1, 7).setBackground('#f59e0b').setFontColor('#7c2d12').setFontSize(11);
        sheet.getRange(hdrRij, 12)
          .setValue('← Helemaal klaar met kwartaal? Klik checkbox, graag!')
          .setFontSize(9).setFontStyle('italic').setFontColor('#1e40af')
          .setHorizontalAlignment('left').setVerticalAlignment('middle');
        // Radio-button: sluit alle andere kwartalen en herstel hun headerkleur + wis instructietekst
        var lastRowKw = sheet.getLastRow();
        if (lastRowKw > 3) {
          var colHIJK = sheet.getRange(4, 8, lastRowKw - 3, 4).getValues(); // H, I, J, K
          for (var kri = 0; kri < colHIJK.length; kri++) {
            var absRij = kri + 4;
            if (absRij === hdrRij) continue;
            var ds = parseInt(colHIJK[kri][1]), de = parseInt(colHIJK[kri][2]);
            if (!ds || !de || ds < 1 || de < ds) continue;
            if (colHIJK[kri][0] === true) {
              sheet.getRange(absRij, 8).setValue(false);
              sheet.hideRows(ds, de - ds + 1);
            }
            var gesloten = colHIJK[kri][3] === true;
            sheet.getRange(absRij, 1, 1, 7)
              .setBackground(gesloten ? '#64748b' : '#ffffff')
              .setFontColor(gesloten ? '#ffffff' : '#94a3b8')
              .setFontSize(10);
            sheet.getRange(absRij, 12).clearContent();
          }
        }
      } else {
        // Gesloten kwartaal: herstel headerkleur en wis instructietekst
        var geslotenDit = sheet.getRange(hdrRij, 11).getValue() === true;
        sheet.getRange(hdrRij, 1, 1, 7)
          .setBackground(geslotenDit ? '#64748b' : '#ffffff')
          .setFontColor(geslotenDit ? '#ffffff' : '#94a3b8')
          .setFontSize(10);
        sheet.getRange(hdrRij, 12).clearContent();
        sheet.hideRows(dataStart, dataEinde - dataStart + 1);
      }
    }

    if (e.range.getColumn() === 11 && e.range.getRow() > 3) {
      var nieuweWaarde = (e.range.getValue()||'').toString().trim();
      var colIWaarde   = sheet.getRange(e.range.getRow(), 9).getValue();
      var colINum      = parseInt(colIWaarde);
      // Header-rij: kolom I bevat een rijnummer (positief getal), geen e-mail
      if (!isNaN(colINum) && colINum > 0) {
        verwerkKwartaalSlot_(sheet, e.range.getRow(), nieuweWaarde);
        return;
      }
      // Melding-rij: kolom I bevat een e-mailadres ÉN kolom J een kwartaalnummer (1–4)
      // Persoonsrijen hebben ook een e-mail in kolom I (nieuw), maar kolom J is daar leeg.
      var notifEmail = (colIWaarde||'').toString().trim();
      var meldKw     = parseInt(sheet.getRange(e.range.getRow(), 10).getValue()) || 0;
      if (notifEmail.indexOf('@') > -1 && meldKw > 0) {
        var meldJaar = parseInt(sheet.getRange('B2').getValue()) || getEffectiveDate().getFullYear();
        clearMeldingViaStatus_(notifEmail, nieuweWaarde, meldKw, meldJaar);
        return;
      }
      // Gewone statusrij (ook persoonsrijen met e-mail in kolom I maar zonder kwartaal in kolom J)
      e.range.setBackground(statusKleur_(nieuweWaarde));
    }
  } catch(err) { Logger.log('onEditJaar fout: ' + err); }
}


// ─── WEKELIJKSE BACKUP ───────────────────────────────────────────────────────

function maakBackupNu() {
  var bestandsnaam = 'Backup Woon-werk ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  DriveApp.getFileById(CONFIG.SPREADSHEET_ID).makeCopy(bestandsnaam, DriveApp.getFolderById(CONFIG.BACKUP_MAP_ID));
  Logger.log('✅ Backup aangemaakt: ' + bestandsnaam);
  try { SpreadsheetApp.getUi().alert('✅ Backup aangemaakt: ' + bestandsnaam); } catch(_) {}
}

function wekelijkseBackup() {
  var bestandsnaam = 'Backup Woon-werk ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  DriveApp.getFileById(CONFIG.SPREADSHEET_ID).makeCopy(bestandsnaam, DriveApp.getFolderById(CONFIG.BACKUP_MAP_ID));
  Logger.log('✅ Wekelijkse backup aangemaakt: ' + bestandsnaam);
}

function installeerWekelijkseBackup() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'wekelijkseBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('wekelijkseBackup').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  Logger.log('✅ Wekelijkse backup trigger geïnstalleerd.');
  try { SpreadsheetApp.getUi().alert('✅ Wekelijkse backup geïnstalleerd.\n\nElke maandag om 6u wordt automatisch een kopie bewaard in de bewijzenmap.'); } catch(_) {}
}
