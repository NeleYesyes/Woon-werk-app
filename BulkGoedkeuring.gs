// ─── BULK GOEDKEURING ────────────────────────────────────────────────────────

function bulkGoedkeurenVoorTabblad_(sheet, goedkKol) {
  var ss             = getSS_();
  var personenMap    = bouwPersonenMap_(ss);
  var emailDomeinMap = bouwEmailDomeinMap_(ss);
  var actieveKw      = bepaalActieveKwartalen_();

  if (sheet.getLastRow() < 2) return 0;
  var aantalRijen = sheet.getLastRow() - 1;
  var idData      = sheet.getRange(2, 1, aantalRijen, 1).getValues();
  var emailData   = sheet.getRange(2, 2, aantalRijen, 1).getValues();
  var naamData    = sheet.getRange(2, 3, aantalRijen, 1).getValues();
  var jaarData    = sheet.getRange(2, 4, aantalRijen, 1).getValues();
  var kwData      = sheet.getRange(2, 5, aantalRijen, 1).getValues();
  var goedkData   = sheet.getRange(2, goedkKol, aantalRijen, 1).getValues();

  var statusValidatie = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ingediend', 'Ja, goedgekeurd', 'Nee, afgekeurd'], true).setAllowInvalid(false).build();

  var goedTeKeurenIds = {};
  var teller = 0;

  // Eerste doorloop: hoofdrijen (ID begint niet met '_')
  for (var r = 0; r < aantalRijen; r++) {
    var idStr = (idData[r][0]||'').toString().trim();
    if (idStr.charAt(0) === '_') continue;
    var rijJaar = parseInt(jaarData[r][0]);
    var rijKw   = parseInt(kwData[r][0]);
    if (!actieveKw.some(function(k) { return k.jaar === rijJaar && k.kwartaal === rijKw; })) continue;
    if ((goedkData[r][0]||'').toString().trim() !== 'Ingediend') continue;
    var naamStr  = (naamData[r][0]||'').toString().trim();
    if (!naamStr) continue;
    var emailStr = (emailData[r][0]||'').toString().toLowerCase().trim();
    var domein   = ((personenMap[naamStr]||{}).domein||'').trim();
    if (!domein && emailStr) domein = (emailDomeinMap[emailStr]||'').trim();
    if (isGroeneCheckKol_(naamStr, domein)) continue;
    sheet.getRange(r + 2, goedkKol)
      .setValue('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
      .setDataValidation(statusValidatie);
    goedTeKeurenIds[idStr] = true;
    teller++;
  }

  // Tweede doorloop: vervolgrijen (woonwerk multi-bewijs) van goedgekeurde hoofdrijen
  for (var r2 = 0; r2 < aantalRijen; r2++) {
    var idStr2 = (idData[r2][0]||'').toString().trim();
    if (idStr2.charAt(0) !== '_') continue;
    if ((goedkData[r2][0]||'').toString().trim() !== 'Ingediend') continue;
    var m = idStr2.match(/^_(.+)_\d+$/);
    if (!m || !goedTeKeurenIds[m[1]]) continue;
    sheet.getRange(r2 + 2, goedkKol)
      .setValue('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
      .setDataValidation(statusValidatie);
    teller++;
  }

  return teller;
}

function bulkGoedkeurenBeeldVoorTabblad_(sheet, beeldGoedkKol) {
  var ss             = getSS_();
  var personenMap    = bouwPersonenMap_(ss);
  var emailDomeinMap = bouwEmailDomeinMap_(ss);
  var actieveKw      = bepaalActieveKwartalen_();

  if (sheet.getLastRow() < 2) return 0;
  var aantalRijen = sheet.getLastRow() - 1;
  var idData      = sheet.getRange(2, 1, aantalRijen, 1).getValues();
  var emailData   = sheet.getRange(2, 2, aantalRijen, 1).getValues();
  var naamData    = sheet.getRange(2, 3, aantalRijen, 1).getValues();
  var jaarData    = sheet.getRange(2, 4, aantalRijen, 1).getValues();
  var kwData      = sheet.getRange(2, 5, aantalRijen, 1).getValues();
  var goedkData   = sheet.getRange(2, beeldGoedkKol, aantalRijen, 1).getValues();

  var statusValidatie = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ingediend', 'Ja, goedgekeurd', 'Nee, afgekeurd'], true).setAllowInvalid(false).build();

  var goedTeKeurenIds = {};
  var teller = 0;

  for (var r = 0; r < aantalRijen; r++) {
    var idStr = (idData[r][0]||'').toString().trim();
    if (idStr.charAt(0) === '_') continue;
    var rijJaarB = parseInt(jaarData[r][0]);
    var rijKwB   = parseInt(kwData[r][0]);
    if (!actieveKw.some(function(k) { return k.jaar === rijJaarB && k.kwartaal === rijKwB; })) continue;
    if ((goedkData[r][0]||'').toString().trim() !== 'Ingediend') continue;
    var naamStr  = (naamData[r][0]||'').toString().trim();
    if (!naamStr) continue;
    var emailStr = (emailData[r][0]||'').toString().toLowerCase().trim();
    var domein   = ((personenMap[naamStr]||{}).domein||'').trim();
    if (!domein && emailStr) domein = (emailDomeinMap[emailStr]||'').trim();
    if (!isGroeneCheckKol_(naamStr, domein)) continue; // alleen Beeld + Inse
    sheet.getRange(r + 2, beeldGoedkKol)
      .setValue('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
      .setDataValidation(statusValidatie);
    goedTeKeurenIds[idStr] = true;
    teller++;
  }

  // Vervolgrijen van goedgekeurde hoofdrijen ook goedkeuren
  for (var r2 = 0; r2 < aantalRijen; r2++) {
    var idStr2 = (idData[r2][0]||'').toString().trim();
    if (idStr2.charAt(0) !== '_') continue;
    if ((goedkData[r2][0]||'').toString().trim() !== 'Ingediend') continue;
    var m = idStr2.match(/^_(.+)_\d+$/);
    if (!m || !goedTeKeurenIds[m[1]]) continue;
    sheet.getRange(r2 + 2, beeldGoedkKol)
      .setValue('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
      .setDataValidation(statusValidatie);
    teller++;
  }

  return teller;
}

function isGroeneCheckKol_(nm, dm) {
  return dm === 'Beeld' || nm === 'Verplanken Inse';
}

function checkKolVoorSheet_(sheetNaam) {
  if (sheetNaam === CONFIG.SHEETS.FIETSVERGOEDING)    return 12;  // L
  if (sheetNaam === CONFIG.SHEETS.WOON_WERK)          return 14;  // N
  if (sheetNaam === CONFIG.SHEETS.DIENSTVERPLAATSING) return 14;  // N
  return null;
}

function kleurBeeldCheckKol_(ss, sheet, checkKol) {
  if (sheet.getLastRow() < 2) return;
  var personenMap    = bouwPersonenMap_(ss);
  var emailDomeinMap = bouwEmailDomeinMap_(ss);
  var aantalRijen = sheet.getLastRow() - 1;
  var nicoKolMap  = { [CONFIG.SHEETS.FIETSVERGOEDING]: 11, [CONFIG.SHEETS.WOON_WERK]: 13, [CONFIG.SHEETS.DIENSTVERPLAATSING]: 13 };
  var beeldKolMap = { [CONFIG.SHEETS.FIETSVERGOEDING]: 13, [CONFIG.SHEETS.WOON_WERK]: 15, [CONFIG.SHEETS.DIENSTVERPLAATSING]: 15 };
  var nicoKol  = nicoKolMap[sheet.getName()];
  var beeldKol = beeldKolMap[sheet.getName()];
  var emailData = sheet.getRange(2, 2, aantalRijen, 1).getValues();
  var naamData  = sheet.getRange(2, 3, aantalRijen, 1).getValues();
  var nicoData  = nicoKol  ? sheet.getRange(2, nicoKol,  aantalRijen, 1).getValues() : null;
  var beeldData = beeldKol ? sheet.getRange(2, beeldKol, aantalRijen, 1).getValues() : null;
  var bgs = [], fclrs = [], afgekeurdLijst = [];
  for (var i = 0; i < aantalRijen; i++) {
    var nicoVal  = nicoData  ? (nicoData[i][0] ||'').toString().trim() : '';
    var beeldVal = beeldData ? (beeldData[i][0]||'').toString().trim() : '';
    afgekeurdLijst.push(nicoVal === 'Nee, afgekeurd' || beeldVal === 'Nee, afgekeurd');
    var nm       = (naamData[i][0]||'').toString().trim();
    var emailStr = (emailData[i][0]||'').toString().toLowerCase().trim();
    var dm       = ((personenMap[nm]||{}).domein||'').trim();
    if (!dm && emailStr) dm = (emailDomeinMap[emailStr]||'').trim();
    bgs.push([isGroeneCheckKol_(nm, dm) ? '#c9f080' : null]); fclrs.push([null]);
  }
  var rng = sheet.getRange(2, checkKol, aantalRijen, 1);
  rng.setBackgrounds(bgs);
  rng.setFontColors(fclrs);
  // Vink afgekeurde rijen aan
  for (var vi = 0; vi < aantalRijen; vi++) {
    if (afgekeurdLijst[vi]) sheet.getRange(2 + vi, checkKol).setValue(true);
  }
}
