// ─── SORTERING PER PERSONEELSLID ─────────────────────────────────────────────

function sorteerVerplaatsingenSheets() {
  var ss      = getSS_();
  var persMap = bouwEmailNaamMap_(ss);
  sorteerPersoneelsSheet_(ss);
  [CONFIG.SHEETS.FIETSVERGOEDING, CONFIG.SHEETS.WOON_WERK, CONFIG.SHEETS.DIENSTVERPLAATSING]
    .forEach(function(naam) {
      var sheet = ss.getSheetByName(naam);
      if (sheet && sheet.getLastRow() >= 2) sorteerSheet_(sheet, persMap);
    });
  Logger.log('✅ Tabbladen gesorteerd op personeelslid.');
  try { SpreadsheetApp.getUi().alert('Klaar! Alle tabbladen zijn gesorteerd op personeelslid (alfabetisch op naam).'); } catch(_) {}
}

function stelGoedkeuringKolommenIn() {
  var ss             = getSS_();
  var personenMap    = bouwPersonenMap_(ss);
  var emailDomeinMap = bouwEmailDomeinMap_(ss);
  var instellingen = [
    { naam: CONFIG.SHEETS.FIETSVERGOEDING,    goedkKol: 11, checkKol: 12, beeldKol: 13 },
    { naam: CONFIG.SHEETS.WOON_WERK,          goedkKol: 13, checkKol: 14, beeldKol: 15 },
    { naam: CONFIG.SHEETS.DIENSTVERPLAATSING, goedkKol: 13, checkKol: 14, beeldKol: 15 },
  ];
  var statusValidatie = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ingediend', 'Ja, goedgekeurd', 'Nee, afgekeurd'], true).setAllowInvalid(false).build();

  instellingen.forEach(function(inst) {
    var sheet = ss.getSheetByName(inst.naam);
    if (!sheet) return;
    var nicoKolLetter  = kolomLetterVan_(inst.goedkKol);
    var beeldKolLetter = kolomLetterVan_(inst.beeldKol);

    // Rij 1: Nico bulk-checkbox (niet-Beeld)
    var bulkCelNico = sheet.getRange(1, inst.goedkKol);
    bulkCelNico.clearContent().clearDataValidations();
    bulkCelNico.insertCheckboxes().setValue(false);
    bulkCelNico.setBackground('#26295a').setFontColor('#ffffff').setHorizontalAlignment('center');
    bulkCelNico.setNote('Klik aan om alle "Ingediend"-rijen van het huidig kwartaal goed te keuren (Muziek / Woord / Kunstkuur / Beleid — Beeld-medewerkers worden overgeslagen).');
    sheet.setColumnWidth(inst.goedkKol, 130);

    // Rij 1: controlekolom-header
    sheet.getRange(1, inst.checkKol)
      .setValue('Gecontroleerd?').setFontWeight('bold')
      .setBackground('#e2e8f0').setFontColor('#475569').setHorizontalAlignment('center');
    sheet.setColumnWidth(inst.checkKol, 105);

    // Rij 1: Inse/Beeld bulk-checkbox
    var bulkCelBeeld = sheet.getRange(1, inst.beeldKol);
    bulkCelBeeld.clearContent().clearDataValidations();
    bulkCelBeeld.insertCheckboxes().setValue(false);
    bulkCelBeeld.setBackground('#26295a').setFontColor('#ffffff').setHorizontalAlignment('center');
    bulkCelBeeld.setNote('Klik aan om alle "Ingediend"-rijen van het huidig kwartaal goed te keuren (uitsluitend Beeld-medewerkers en Verplanken Inse).');
    sheet.setColumnWidth(inst.beeldKol, 130);

    if (sheet.getLastRow() < 2) return;

    var lastRow     = sheet.getLastRow();
    var aantalRijen = lastRow - 1;
    var emailData   = sheet.getRange(2, 2, aantalRijen, 1).getValues();
    var naamData    = sheet.getRange(2, 3, aantalRijen, 1).getValues();
    var goedkData   = sheet.getRange(2, inst.goedkKol, aantalRijen, 1).getValues();
    var beeldData   = sheet.getRange(2, inst.beeldKol, aantalRijen, 1).getValues();
    var checkData   = sheet.getRange(2, inst.checkKol, aantalRijen, 1).getValues();

    // Wis alle drie kolommen — witte achtergrond op goedkeuringskolommen zodat hyperlinks niet doorbloeden
    sheet.getRange(2, inst.goedkKol, aantalRijen, 1).clearDataValidations().clearContent().setBackground('#ffffff').setFontColor(null);
    sheet.getRange(2, inst.beeldKol, aantalRijen, 1).clearDataValidations().clearContent().setBackground('#ffffff').setFontColor(null);
    sheet.getRange(2, inst.checkKol, aantalRijen, 1).clearDataValidations().clearContent().setBackground(null).setFontColor(null);

    for (var r = 0; r < emailData.length; r++) {
      if (!(emailData[r][0]||'').toString().trim()) continue;
      var rijnr    = r + 2;
      var nm       = (naamData[r][0]||'').toString().trim();
      var emailStr = (emailData[r][0]||'').toString().toLowerCase().trim();
      var dm       = ((personenMap[nm]||{}).domein||'').trim();
      if (!dm && emailStr) dm = (emailDomeinMap[emailStr]||'').trim();
      var isBeeld  = isGroeneCheckKol_(nm, dm);

      // Goedkeuringswaarde: lees uit de juiste bronkolom, migreer 'Ja'/'Nee'
      var bronWaarde = isBeeld ? (beeldData[r][0]||'').toString().trim()
                               : (goedkData[r][0]||'').toString().trim();
      // Bij migratie: als beeldKol leeg is maar goedkKol gevuld, neem die mee
      if (isBeeld && !bronWaarde) bronWaarde = (goedkData[r][0]||'').toString().trim();
      if (bronWaarde === 'Ja')        bronWaarde = 'Ja, goedgekeurd';
      else if (bronWaarde === 'Nee') bronWaarde = 'Nee, afgekeurd';
      else if (!bronWaarde)           bronWaarde = 'Ingediend';

      var kleur    = goedkeuringStatusKleur_(bronWaarde);
      var doelKol  = isBeeld ? inst.beeldKol : inst.goedkKol;
      var leegKol  = isBeeld ? inst.goedkKol : inst.beeldKol;
      var doelCelSetup = sheet.getRange(rijnr, doelKol).setValue(bronWaarde).setBackground(kleur.bg).setFontColor(kleur.txt);
      // 'Beeld' is een overgebleven plaatshouder uit de andere kolom, geen echte status: geen keuzelijst erop
      if (bronWaarde === 'Beeld') doelCelSetup.clearDataValidations();
      else doelCelSetup.setDataValidation(statusValidatie);
      // Nico-kolom voor Beeld-rijen: 'Beeld' in witte letter zodat hyperlink niet visueel doorloopt
      var leegCelSetup = sheet.getRange(rijnr, leegKol);
      leegCelSetup.clearDataValidations().setBackground('#ffffff');
      if (isBeeld) {
        leegCelSetup.setValue('Beeld').setFontColor('#ffffff');
      } else {
        leegCelSetup.clearContent().setFontColor(null);
      }

      // Controle-checkbox (bestaand vinkje bewaren)
      var checkCel = sheet.getRange(rijnr, inst.checkKol);
      checkCel.insertCheckboxes();
      if (checkData[r][0] === true) checkCel.setValue(true);
      if (bronWaarde === 'Nee, afgekeurd') checkCel.setValue(true);
      checkCel.setBackground(isBeeld ? '#c9f080' : null).setFontColor(null);
    }

    // Voorwaardelijke opmaak voor beide goedkeuringskolommen
    var nicoGoedkBereik  = sheet.getRange(2, inst.goedkKol,  500, 1);
    var beeldGoedkBereik = sheet.getRange(2, inst.beeldKol, 500, 1);
    var geheelBereik     = sheet.getRange(2, 1, 500, inst.beeldKol - 1);
    sheet.setConditionalFormatRules([
      // Hele rij rood in lettertype bij afgekeurd in Nico-kolom
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + nicoKolLetter + '2="Nee, afgekeurd"')
        .setFontColor('#dc2626')
        .setRanges([sheet.getRange(2, 1, 500, inst.goedkKol - 1)])
        .build(),
      // Hele rij rood in lettertype bij afgekeurd in Beeld-kolom (range stopt vóór nicoGoedkKol zodat 'Beeld'-tekst wit blijft)
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + beeldKolLetter + '2="Nee, afgekeurd"')
        .setFontColor('#dc2626')
        .setRanges([sheet.getRange(2, 1, 500, inst.goedkKol - 1)])
        .build(),
      // Celkleur Nico-goedkeuringskolom
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
        .setRanges([nicoGoedkBereik]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Nee, afgekeurd').setBackground('#fee2e2').setFontColor('#1e293b')
        .setRanges([nicoGoedkBereik]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Ingediend').setBackground('#f1f5f9').setFontColor('#64748b')
        .setRanges([nicoGoedkBereik]).build(),
      // Celkleur Beeld-goedkeuringskolom
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Ja, goedgekeurd').setBackground('#bbf7d0').setFontColor('#14532d')
        .setRanges([beeldGoedkBereik]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Nee, afgekeurd').setBackground('#fee2e2').setFontColor('#1e293b')
        .setRanges([beeldGoedkBereik]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Ingediend').setBackground('#f1f5f9').setFontColor('#64748b')
        .setRanges([beeldGoedkBereik]).build(),
    ]);
  });
  // Herstel getalopmaak voor parkeerBedrag (kolom J) in Dienstverplaatsing
  var dienstSheet = ss.getSheetByName(CONFIG.SHEETS.DIENSTVERPLAATSING);
  if (dienstSheet && dienstSheet.getLastRow() >= 2) {
    dienstSheet.getRange(2, 10, dienstSheet.getLastRow() - 1, 1).setNumberFormat('0.00');
  }
  Logger.log('✅ Goedkeuring- en controlekolommen ingesteld op de drie tabbladen.');
}

function kolomLetterVan_(n) {
  var s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
