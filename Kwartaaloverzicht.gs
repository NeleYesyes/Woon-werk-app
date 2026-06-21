// ─── KWARTAALOVERZICHT ────────────────────────────────────────────────────────

function getTarieven_(jaar) {
  var ss    = getSS_();
  var sheet = ss.getSheetByName('Tarieven');
  var fiets = 0.35, dienst = 0.42;
  if (!sheet || sheet.getLastRow() < 2) return { fiets: fiets, dienst: dienst };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(jaar)) {
      fiets  = parseFloat(data[i][1]) || fiets;
      dienst = parseFloat(data[i][2]) || dienst;
      break;
    }
  }
  return { fiets: fiets, dienst: dienst };
}

function leesStatussenUitBestaandeSheet_(sheet) {
  if (!sheet) return {};
  var result = {};
  if (sheet.getLastRow() < 4) return result;
  var jaarVal = parseInt(sheet.getRange('B2').getValue());
  var jaar    = isNaN(jaarVal) ? new Date().getFullYear() : jaarVal;
  var data    = sheet.getDataRange().getValues();
  var huidigKw = 0, huidigeSectie = '';
  for (var i = 0; i < data.length; i++) {
    var col0 = (data[i][0]||'').toString();
    for (var q = 1; q <= 4; q++) { if (col0.indexOf('Kwartaal ' + q) > -1) { huidigKw = q; break; } }
    if (col0.indexOf('Fietsvergoeding')    > -1) huidigeSectie = 'fiets';
    if (col0.indexOf('Woon-Werk')          > -1) huidigeSectie = 'woonwerk';
    if (col0.indexOf('Dienstverplaatsing') > -1) huidigeSectie = 'dienst';
    var status = (data[i][10]||'').toString().trim();
    if ((status === STATUS_INGEDIEND || status === STATUS_CONTROLE || status === STATUS_BETALING) && huidigKw > 0 && huidigeSectie) {
      var email  = (data[i][8]||'').toString().trim().toLowerCase(); // col 9 (index 8): e-mail (nieuw)
      var naam   = (data[i][1]||'').toString().trim();               // col B (index 1): persoonsnaam
      var prefix = jaar + '|' + huidigKw + '|' + huidigeSectie + '|';
      // Primaire sleutel: e-mail of persoonsnaam (nooit domein)
      var primSleutel = prefix + (email || naam);
      if (email || naam) result[primSleutel] = status;
      // Overgangsbruggetje: bewaar ook de domein-sleutel als col 9 nog leeg was (oud overzicht)
      if (!email) {
        var domein = col0.replace(/^Totaal\s+/,'').trim();
        if (domein) result[prefix + domein] = status;
      }
    }
  }
  return result;
}

// Overzicht per persoon voor e-mailrapport
function getQuarterlyOverview(quarter, year) {
  var ss       = getSS_();
  var fietsR   = leesRittenVoorJaar_(ss, CONFIG.SHEETS.FIETSVERGOEDING,    year, 3);
  var woonwR   = leesRittenVoorJaar_(ss, CONFIG.SHEETS.WOON_WERK,          year, 3);
  var dienstR  = leesRittenVoorJaar_(ss, CONFIG.SHEETS.DIENSTVERPLAATSING, year, 3);
  var tarieven = getTarieven_(year);

  var map = {};
  fietsR.forEach(function(r) {
    if (parseInt(r[4]) !== quarter) return;
    var naam = (r[2]||'').toString().trim();
    if (!map[naam]) map[naam] = { name: naam, bikeTotal: 0, transportTotal: 0, serviceTotal: 0 };
    map[naam].bikeTotal += parseFloat(r[7]||0) * tarieven.fiets;
  });
  woonwR.forEach(function(r) {
    if (parseInt(r[4]) !== quarter) return;
    var naam = (r[2]||'').toString().trim();
    if (!map[naam]) map[naam] = { name: naam, bikeTotal: 0, transportTotal: 0, serviceTotal: 0 };
    map[naam].transportTotal += 0;  // bedrag niet meer opgeslagen in het tabblad
  });
  dienstR.forEach(function(r) {
    if (parseInt(r[4]) !== quarter) return;
    var naam = (r[2]||'').toString().trim();
    if (!map[naam]) map[naam] = { name: naam, bikeTotal: 0, transportTotal: 0, serviceTotal: 0 };
    map[naam].serviceTotal += parseFloat(r[8]||0) * tarieven.dienst;
  });
  return Object.keys(map).sort().map(function(n) {
    var m = map[n];
    m.total = m.bikeTotal + m.transportTotal + m.serviceTotal;
    return m;
  });
}

function maakKwartaaloverzicht(terugNaamParam) {
  var terugNaam;
  if (terugNaamParam) {
    terugNaam = terugNaamParam;
  } else {
    try { var _a = SpreadsheetApp.getActiveSheet(); terugNaam = _a ? _a.getName() : null; } catch(_) { terugNaam = null; }
  }
  var ss       = getSS_();
  var naam     = 'Kwartaaloverzicht';
  var nu       = getEffectiveDate();
  var huidigKw = huidigKwartaal_();

  var bestaand    = ss.getSheetByName(naam);
  var jaarCel     = bestaand ? bestaand.getRange('B2').getValue() : null;
  var jaar        = (jaarCel && !isNaN(parseInt(jaarCel))) ? parseInt(jaarCel) : nu.getFullYear();
  var isHuidigJaar = (jaar === nu.getFullYear());

  var FIETS_TARIEF  = 0;
  var DIENST_TARIEF = 0;
  var vorigeOpenKw  = 0;
  var hData;
  if (bestaand) {
    hData = bestaand.getDataRange().getValues();
    for (var hi = 0; hi < hData.length - 1; hi++) {
      var hLabel = (hData[hi][0]||'').toString();
      // Het tarief staat in de headerrij (hi+1), niet in de titelrij (hi) zelf
      if (hLabel.indexOf('Fietsvergoeding') > -1) {
        var fVal = parseFloat(hData[hi+1][7]);
        if (!isNaN(fVal) && fVal > 0) FIETS_TARIEF = fVal;
      }
      if (hLabel.indexOf('Dienstverplaatsing') > -1) {
        var dVal = parseFloat(hData[hi+1][7]);
        if (!isNaN(dVal) && dVal > 0) DIENST_TARIEF = dVal;
      }
    }
  }
  // Bepaal welk kwartaal open moet staan:
  // 1. PropertiesService: door gebruiker expliciet gekozen via H-checkbox (meest betrouwbaar)
  // 2. H-checkbox in oud blad (enkel voor vorige jaren als fallback)
  // 3. Standaard: huidig kwartaal (huidig jaar) of K1 (vorig jaar)
  var props = PropertiesService.getScriptProperties();
  var storedKw = parseInt(props.getProperty('lastOpenKw_' + jaar)) || 0;
  if (storedKw >= 1 && storedKw <= 4) {
    vorigeOpenKw = storedKw;
  } else if (!isHuidigJaar && hData) {
    for (var vi = 3; vi < hData.length; vi++) {
      if (hData[vi][7] === true && parseInt(hData[vi][8]) > 0) {
        var kwLabelVI = (hData[vi][0]||'').toString();
        var kwMVI = kwLabelVI.match(/Kwartaal\s+(\d)/);
        if (kwMVI) { vorigeOpenKw = parseInt(kwMVI[1]); break; }
      }
    }
  }
  if (!vorigeOpenKw) vorigeOpenKw = isHuidigJaar ? huidigKw : 1;
  // Altijd persisteren: ook als het kwartaal via H-checkbox of standaard bepaald werd,
  // zodat opeenvolgende rebuilds (bv. na plakken) steeds hetzelfde kwartaal heropenen.
  props.setProperty('lastOpenKw_' + jaar, vorigeOpenKw.toString());

  var statusLookup = leesStatussenUitBestaandeSheet_(bestaand);

  // Lees vergrendelingstatus per kwartaal uit PropertiesService (blijft bewaard bij herbouw)
  var vergrendelingLookup = {};
  for (var vk = 1; vk <= 4; vk++) {
    if (props.getProperty('lock_' + jaar + '_' + vk) === 'true') vergrendelingLookup[vk] = true;
  }
  var nicoTarief = parseFloat(props.getProperty('tariefFiets_Logghe Nico')) || 0;
  var propsF = parseFloat(props.getProperty('tariefFiets'));
  var propsD = parseFloat(props.getProperty('tariefDienst'));
  if (!isNaN(propsF) && propsF > 0) FIETS_TARIEF  = propsF;
  if (!isNaN(propsD) && propsD > 0) DIENST_TARIEF = propsD;

  var positie = bestaand ? bestaand.getIndex() - 1 : 0;
  var tempNaam = naam + '_TEMP';
  var oudTemp = ss.getSheetByName(tempNaam);
  if (oudTemp) { try { ss.deleteSheet(oudTemp); } catch(_) {} }
  var sheet = ss.insertSheet(tempNaam, positie);
  sheet.setTabColor('#26295a');

  // Rij 1: Titel
  sheet.getRange('A1:G1').merge()
    .setValue('Kwartaaloverzicht — Academie Ieper')
    .setBackground('#26295a').setFontColor('#ffffff')
    .setFontSize(15).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 46);

  // Rij 2: Jaar
  sheet.getRange('A2').setValue('Jaar:').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('B2').setValue(jaar).setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground('#ede9fe').setFontColor('#5b21b6');
  var bijgewerktOp = Utilities.formatDate(getEffectiveDate(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  sheet.getRange('C2:G2').merge()
    .setValue('← Pas het jaar aan in cel B2 — het overzicht ververst automatisch   ·   Bijgewerkt op ' + bijgewerktOp)
    .setFontSize(9).setFontColor('#94a3b8').setVerticalAlignment('middle');
  sheet.getRange('H2').setValue('Tarief ' + jaar).setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#ede9fe').setFontColor('#5b21b6');
  sheet.getRange(2, 11).setValue('Stad Ieper').setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#ede9fe').setFontColor('#5b21b6');
  sheet.setRowHeight(2, 30);

  var jaren = [];
  for (var y = jaar - 4; y <= jaar + 2; y++) jaren.push(y);
  sheet.getRange('B2').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(jaren.map(String), true).setAllowInvalid(false).build()
  );
  sheet.setRowHeight(3, 10);

  var fietsR   = leesRittenVoorJaar_(ss, CONFIG.SHEETS.FIETSVERGOEDING,    jaar, 3);
  var woonwR   = leesRittenVoorJaar_(ss, CONFIG.SHEETS.WOON_WERK,          jaar, 3);
  var dienstR  = leesRittenVoorJaar_(ss, CONFIG.SHEETS.DIENSTVERPLAATSING, jaar, 3);
  var persSheetShared = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  var persDataGedeeld = (persSheetShared && persSheetShared.getLastRow() >= 2) ? persSheetShared.getDataRange().getValues() : [[]];
  var personenMap = bouwPersonenMap_(ss, persDataGedeeld);
  var personenMapByEmail = bouwPersonenMapByEmail_(ss, persDataGedeeld);

  // Verberg/wis oude personeelsrijen uit vorige jaren
  ruimPersoneelsOudRijenOp_(ss, jaar, persDataGedeeld.length > 1 ? persDataGedeeld.slice(1) : []);

  // Personeelsmeldingen per kwartaal — enkel voor het weergegeven jaar
  var meldingPerKw = { 1:[], 2:[], 3:[], 4:[] };
  if (persDataGedeeld.length >= 2) {
    for (var mi = 1; mi < persDataGedeeld.length; mi++) {
      var mStatus = (persDataGedeeld[mi][10]||'').toString().trim().toLowerCase();
      if (mStatus !== 'nieuw' && mStatus !== 'gewijzigd') continue;
      var geldigheid = (persDataGedeeld[mi][8]||'').toString().trim();
      var meldDetail = (persDataGedeeld[mi][9]||'').toString().trim();
      var mVanaf = geldigheid.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (!mVanaf) mVanaf = meldDetail.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (!mVanaf || parseInt(mVanaf[3]) !== jaar) continue;
      var meldDatum = new Date(parseInt(mVanaf[3]), parseInt(mVanaf[2])-1, parseInt(mVanaf[1]));
      var mKw = getQuarterFromDate(meldDatum);
      var melEmail   = (persDataGedeeld[mi][7]||persDataGedeeld[mi][5]||'').toString().toLowerCase().trim();
      var normEmail  = melEmail.replace(/[^a-z0-9@._-]/g, '_');
      var handledKw  = parseInt(props.getProperty('meldingHandledKw_' + normEmail + '_' + jaar)) || 0;
      var melObj = { domein: (persDataGedeeld[mi][0]||'').toString().trim(), naam: (persDataGedeeld[mi][1]||'').toString().trim(), email: melEmail, type: mStatus, handledKw: handledKw };
      for (var mq = mKw; mq <= 4; mq++) {
        if (handledKw > 0 && mq > handledKw) continue;
        meldingPerKw[mq].push(melObj);
      }
    }
  }

  var rij = 4;

  var kwNamen = ['', 'Kwartaal 1  —  jan / feb / mrt', 'Kwartaal 2  —  apr / mei / jun', 'Kwartaal 3  —  jul / aug / sep', 'Kwartaal 4  —  okt / nov / dec'];
  var kwartaalRijen = [];

  for (var kw = 1; kw <= 4; kw++) {
    if (kw > 1) { sheet.getRange(rij, 1, 1, 11).setBackground('#ffffff'); sheet.setRowHeight(rij, 18); rij++; }
    var isHuidig   = isHuidigJaar && (kw === huidigKw);
    var isGeopend  = (kw === vorigeOpenKw);
    var isGesloten = isHuidigJaar && isLocked(kw, jaar);
    var kwLabel    = kwNamen[kw];
    if (isHuidig)   kwLabel += '   ★ huidig kwartaal';
    if (isGesloten) kwLabel += '   🔒 Afgesloten op ' + formatDatumWeergave_(getQuarterDeadline(kw, jaar));

    var headerRij = rij;
    var kwHeader  = sheet.getRange(rij, 1, 1, 7);
    kwHeader.merge()
      .setValue(kwLabel)
      .setBackground(isGeopend ? '#f59e0b' : isGesloten ? '#64748b' : '#ffffff')
      .setFontColor(isGeopend ? '#7c2d12' : isGesloten ? '#ffffff' : '#94a3b8')
      .setFontSize(isGeopend ? 11 : 10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    kwHeader.setBorder(true, null, null, null, null, null, isGeopend ? '#ea580c' : '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(headerRij, 8)
      .setValue(isGeopend)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())
      .setBackground(isGeopend ? '#ede9fe' : isGesloten ? '#f1f5f9' : '#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Kolom K: checkbox Stad Ieper
    var kwVergrendeld = vergrendelingLookup[kw] || false;
    sheet.getRange(headerRij, 11)
      .setValue(kwVergrendeld)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())
      .setBackground(kwVergrendeld ? '#bbf7d0' : '#dbeafe')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Kolom L: instructietekst naast Stad Ieper checkbox — voor het initieel geopende kwartaal
    // Huidig jaar → huidig kwartaal; ander jaar → kwartaal 1 (dat wordt standaard geopend)
    var isInitieelGeopend = isGeopend;
    if (isInitieelGeopend) {
      sheet.getRange(headerRij, 12)
        .setValue('← Helemaal klaar met kwartaal? Klik checkbox, graag!')
        .setFontSize(9).setFontStyle('italic').setFontColor('#1e40af')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    } else {
      sheet.getRange(headerRij, 12).clearContent();
    }
    sheet.setRowHeight(rij, isGeopend ? 42 : 26); rij++;

    var kwDataStart = rij;

    // Meldingen die horen bij dit kwartaal
    var kwMeldingen = meldingPerKw[kw] || [];
    if (kwMeldingen.length > 0) {
      var meldValidatie = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Opgelet: wijziging', 'Wijziging doorgevoerd'], true).setAllowInvalid(false).build();
      for (var mi2 = 0; mi2 < kwMeldingen.length; mi2++) {
        var mel        = kwMeldingen[mi2];
        var isAfgedaan = mel.handledKw > 0 && kw === mel.handledKw;
        var isNieuw    = mel.type === 'nieuw';
        var bgKleur    = isAfgedaan ? '#bbf7d0' : (isNieuw ? '#fef3c7' : '#fee2e2');
        var tekst      = isNieuw ? 'Nieuwe medewerker — gegevens nog te verwerken door Stad Ieper' : 'Gewijzigde gegevens — aanpassing vereist door Stad Ieper';
        sheet.getRange(rij, 1).setValue(mel.domein).setFontSize(9).setFontWeight('normal').setFontColor('#94a3b8').setBackground(bgKleur);
        sheet.getRange(rij, 2).setValue(mel.naam).setFontSize(9).setFontWeight('bold').setFontColor('#334155').setBackground(bgKleur);
        sheet.getRange(rij, 3, 1, 5).merge().setValue(tekst).setFontSize(9).setFontColor('#475569').setBackground(bgKleur);
        sheet.getRange(rij, 9).setValue(mel.email);
        sheet.getRange(rij, 10).setValue(kw);
        var meldStatus = isAfgedaan ? 'Wijziging doorgevoerd' : 'Opgelet: wijziging';
        sheet.getRange(rij, 11).setValue(meldStatus).setDataValidation(meldValidatie)
          .setBackground(isAfgedaan ? '#bbf7d0' : '#dc2626')
          .setFontColor(isAfgedaan ? '#14532d' : '#ffffff')
          .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setRowHeight(rij, 24); rij++;
      }
      sheet.setRowHeight(rij, 6); rij++;
    }

    rij = schrijfCategorie_(sheet, rij,
      '🚴  Fietsvergoeding', '#26295a', '#ffffff',
      aggregeerFiets_(fietsR, kw, FIETS_TARIEF, personenMap, personenMapByEmail),
      ['Domein', 'Naam leerkracht', 'Rijksregisternummer', 'Bestemming', 'KM', 'Vergoeding', 'IBAN'],
      true, FIETS_TARIEF, 'fiets', jaar, kw, statusLookup, nicoTarief);

    rij = schrijfCategorie_(sheet, rij,
      '🚆  Woon-Werk (trein-De Lijn)', '#26295a', '#ffffff',
      aggregeerWoonwerk_(woonwR, kw, personenMap, personenMapByEmail),
      ['Domein', 'Naam leerkracht', 'Rijksregisternummer', 'Ritten', '', 'Vergoeding', 'IBAN'],
      true, null, 'woonwerk', jaar, kw, statusLookup);

    rij = schrijfCategorie_(sheet, rij,
      '🚗  Dienstverplaatsing', '#26295a', '#ffffff',
      aggregeerDienst_(dienstR, kw, DIENST_TARIEF, personenMap, personenMapByEmail),
      ['Domein', 'Naam leerkracht', 'Rijksregisternummer', 'Verplaatsingen', 'KM', 'Vergoeding', 'Parkeerticket (€)'],
      true, DIENST_TARIEF, 'dienst', jaar, kw, statusLookup);

    sheet.setRowHeight(rij, 24); rij++;
    kwartaalRijen.push({ kw: kw, dataStart: kwDataStart, dataEinde: rij - 1, isHuidig: isHuidig, isGeopend: isGeopend });
  }

  var defaultBreedtes = { 1:60, 2:100, 3:85, 4:115, 5:60, 6:65, 7:135, 8:80, 11:155, 12:310 };
  for (var ci = 1; ci <= 12; ci++) {
    if (defaultBreedtes[ci]) sheet.setColumnWidth(ci, defaultBreedtes[ci]);
  }
  sheet.setFrozenRows(3);

  var kBereik = sheet.getRange(4, 11, rij - 4, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Wijziging doorgevoerd').setBackground('#bbf7d0').setFontColor('#14532d').setRanges([kBereik]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS_BETALING).setBackground('#bbf7d0').setFontColor('#14532d').setRanges([kBereik]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS_CONTROLE).setBackground('#fef9c3').setFontColor('#713f12').setRanges([kBereik]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS_INGEDIEND).setBackground('#f1f5f9').setFontColor('#475569').setRanges([kBereik]).build(),
  ]);

  kwartaalRijen.forEach(function(k) {
    var hdrRij = k.dataStart - 1;
    sheet.getRange(hdrRij, 9).setValue(k.dataStart);
    sheet.getRange(hdrRij, 10).setValue(k.dataEinde);
    if (isHuidigJaar) {
      if (!k.isGeopend) sheet.hideRows(k.dataStart, k.dataEinde - k.dataStart + 1);
    } else {
      sheet.getRange(hdrRij, 8).setValue(k.isGeopend);
      if (!k.isGeopend) sheet.hideRows(k.dataStart, k.dataEinde - k.dataStart + 1);
    }
  });
  sheet.hideColumns(9, 2);

  // Herstel verberging van rit-rijen voor vergrendelde kwartalen (ook na herbouw)
  for (var hk = 1; hk <= 4; hk++) {
    if (vergrendelingLookup[hk]) {
      try { verbergRitRijenVoorKwartaal_(ss, jaar, hk); } catch(_) {}
    }
  }

  // Atomische swap: hernoem oud tabblad, hernoem nieuw naar definitieve naam, verwijder oud
  if (bestaand) { try { bestaand.setName(naam + '_OUD'); } catch(_) {} }
  sheet.setName(naam);
  if (bestaand) {
    try { var oudSheet = ss.getSheetByName(naam + '_OUD'); if (oudSheet) ss.deleteSheet(oudSheet); } catch(_) {}
  }

  if (terugNaam && terugNaam !== naam) {
    var herstel = ss.getSheetByName(terugNaam);
    try { if (herstel) ss.setActiveSheet(herstel, true); } catch(_) {}
  } else if (terugNaam === naam) {
    try { ss.setActiveSheet(sheet, false); } catch(_) {}
  } else {
    try { SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet); } catch(_) {}
  }
  try { stelKwartaaloverzichtBeveiligingIn(); } catch(_) {}
  Logger.log('✅ Kwartaaloverzicht bijgewerkt voor ' + jaar + '.');
}

function verversKwartaaloverzichtAlsBestaat_(terugNaamParam) {
  try {
    // Eventuele nog geplande refresh annuleren (na directe aanroep niet meer nodig)
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'verversKwartaaloverzichtGetriggerd') ScriptApp.deleteTrigger(t);
    });
    // Direct herbouwen: time-based triggers hebben een minimumvertraging van ~1 minuut,
    // waardoor .after(5000) toch pas na ±1 minuut vuurde. Installable triggers draaien
    // 6 minuten, dus maakKwartaaloverzicht() past hier gewoon in.
    maakKwartaaloverzicht(terugNaamParam);
  } catch(e) { Logger.log('⚠️ Kwartaaloverzicht vernieuwen mislukt: ' + e); }
}

function verversKwartaaloverzichtGetriggerd() {
  try { maakKwartaaloverzicht(); } catch(e) { Logger.log('⚠️ Getriggerde refresh mislukt: ' + e); }
}

function stelKwartaaloverzichtBeveiligingIn() {
  var ss = getSS_();
  var sheet = ss.getSheetByName('Kwartaaloverzicht');
  if (!sheet) {
    try { SpreadsheetApp.getUi().alert('Tabblad Kwartaaloverzicht niet gevonden.'); } catch(_) {}
    return;
  }

  // Wie mag NIET bewerken (maar wel bekijken) — vul aan bij nieuwe directeurs
  var DENYLIST = [
    'myriam.demeester@ieper.be',
    'myriam.demeester@academie-ieper.be',
    'nele.moerman@ieper.be',
    'nele.moerman@academie-ieper.be',
    'nico.logghe@academie-ieper.be',
    'nico.logghe@ieper.be',
    'inse.verplanken@academie-ieper.be',
    'inse.verplanken@ieper.be'
  ].map(function(e) { return e.toLowerCase(); });

  // Verwijder bestaande beveiligingen op dit tabblad
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
  var prot = sheet.protect().setDescription('Kwartaaloverzicht — directeurs kunnen niet bewerken');

  // Allowlist = alle spreadsheet-editors min de denylist
  var editors = ss.getEditors().map(function(u) { return u.getEmail().toLowerCase(); });
  try { var owner = ss.getOwner(); if (owner) editors.push(owner.getEmail().toLowerCase()); } catch(_) {}
  try { editors.push(Session.getEffectiveUser().getEmail().toLowerCase()); } catch(_) {}
  var toegelaten = editors.filter(function(email, idx, arr) {
    return arr.indexOf(email) === idx && DENYLIST.indexOf(email) === -1;
  });

  prot.removeEditors(prot.getEditors());
  prot.addEditors(toegelaten);

  Logger.log('✅ Beveiliging Kwartaaloverzicht — toegelaten: ' + toegelaten.join(', ') + ' | geblokkeerd: ' + DENYLIST.join(', '));
  try {
    SpreadsheetApp.getUi().alert(
      'Beveiliging ingesteld.\n\nKunnen bewerken: ' + toegelaten.join('\n') +
      '\n\nKunnen enkel bekijken (niet bewerken):\n' + DENYLIST.join('\n')
    );
  } catch(_) {}
}

function controleerKwartaaloverzicht_() {
  try {
    var ss = getSS_();
    if (!ss.getSheetByName('Kwartaaloverzicht')) {
      Logger.log('⚠️ Kwartaaloverzicht ontbreekt — automatische herbouw gestart.');
      maakKwartaaloverzicht();
    }
  } catch(e) { Logger.log('⚠️ Controle Kwartaaloverzicht mislukt: ' + e); }
}

function verversKwartaaloverzicht() {
  verversKwartaaloverzichtAlsBestaat_();
}

function ruimPersoneelsOudRijenOp_(ss, jaar, persDataPreloaded) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = Array.isArray(persDataPreloaded) ? persDataPreloaded : sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  try { sheet.showRows(2, data.length); } catch(_) {}
  for (var i = 0; i < data.length; i++) {
    var geldigheid = (data[i][8]||'').toString().trim();
    var melding    = (data[i][10]||'').toString().trim().toLowerCase();
    var meldDetail = (data[i][9]||'').toString().trim();
    var rijnr = i + 2;

    // Archiefrij (Tot datum): permanent verbergen als jaar < weergegeven jaar
    var totMatch = geldigheid.match(/^Tot\s+\d{2}\/\d{2}\/(\d{4})/);
    if (totMatch && parseInt(totMatch[1]) < jaar) {
      sheet.hideRows(rijnr);
      continue;
    }

    // Actieve gewijzigde rij (Vanaf datum): kolom I, J en K wissen als jaar < weergegeven jaar
    var vanafMatch = geldigheid.match(/^Vanaf\s+\d{2}\/\d{2}\/(\d{4})/);
    if (vanafMatch && parseInt(vanafMatch[1]) < jaar) {
      sheet.getRange(rijnr, 9).clearContent();
      sheet.getRange(rijnr, 10).clearContent();
      sheet.getRange(rijnr, 11).clearContent();
      continue;
    }

    // Nieuwe medewerker: kolom K wissen als invoerdatum < weergegeven jaar
    if (melding === 'nieuw') {
      var datumMatch = meldDetail.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (datumMatch && parseInt(datumMatch[3]) < jaar) {
        sheet.getRange(rijnr, 11).clearContent();
      }
    }
  }
}

function leesRittenVoorJaar_(ss, sheetNaam, jaar, jaarKolom) {
  jaarKolom = (jaarKolom !== undefined) ? jaarKolom : 6;
  var sheet = ss.getSheetByName(sheetNaam);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) { if (parseInt(data[i][jaarKolom]) === jaar) result.push(data[i]); }
  return result;
}


// ─── AGGREGATIE (1 rij per persoon) ──────────────────────────────────────────

function aggregeerFiets_(ritten, kw, tarief, personenMap, personenMapByEmail) {
  var map = {}, volgorde = [];
  ritten.forEach(function(r) {
    if (parseInt(r[4]) !== kw) return;
    var gs10 = (r[10]||'').toString().trim().toLowerCase(); // Nico-goedkKol (K, index 10)
    var gs12 = (r[12]||'').toString().trim().toLowerCase(); // Beeld-goedkKol (M, index 12)
    if (gs10.indexOf('afgekeurd') > -1 || gs12.indexOf('afgekeurd') > -1) return;
    var email   = (r[1]||'').toString().trim().toLowerCase();
    var naam    = (r[2]||'').toString().trim(); if (!naam && !email) return;
    var p = (personenMapByEmail && email && personenMapByEmail[email]) || (personenMap && naam && personenMap[naam]) || {};
    var displayNaam = p.naam || naam;
    var sleutel = (displayNaam || naam).toLowerCase().replace(/\s+/g, ' ').trim() || email;
    if (!map[sleutel]) {
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', email:email||'', totKm:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].totKm += parseFloat(r[7]||0);
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:g.iban, email:g.email||'',
             ritten:[{ col4:'Thuis ↔ Academie', col5:g.totKm, col6:g.totKm*tarief }], fmt5:'0.0 "km"', fmt6:'"€ "#,##0.00' };
  });
}

function aggregeerWoonwerk_(ritten, kw, personenMap, personenMapByEmail) {
  var map = {}, volgorde = [];
  ritten.forEach(function(r) {
    if ((r[0]||'').toString().charAt(0) === '_') return; // vervolrij (extra bewijs), niet meetellen
    if (parseInt(r[4]) !== kw) return;
    var _gs12 = (r[12]||'').toString().trim().toLowerCase(); // Nico-goedkKol (M, index 12)
    var _gs14 = (r[14]||'').toString().trim().toLowerCase(); // Beeld-goedkKol (O, index 14)
    if (_gs12.indexOf('afgekeurd') > -1 || _gs14.indexOf('afgekeurd') > -1) return;
    var email   = (r[1]||'').toString().trim().toLowerCase();
    var naam    = (r[2]||'').toString().trim(); if (!naam && !email) return;
    var p = (personenMapByEmail && email && personenMapByEmail[email]) || (personenMap && naam && personenMap[naam]) || {};
    var displayNaam = p.naam || naam;
    var sleutel = (displayNaam || naam).toLowerCase().replace(/\s+/g, ' ').trim() || email;
    if (!map[sleutel]) {
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', email:email||'', aantalRitten:0, totBedrag:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].aantalRitten++;
    map[sleutel].totBedrag += parseFloat(r[9]||0);
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:g.iban, email:g.email||'',
             ritten:[{ col4:g.aantalRitten + ' rit(ten)', col5:'', col6:g.totBedrag }], fmt5:null, fmt6:'"€ "#,##0.00' };
  });
}

function aggregeerDienst_(ritten, kw, tarief, personenMap, personenMapByEmail) {
  var map = {}, volgorde = [];
  ritten.forEach(function(r) {
    if (parseInt(r[4]) !== kw) return;
    var _dgs12 = (r[12]||'').toString().trim().toLowerCase(); // Nico-goedkKol (M, index 12)
    var _dgs14 = (r[14]||'').toString().trim().toLowerCase(); // Beeld-goedkKol (O, index 14)
    if (_dgs12.indexOf('afgekeurd') > -1 || _dgs14.indexOf('afgekeurd') > -1) return;
    var email   = (r[1]||'').toString().trim().toLowerCase();
    var naam    = (r[2]||'').toString().trim(); if (!naam && !email) return;
    var p = (personenMapByEmail && email && personenMapByEmail[email]) || (personenMap && naam && personenMap[naam]) || {};
    var displayNaam = p.naam || naam;
    var sleutel = (displayNaam || naam).toLowerCase().replace(/\s+/g, ' ').trim() || email;
    if (!map[sleutel]) {
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', email:email||'', totKm:0, aantalVerpl:0, totParkeer:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].totKm      += parseFloat(r[8]||0);
    map[sleutel].totParkeer += parseFloat(r[9]||0);
    map[sleutel].aantalVerpl++;
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    var parkeerTekst = g.totParkeer > 0 ? '€ ' + g.totParkeer.toFixed(2) : '—';
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:parkeerTekst, email:g.email||'',
             ritten:[{ col4:g.aantalVerpl + ' verplaatsing(en)', col5:g.totKm, col6:g.totKm*tarief }], fmt5:'0.0 "km"', fmt6:'"€ "#,##0.00' };
  });
}


// ─── SCHRIJF CATEGORIE ────────────────────────────────────────────────────────

function schrijfCategorie_(sheet, rij, titel, accentKleur, tekstkleur, groepen, headers, enkeleRij, tarief, sectie, jaar, kw, statusLookup, nicoTarief) {
  var statusValidatie = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_INGEDIEND, STATUS_CONTROLE, STATUS_BETALING], true).setAllowInvalid(false).build();

  sheet.getRange(rij, 1, 1, 7).setBackground('#ffffff'); sheet.setRowHeight(rij, 18); rij++;

  var titelRij = rij;
  sheet.getRange(rij, 1, 1, 7).merge()
    .setValue(titel).setBackground(accentKleur).setFontColor(tekstkleur)
    .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('left');
  sheet.setRowHeight(rij, 28); rij++;

  var tariefRij = rij;
  var headerRange = sheet.getRange(rij, 1, 1, 7);
  headerRange.setValues([headers]);
  headerRange.setBackgrounds([headers.map(function() { return '#e2e8f0'; })]);
  headerRange.setFontColors([headers.map(function() { return '#475569'; })]);
  headerRange.setFontWeights([headers.map(function() { return 'bold'; })]);
  headerRange.setFontSizes([headers.map(function() { return 8; })]);
  headerRange.setVerticalAlignments([headers.map(function() { return 'middle'; })]);
  if (tarief !== null && tarief !== undefined) {
    sheet.getRange(tariefRij, 8).setValue(tarief).setNumberFormat('"€ "#,##0.0000').setFontSize(9).setFontWeight('bold')
      .setFontColor('#7d2568').setBackground('#fce7f5').setHorizontalAlignment('center').setVerticalAlignment('middle');
    var tariefTekst = sectie === 'fiets' ? '← Vul het juiste tarief fietsvergoeding aan'
                    : sectie === 'dienst' ? '← Vul het juiste tarief dienstverplaatsing aan'
                    : null;
    if (tariefTekst) {
      sheet.getRange(tariefRij, 12).setValue(tariefTekst)
        .setFontSize(9).setFontStyle('italic').setFontColor('#1e40af')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    }
  }
  sheet.setRowHeight(rij, 22); rij++;

  if (groepen.length === 0) {
    sheet.getRange(rij, 1, 1, 7).merge()
      .setValue('Geen gegevens dit kwartaal').setFontColor('#94a3b8').setFontStyle('italic').setFontSize(9).setBackground('#ffffff');
    sheet.setRowHeight(rij, 24); rij++;
  } else {
    var grandTot5 = 0, grandTot6 = 0;
    var heeftCol6 = groepen[0].ritten[0].col6 !== null;

    // Batch-verzamelaars voor gewone kolommen 1–5, 7 en 9 (kolom 6, 8 en 11 blijven per-cel)
    var blokVals    = [];
    var blokBgs     = [];
    var blokFc      = [];
    var blokIban    = [];
    var blokIbanBgs = [];
    var blokIbanFc  = [];
    var blokEmail   = []; // col 9: e-mail als stabiele sleutel voor betaalstatus
    var startRijBlok = rij;

    groepen.forEach(function(g) {
      var tot5 = g.ritten.reduce(function(s,r){ return s+(r.col5||0); }, 0);
      var tot6 = heeftCol6 ? g.ritten.reduce(function(s,r){ return s+(r.col6||0); }, 0) : null;
      grandTot5 += tot5; if (tot6 !== null) grandTot6 += tot6;

      // Rij-brede achtergrond (identiek aan vóór; raakt ook col 6 — bewust ongewijzigd)
      sheet.getRange(rij, 1, 1, 7).setBackground('#ffffff');

      // Gewone kolommen 1–5, 7 en 9: waarden + kleuren verzamelen voor post-lus batch
      blokVals.push([g.naam, g.adres, g.pcGem, g.ritten[0].col4, tot5]);
      blokBgs.push( ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']);
      blokFc.push(  ['#334155', '#26295a', '#94a3b8', '#334155', '#26295a']);
      blokIban.push([g.iban]);
      blokIbanBgs.push(['#ffffff']);
      blokIbanFc.push(['#26295a']);
      blokEmail.push([g.email || '']); // col 9: e-mailadres voor statussleutel

      // Per-cel opmaak die per kolom verschilt (weight, size, numberFormat) — blijft per-cel
      sheet.getRange(rij, 1).setFontWeight('normal').setFontSize(9);
      sheet.getRange(rij, 2).setBackground('#ffffff').setFontWeight('bold').setFontSize(9);
      sheet.getRange(rij, 3).setBackground('#ffffff').setFontSize(8);
      sheet.getRange(rij, 4).setFontSize(9);
      sheet.getRange(rij, 5).setNumberFormat(g.fmt5).setFontWeight('bold').setFontSize(9);
      sheet.getRange(rij, 7).setFontSize(9).setFontWeight('bold');

      // Kolom 6 — ONGEWIJZIGD (formule / woon-werk bedrag)
      var isNico = (sectie === 'fiets' && g.adres === 'Logghe Nico');
      if (isNico) {
        sheet.getRange(rij, 8).setValue(nicoTarief || 0)
          .setNumberFormat('"€ "#,##0.0000').setFontSize(9).setFontWeight('bold')
          .setFontColor('#92400e').setBackground('#fef3c7')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRange(rij, 6).setFormula('=E'+rij+'*H'+rij)
          .setNumberFormat('"€ "#,##0.00;-"€ "#,##0.00;').setFontWeight('bold').setFontSize(9).setFontColor('#26295a');
      } else if (tarief !== null && tarief !== undefined) {
        sheet.getRange(rij, 6).setFormula('=E'+rij+'*$H$'+tariefRij).setNumberFormat(g.fmt6).setFontWeight('bold').setFontSize(9).setFontColor('#26295a');
      } else if (tot6 !== null) {
        sheet.getRange(rij, 6).setValue(tot6).setNumberFormat(g.fmt6).setFontWeight('bold').setFontSize(9).setFontColor('#26295a');
      }
      sheet.setRowHeight(rij, 22);

      // Kolom 11: betaalstatus — primaire sleutel op e-mail of naam, domein als overgangsval
      var primSleutel = jaar+'|'+kw+'|'+sectie+'|'+(g.email || g.adres);
      var oudeSleutel = jaar+'|'+kw+'|'+sectie+'|'+g.naam;
      var bestaandeStatus = (statusLookup && statusLookup[primSleutel]) ? statusLookup[primSleutel]
                          : (statusLookup && statusLookup[oudeSleutel]) ? statusLookup[oudeSleutel]
                          : STATUS_INGEDIEND;
      sheet.getRange(rij, 11).setValue(bestaandeStatus).setDataValidation(statusValidatie)
        .setBackground(statusKleur_(bestaandeStatus)).setFontSize(9).setFontColor('#334155')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      rij++;

      // Spacer-rij — logica ongewijzigd; lege rij in batch-arrays
      blokVals.push(['', '', '', '', '']);
      blokBgs.push( ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']);
      blokFc.push(  ['#334155', '#26295a', '#94a3b8', '#334155', '#26295a']);
      blokIban.push(['']);
      blokIbanBgs.push(['#ffffff']);
      blokIbanFc.push(['#26295a']);
      blokEmail.push(['']);
      sheet.getRange(rij, 1, 1, 7).setBackground('#ffffff'); sheet.setRowHeight(rij, 6); rij++;
    });

    // Gebatched schrijven: gewone kolommen 1–5, 7 en 9 in één blok over alle personen
    if (blokVals.length > 0) {
      var aantalRijen = blokVals.length;
      sheet.getRange(startRijBlok, 1, aantalRijen, 5)
        .setValues(blokVals).setBackgrounds(blokBgs).setFontColors(blokFc);
      sheet.getRange(startRijBlok, 7, aantalRijen, 1)
        .setValues(blokIban).setBackgrounds(blokIbanBgs).setFontColors(blokIbanFc);
      sheet.getRange(startRijBlok, 9, aantalRijen, 1).setValues(blokEmail);
    }
  }
  return rij;
}
