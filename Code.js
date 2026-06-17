// ============================================
// CONFIGURATIE — pas hier elk jaar aan
// ============================================
const CONFIG = {
  SPREADSHEET_ID: '1zVg9MSZVE_EDhLI19rwBKF_tnaUGdNAa00Ls-2vSEQQ',
  SHEETS: {
    PERSONEELSGEGEVENS: 'Personeelsgegevens',
    FIETSVERGOEDING:    'Fietsvergoeding',
    WOON_WERK:          'Woon-Werk (trein-De Lijn)',
    DIENSTVERPLAATSING: 'Dienstverplaatsing',
  },
  DRIVE_MAP_ID:      '14LoAJybVBjvEo3n-KxYmwoeonX6qEUab',  // map voor bewijsstukken (uploads)
  BACKUP_MAP_ID:     '14LoAJybVBjvEo3n-KxYmwoeonX6qEUab',  // map voor wekelijkse backups — pas aan naar eigen backup-map
  TEST_DATUM:        null,   // 'YYYY-MM-DD' of null
  TEST_MODUS:        true,
  TEST_EMAIL:        'Nele.Moerman@academie-ieper.be',
  RECIPIENT_EMAILS:  'Nele.Delameilleure@ieper.be',  // komma-gescheiden bij meerdere ontvangers
  ADMIN_EMAILS:      'nele.delameilleure@academie-ieper.be',
};
// ============================================
// EINDE CONFIGURATIE
// ============================================

// ─── KOLOMINDEXEN (0-gebaseerd) ──────────────────────────────────────────────
//
// PERSONEELSGEGEVENS:
//   0:domein  1:naam  2:adres  3:postcodeGem  4:iban  5:privémail  6:rijksreg  7:werkemail
//   8:geldigheid  9:meldingDetail (omschrijving)  10:melding ('nieuw'/'gewijzigd')
//
// FIETSVERGOEDING:
//   0:id  1:email  2:naam  3:jaar  4:kwartaal  5:maand  6:bestemming  7:km  8:bestand  9:datum
//
// WOON-WERK:
//   0:id  1:email  2:naam  3:jaar  4:kwartaal  5:maand  6:vertrekpunt  7:bestemming  8:vervoersmiddel  9:bedrag  10:bestand  11:invoerdatum
//
// DIENSTVERPLAATSING:
//   0:id  1:email  2:naam  3:jaar  4:kwartaal  5:datum  6:bestemming  7:doel  8:km  9:parkeerBedrag  10:bestand  11:invoerdatum
// ─────────────────────────────────────────────────────────────────────────────


// ─── SPREADSHEET CACHE ───────────────────────────────────────────────────────
var _ss = null;
function getSS_() {
  if (!_ss) _ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return _ss;
}

// ─── DATUM / KWARTAAL HULPFUNCTIES ───────────────────────────────────────────

function getEffectiveDate() {
  if (CONFIG.TEST_DATUM) {
    var p = String(CONFIG.TEST_DATUM).split('-');
    if (p.length === 3) return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  }
  var testDate = getConfig('TEST_DATE');
  if (testDate && testDate.trim() !== '') {
    var parts = testDate.trim().split('-');
    if (parts.length === 3) return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return new Date();
}

function getQuarterFromDate(date) {
  var m = date.getMonth() + 1;
  return m <= 3 ? 1 : m <= 6 ? 2 : m <= 9 ? 3 : 4;
}

function getQuarterEndDate(quarter, year) {
  var ends = { 1: [2, 31], 2: [5, 30], 3: [8, 30], 4: [11, 31] };
  var e = ends[quarter];
  return new Date(year, e[0], e[1], 23, 59, 59);
}

function getQuarterDeadline(quarter, year) {
  if (quarter === 2) {
    return new Date(year, 6, 5, 23, 59, 59);  // Uiterste indieningstermijn Q2 = altijd 5 juli
  }
  var end = getQuarterEndDate(quarter, year);
  var dl = new Date(end);
  dl.setDate(dl.getDate() + 10);
  return dl;
}

function isGracePeriod(quarter, year, referenceDate) {
  var date = referenceDate || getEffectiveDate();
  var end  = getQuarterEndDate(quarter, year);
  var dl   = getQuarterDeadline(quarter, year);
  return date > end && date <= dl;
}

function isLocked(quarter, year, referenceDate) {
  var date = referenceDate || getEffectiveDate();
  return date > getQuarterDeadline(quarter, year);
}

function getPreviousQuarter(quarter, year) {
  return quarter === 1 ? { quarter: 4, year: year - 1 } : { quarter: quarter - 1, year: year };
}

function getCurrentQuarterInfo() {
  var today   = getEffectiveDate();
  var curQ    = getQuarterFromDate(today);
  var curYear = today.getFullYear();
  var prev    = getPreviousQuarter(curQ, curYear);
  return {
    today:          formatDatumWeergave_(today),
    currentQuarter: curQ,
    currentYear:    curYear,
    prevQuarter:    prev.quarter,
    prevYear:       prev.year,
    isGracePeriod:  isGracePeriod(prev.quarter, prev.year, today),
    isPrevLocked:   isLocked(prev.quarter, prev.year, today),
    prevDeadline:   formatDatumWeergave_(getQuarterDeadline(prev.quarter, prev.year)),
    testDateActive: !!(CONFIG.TEST_DATUM || (getConfig('TEST_DATE') && getConfig('TEST_DATE').trim() !== '')),
  };
}

function validateEntryDate(dateString) {
  var date  = parseDateStr_(dateString);
  var today = getEffectiveDate();
  if (date > today) return { valid: false, error: 'De datum mag niet in de toekomst liggen.' };
  var kw   = getQuarterFromDate(date);
  var jaar = date.getFullYear();
  if (isLocked(kw, jaar)) {
    return {
      valid: false,
      error: '⛔ De indieningstermijn voor K' + kw + ' ' + jaar + ' is verstreken op ' +
             formatDatumWeergave_(getQuarterDeadline(kw, jaar)) + '. Aanpassing is niet meer mogelijk.',
    };
  }
  return { valid: true, quarter: kw, year: jaar };
}

function formatDatumWeergave_(date) {
  var d = new Date(date);
  return String(d.getDate()).padStart(2,'0') + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
}

function parseDateStr_(str) {
  var p = str.split('-');
  return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
}

var MAANDNAMEN_SERVER_ = ['januari','februari','maart','april','mei','juni',
                          'juli','augustus','september','oktober','november','december'];

function formatMaandJaarServer_(datumStr) {
  if (!datumStr) return '';
  var p = datumStr.toString().split('-');
  var m = parseInt(p[1]) - 1;
  var j = parseInt(p[0]);
  return MAANDNAMEN_SERVER_[m] + ' ' + j;
}


// ─── STATUS-WAARDEN ──────────────────────────────────────────────────────────
var STATUS_INGEDIEND = 'Ingediend';
var STATUS_CONTROLE  = 'Controle uitgevoerd';
var STATUS_BETALING  = 'Betaling verwerkt';

function statusKleur_(status) {
  if (status === STATUS_BETALING) return '#fef9c3';
  if (status === STATUS_CONTROLE) return '#bbf7d0';
  return '#f1f5f9';
}

function goedkeuringStatusKleur_(status) {
  if (status === 'Ja, goedgekeurd') return { bg: '#bbf7d0', txt: '#14532d' };
  if (status === 'Nee, afgekeurd')  return { bg: '#fee2e2', txt: '#1e293b' };
  return { bg: '#f1f5f9', txt: '#64748b' };
}


// ─── SPREADSHEET MENU (verschijnt automatisch bij openen van het spreadsheet) ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Academie Ieper')
    .addItem('Kwartaaloverzicht vernieuwen', 'maakKwartaaloverzicht')
    .addItem('Tabbladen sorteren op personeelslid', 'sorteerVerplaatsingenSheets')
    .addItem('Goedkeuringkolommen instellen', 'stelGoedkeuringKolommenIn')
    .addSeparator()
    .addItem('Installeer wekelijkse backup', 'installeerWekelijkseBackup')
    .addItem('Maak nu een backup', 'maakBackupNu')
    .addSeparator()
    .addItem('Tabblad Aanpak aanmaken', 'maakAanpakTabblad')
    .addSeparator()
    .addItem('Beveiliging Kwartaaloverzicht instellen', 'stelKwartaaloverzichtBeveiligingIn')
    .addToUi();
}


// ─── WEB APP ENTRY POINT ─────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Terugbetaling Woon-Werkverkeer — Academie Ieper')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// ─── HELPERS ─────────────────────────────────────────────────────────────────
function getEmail_() {
  return CONFIG.TEST_MODUS ? CONFIG.TEST_EMAIL : Session.getActiveUser().getEmail();
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function generateId_() {
  return new Date().getTime().toString(36) + Math.random().toString(36).substr(2, 5);
}

function appendRijNaData_(sheet, rij) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, rij.length).setValues([rij]);
    return 2;
  }
  // Scan e-mailkolom (kolom B) — die is altijd gevuld bij echte data en leeg na wissen
  var kolB    = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var doelRij = 2;
  for (var r = 0; r < kolB.length; r++) {
    if ((kolB[r][0]||'').toString().trim() !== '') doelRij = r + 3;
  }
  sheet.getRange(doelRij, 1, 1, rij.length).setValues([rij]);
  return doelRij;
}

function zetGoedkeuringInOpRij_(sheet, rijnr, categorie, isBeeld) {
  var nicoKol  = (categorie === 'fiets') ? 11 : 13;
  var checkKol = (categorie === 'fiets') ? 12 : 14;
  var beeldKol = (categorie === 'fiets') ? 13 : 15;
  var doelKol  = isBeeld ? beeldKol : nicoKol;
  var leegKol  = isBeeld ? nicoKol  : beeldKol;
  var validatie = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ingediend', 'Ja, goedgekeurd', 'Nee, afgekeurd'], true).setAllowInvalid(false).build();
  var goedkCel = sheet.getRange(rijnr, doelKol);
  var huidig = (goedkCel.getValue()||'').toString().trim();
  if (huidig === 'Ja')        huidig = 'Ja, goedgekeurd';
  else if (huidig === 'Nee') huidig = 'Nee, afgekeurd';
  else if (!huidig)           huidig = 'Ingediend';
  var kleur = goedkeuringStatusKleur_(huidig);
  goedkCel.setValue(huidig).setBackground(kleur.bg).setFontColor(kleur.txt);
  // 'Beeld' is een overgebleven plaatshouder uit de andere kolom, geen echte status: geen keuzelijst erop
  if (huidig === 'Beeld') goedkCel.clearDataValidations();
  else goedkCel.setDataValidation(validatie);
  // Lege goedkeuringskolom: 'Beeld' in witte letter voor Beeld-rijen (blokkeert doorloop van hyperlink)
  var leegCel = sheet.getRange(rijnr, leegKol);
  leegCel.clearDataValidations().setBackground('#ffffff');
  if (isBeeld) {
    leegCel.setValue('Beeld').setFontColor('#ffffff');
  } else {
    leegCel.clearContent().setFontColor(null);
  }
  sheet.getRange(rijnr, checkKol).insertCheckboxes();
}

function huidigKwartaal_() {
  return getQuarterFromDate(getEffectiveDate());
}

// Geeft een array van {kwartaal, jaar} terug die in aanmerking komen voor bulk-goedkeuring.
// Normaal is dat één item (het huidige kwartaal). In de overgangsperiode — de eerste dagen
// van een nieuw kwartaal, voordat de invulstop voor leerkrachten ingaat — worden ook de
// rijen van het vorige kwartaal meegenomen (inclusief jaarwissel bij Q4→Q1).
function bepaalActieveKwartalen_() {
  var nu        = getEffectiveDate();
  var huidigKw  = huidigKwartaal_();
  var huidigJaar = nu.getFullYear();
  var maand     = nu.getMonth() + 1;
  var dag       = nu.getDate();

  // stopDag = eerste dag waarop leerkrachten NIET meer kunnen invullen (cfr. instructiekader)
  var kwInfo = {
    1: { startMaand: 1,  stopDag: 11 },
    2: { startMaand: 4,  stopDag: 11 },
    3: { startMaand: 7,  stopDag: 6  },
    4: { startMaand: 10, stopDag: 11 }
  };

  var actief = [{ kwartaal: huidigKw, jaar: huidigJaar }];
  var info = kwInfo[huidigKw];
  if (maand === info.startMaand && dag < info.stopDag) {
    var vorigKw   = huidigKw === 1 ? 4 : huidigKw - 1;
    var vorigJaar = huidigKw === 1 ? huidigJaar - 1 : huidigJaar;
    actief.push({ kwartaal: vorigKw, jaar: vorigJaar });
  }
  return actief;
}

function sheetNaamVoorCategorie_(categorie) {
  return { fiets: CONFIG.SHEETS.FIETSVERGOEDING, woonwerk: CONFIG.SHEETS.WOON_WERK, dienst: CONFIG.SHEETS.DIENSTVERPLAATSING }[categorie];
}

function formatDatum_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0') + '-' + String(v.getDate()).padStart(2,'0');
  }
  return v.toString();
}

function formatDatumNL_(date) {
  var d = date instanceof Date ? date : new Date(date);
  return String(d.getDate()).padStart(2,'0') + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
}

function fileIdUitUrl_(url) {
  if (!url) return '';
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

var MAANDNAMEN_NL_ = ['Januari','Februari','Maart','April','Mei','Juni',
                      'Juli','Augustus','September','Oktober','November','December'];

function bouwBestandsnaam_(data, ritIdx, email, categorie, fileNumInRit) {
  var rij  = data[ritIdx];
  var jaar = parseInt(rij[3]) || 0;
  var kw   = parseInt(rij[4]) || 0;
  var naam = (rij[2]||'').toString().trim();

  var maandStr;
  if (categorie === 'dienst') {
    var datRaw = rij[5] instanceof Date
      ? Utilities.formatDate(rij[5], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (rij[5]||'').toString().trim();
    var mIdx = /^\d{4}-(\d{2})-\d{2}$/.test(datRaw) ? parseInt(datRaw.split('-')[1]) - 1 : -1;
    maandStr = mIdx >= 0 ? MAANDNAMEN_NL_[mIdx] : datRaw;
  } else {
    if (rij[5] instanceof Date) {
      maandStr = MAANDNAMEN_NL_[rij[5].getMonth()];
    } else {
      var maandRaw = (rij[5]||'').toString().trim().toLowerCase();
      var mIdx2    = MAANDNAMEN_SERVER_.indexOf(maandRaw.split(' ')[0]);
      maandStr = mIdx2 >= 0 ? MAANDNAMEN_NL_[mIdx2] : (maandRaw.charAt(0).toUpperCase() + maandRaw.slice(1));
    }
  }

  // fileNumInRit 0 = enkel bestand (geen getal), 1+ = volgnummer binnen rit
  var type = categorie === 'fiets' ? 'F' : categorie === 'dienst' ? 'D'
           : fileNumInRit === 0 ? 'WW' : 'WW-' + fileNumInRit;

  // Rit-volgnummer: hoeveel hoofdrijen van dezelfde persoon/kwartaal staan vóór deze rij
  var ritSeqNr = 1;
  for (var r = 1; r < ritIdx; r++) {
    if ((data[r][0]||'').toString().charAt(0) === '_') continue;
    if ((data[r][1]||'').toString().toLowerCase().trim() !== email.toLowerCase().trim()) continue;
    if (parseInt(data[r][3]) !== jaar || parseInt(data[r][4]) !== kw) continue;
    ritSeqNr++;
  }

  var basis = jaar + ' (K' + kw + ') - ' + maandStr + ' - ' + naam + ' (' + type + ')';
  if (categorie !== 'fiets' || ritSeqNr > 1) basis += ' (' + ritSeqNr + ')';
  return basis;
}

function isZelfdePersoon_(naam1, naam2) {
  var norm = function(n) { return (n||'').toLowerCase().trim().split(/\s+/).sort().join(' '); };
  return norm(naam1) === norm(naam2);
}

function naamUitEmail_(email) {
  var lokaal = (email || '').split('@')[0];
  var delen  = lokaal.split('.');
  if (delen.length < 2) return lokaal;
  var kapitaliseer = function(s) { return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); };
  var voornaam    = kapitaliseer(delen[0]);
  var familienaam = delen.slice(1).map(kapitaliseer).join(' ');
  return familienaam + ' ' + voornaam;
}

// Twee e-mailadressen matchen als ze exact gelijk zijn, OF als ze hetzelfde lokaal deel
// hebben én allebei tot de ieper.be-domeinfamilie behoren (academie-ieper.be ↔ ieper.be).
function emailsKommenOvereen_(opgeslagen, sessie) {
  var a = (opgeslagen || '').toString().toLowerCase().trim();
  var b = (sessie     || '').toString().toLowerCase().trim();
  if (!a || !b) return false;
  if (a === b) return true;
  var isIeper = function(d) { return d === 'ieper.be' || d === 'academie-ieper.be'; };
  var pA = a.split('@'), pB = b.split('@');
  return pA.length === 2 && pB.length === 2 &&
         isIeper(pA[1]) && isIeper(pB[1]) && pA[0] === pB[0];
}


// ─── MELDING WISSEN ──────────────────────────────────────────────────────────
function clearMelding_(email) {
  try {
    var ss    = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getDataRange().getValues();
    var lastMatch = -1;
    for (var i = 1; i < data.length; i++) {
      if (emailsKommenOvereen_(data[i][7], email)) lastMatch = i;
    }
    if (lastMatch >= 0) sheet.getRange(lastMatch + 1, 10, 1, 2).setValues([['', '']]);
  } catch(e) { Logger.log('clearMelding_: ' + e); }
}


// ─── CONFIG SHEET ────────────────────────────────────────────────────────────
function getConfig(key) {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('cfg_' + key);
    if (cached !== null) return cached;
    var ss    = getSS_();
    var sheet = ss.getSheetByName('Config');
    if (!sheet || sheet.getLastRow() < 1) { cache.put('cfg_' + key, '', 120); return ''; }
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0]||'').toString().trim() === key) {
        var val = (data[i][1]||'').toString().trim();
        cache.put('cfg_' + key, val, 120);
        return val;
      }
    }
    cache.put('cfg_' + key, '', 120);
    return '';
  } catch(e) { return ''; }
}

function setConfig(key, value) {
  try {
    try { CacheService.getScriptCache().remove('cfg_' + key); } catch(_) {}
    var ss    = getSS_();
    var sheet = getOrCreateSheet_(ss, 'Config');
    var data  = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0]||'').toString().trim() === key) { sheet.getRange(i+1, 2).setValue(value); return; }
    }
    sheet.appendRow([key, value]);
  } catch(e) { Logger.log('setConfig: ' + e); }
}


function isAdminEmail(email) {
  var admins = CONFIG.ADMIN_EMAILS || '';
  return admins.toLowerCase().indexOf((email||'').toLowerCase().trim()) > -1;
}


// ─── PERSONEELSGEGEVENS ───────────────────────────────────────────────────────

function getUserData() {
  try {
    var email = getEmail_();
    if (!email) return { error: 'Kon je e-mailadres niet ophalen.' };
    var ss       = getSS_();
    var gevonden = null;

    var persSheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
    if (persSheet && persSheet.getLastRow() >= 2) {
      var persData = persSheet.getDataRange().getValues();

      // Primaire opzoeking: op e-mailadres (exact of zelfde lokaal deel binnen *.ieper.be)
      for (var i = 1; i < persData.length; i++) {
        if (emailsKommenOvereen_(persData[i][7], email)) {
          gevonden = {
            domein:              (persData[i][0]||'').toString().trim(),
            naam:                (persData[i][1]||'').toString().trim(),
            adres:               (persData[i][2]||'').toString().trim(),
            postcodeGem:         (persData[i][3]||'').toString().trim(),
            iban:                (persData[i][4]||'').toString().trim(),
            email:               email,
            rijksregisternummer: (persData[i][6]||'').toString().trim(),
            priveEmail:          (persData[i][5]||'').toString().trim(),
          };
          // geen break: laatste (meest recente) rij wint
        }
      }

      // Fallback: naam afleiden uit e-mailadres (bv. Nele.Moerman@ → "Nele Moerman")
      // Handig voor handmatig ingevoerd personeel zonder e-mail in kolom H
      if (!gevonden) {
        var lokalDeel = email.split('@')[0].replace(/[._-]/g, ' ');
        for (var j = 1; j < persData.length; j++) {
          var rijnaam = (persData[j][1]||'').toString();
          if (isZelfdePersoon_(rijnaam, lokalDeel)) {
            gevonden = {
              domein:              (persData[j][0]||'').toString().trim(),
              naam:                rijnaam.trim(),
              adres:               (persData[j][2]||'').toString().trim(),
              postcodeGem:         (persData[j][3]||'').toString().trim(),
              iban:                (persData[j][4]||'').toString().trim(),
              email:               email,
              rijksregisternummer: (persData[j][6]||'').toString().trim(),
              priveEmail:          (persData[j][5]||'').toString().trim(),
            };
            // geen break: laatste rij wint
          }
        }
      }
    }

    var nu = getEffectiveDate();
    return {
      email:          email,
      personData:     gevonden,
      naamSuggestie:  gevonden ? null : naamUitEmail_(email),
      kwartaal:       huidigKwartaal_(),
      jaar:           nu.getFullYear(),
      kwartaalInfo:   getCurrentQuarterInfo(),
      isAdmin:        isAdminEmail(email),
    };
  } catch(e) { Logger.log('❌ getUserData: ' + e); return { error: e.toString() }; }
}

function saveUserDataAlle(personData) {
  try {
    var email     = getEmail_();
    var ss        = getSS_();
    var persSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.PERSONEELSGEGEVENS);
    var persData  = persSheet.getLastRow() >= 2 ? persSheet.getDataRange().getValues() : [[]];

    var domein        = (personData.domein         ||'').trim();
    var naam          = (personData.naam          ||'').trim();
    var adres         = (personData.adres         ||'').trim();
    var postcodeGem   = (personData.postcodeGem   ||'').trim();
    var iban          = (personData.iban          ||'').trim();
    var rijksreg      = (personData.rijksregisternummer||'').trim();
    var privmail      = (personData.priveEmail    ||'').trim();

    var bestaandRij = -1;
    for (var i = 1; i < persData.length; i++) {
      if (emailsKommenOvereen_(persData[i][7], email)) bestaandRij = i + 1;
    }
    // Fallback op naam als e-mailadres niet teruggevonden werd (bv. handmatig ingevoerd personeelslid)
    if (bestaandRij < 0) {
      for (var i = 1; i < persData.length; i++) {
        if (isZelfdePersoon_((persData[i][1]||'').toString(), naam)) bestaandRij = i + 1;
      }
    }

    if (bestaandRij > 0) {
      var oudData         = persData[bestaandRij - 1];
      var oudDomein       = (oudData[0]||'').toString().trim();
      var oudeNaam        = (oudData[1]||'').toString().trim();
      var oudeAdres       = (oudData[2]||'').toString().trim();
      var oudePostcodeGem = (oudData[3]||'').toString().trim();
      var oudeIban        = (oudData[4]||'').toString().trim();
      var oudeRijks       = (oudData[6]||'').toString().trim();
      var oudePrivmail    = (oudData[5]||'').toString().trim();

      var naamTeOpslaan  = oudeNaam || naamUitEmail_(email);
      var adresGewijzigd = (adres !== oudeAdres || postcodeGem !== oudePostcodeGem);
      var andereWijziging = (!isZelfdePersoon_(naam, oudeNaam) || iban !== oudeIban || rijksreg !== oudeRijks || privmail !== oudePrivmail || domein !== oudDomein);

      if (!adresGewijzigd && !andereWijziging) return { ok: true };

      var ingangsDatum = (adresGewijzigd && personData.effectiveDate)
        ? parseDateStr_(personData.effectiveDate)
        : getEffectiveDate();

      var geldigTotStr   = 'Tot '   + formatDatumNL_(ingangsDatum);
      var geldigVanafStr = 'Vanaf ' + formatDatumNL_(ingangsDatum);

      var wijzigingen = [];
      if (domein !== oudDomein)               wijzigingen.push('Domein: ' + oudDomein + ' → ' + domein);
      if (!isZelfdePersoon_(naam, oudeNaam))  wijzigingen.push('Naam: ' + oudeNaam + ' → ' + naamTeOpslaan);
      if (adres !== oudeAdres)                wijzigingen.push('Adres: ' + oudeAdres + ' → ' + adres);
      if (postcodeGem !== oudePostcodeGem)    wijzigingen.push('Postcode/Gem: ' + oudePostcodeGem + ' → ' + postcodeGem);
      if (iban !== oudeIban)                  wijzigingen.push('IBAN gewijzigd');
      if (rijksreg !== oudeRijks)             wijzigingen.push('Rijksregister gewijzigd');
      if (privmail !== oudePrivmail)          wijzigingen.push('Privé-e-mail gewijzigd');

      // Archiveer huidige rij
      persSheet.getRange(bestaandRij, 9).setValue(geldigTotStr);
      persSheet.getRange(bestaandRij, 10, 1, 2).setValues([['', '']]);

      // Voeg nieuwe rij in direct ONDER de huidige rij
      persSheet.insertRowAfter(bestaandRij);
      persSheet.getRange(bestaandRij + 1, 1, 1, 11).setValues([[
        domein, naamTeOpslaan, adres, postcodeGem, iban, privmail, rijksreg, email,
        geldigVanafStr, wijzigingen.join('; '), 'gewijzigd'
      ]]);

    } else {
      var naamOpgeslagen = naamUitEmail_(email);
      persSheet.appendRow([domein, naamOpgeslagen, adres, postcodeGem, iban, privmail, rijksreg, email, '',
        'Nieuw personeelslid ingevoerd via de app op ' + formatDatumNL_(getEffectiveDate()), 'nieuw']);
    }

    // Meteen vernieuwen (niet via timer) zodat de aanpassing zichtbaar is
    try { sorteerPersoneelsSheet_(ss); } catch(_) {}
    try { verversKwartaaloverzichtAlsBestaat_(); } catch(_) {}
    return { ok: true };
  } catch(e) { Logger.log('❌ saveUserDataAlle: ' + e); return { ok: false, error: e.toString() }; }
}

function clearMeldingViaStatus_(email, nieuweStatus, handledKw, jaar) {
  // Aangeroepen vanuit onEditJaar wanneer Stad Ieper "Wijziging doorgevoerd" instelt
  if (nieuweStatus !== 'Wijziging doorgevoerd') return;
  if (handledKw > 0 && jaar > 0) {
    var normEmail = email.toLowerCase().trim().replace(/[^a-z0-9@._-]/g, '_');
    PropertiesService.getScriptProperties().setProperty('meldingHandledKw_' + normEmail + '_' + jaar, handledKw.toString());
  }
  verversKwartaaloverzichtAlsBestaat_();
}

function getPersonData_(email, ss) {
  var persSheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!persSheet || persSheet.getLastRow() < 2) return {};
  var data = persSheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    if (emailsKommenOvereen_(data[i][7], email)) {
      result = {
        domein:      (data[i][0]||'').toString().trim(),
        naam:        (data[i][1]||'').toString().trim(),
        adres:       (data[i][2]||'').toString().trim(),
        postcodeGem: (data[i][3]||'').toString().trim(),
        iban:        (data[i][4]||'').toString().trim(),
      };
      // geen break: laatste rij wint
    }
  }
  return result;
}

function bouwPersonenMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var naam = (data[i][1]||'').toString().trim();
    if (naam) {
      map[naam] = {
        domein:      (data[i][0]||'').toString().trim(),
        adres:       (data[i][2]||'').toString().trim(),
        postcodeGem: (data[i][3]||'').toString().trim(),
        iban:        (data[i][4]||'').toString().trim(),
        rijksreg:    (data[i][6]||'').toString().trim(),
        melding:     (data[i][10]||'').toString().trim(),
      };
    }
  }
  return map;
}

function bouwPersonenMapByEmail_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var email = (data[i][7]||'').toString().trim().toLowerCase();
    if (!email) continue;
    var entry = {
      naam:     (data[i][1]||'').toString().trim(),
      domein:   (data[i][0]||'').toString().trim(),
      rijksreg: (data[i][6]||'').toString().trim(),
      iban:     (data[i][4]||'').toString().trim(),
    };
    map[email] = entry;
    // Sla ook op onder alternatief domein (academie-ieper.be ↔ ieper.be)
    if (email.indexOf('@academie-ieper.be') > -1)
      map[email.replace('@academie-ieper.be', '@ieper.be')] = entry;
    else if (email.indexOf('@ieper.be') > -1)
      map[email.replace('@ieper.be', '@academie-ieper.be')] = entry;
  }
  return map;
}


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
  try {
    var email = getEmail_();
    var ss    = getSS_();
    var sheet = ss.getSheetByName(sheetNaamVoorCategorie_(categorie));
    if (!sheet) return { ok: false, error: 'Tabblad niet gevonden.' };

    var data    = sheet.getDataRange().getValues();
    var jaar = 0, kwartaal = 0;
    // Verwijder vervolgrijen (woonwerk multi-bewijs) én hoofdrij, van onder naar boven
    for (var i = data.length - 1; i >= 1; i--) {
      var rijId    = (data[i][0]||'').toString();
      var rijEmail = (data[i][1]||'').toString().toLowerCase().trim();
      if (rijEmail !== email.toLowerCase().trim()) continue;
      var isHoofd  = rijId === id;
      var isVervol = rijId === '_' + id + '_2' || rijId === '_' + id + '_3' || rijId === '_' + id + '_4';
      if (!isHoofd && !isVervol) continue;
      if (isHoofd) { jaar = parseInt(data[i][3]); kwartaal = parseInt(data[i][4]); }
      sheet.deleteRow(i + 1);
    }
    if (jaar) invalideerRittenCache_(email, categorie, jaar, kwartaal);
    try { sorteerSheet_(sheet, bouwEmailNaamMap_(ss)); } catch(_) {}
    var _ckk2 = checkKolVoorSheet_(sheet.getName()); if (_ckk2) { try { kleurBeeldCheckKol_(ss, sheet, _ckk2); } catch(_e) {} }
    return { ok: true };
  } catch(e) { Logger.log('❌ deleteRit: ' + e); return { ok: false, error: e.toString() }; }
}


// ─── TESTDATUM BEHEER (admin UI) ─────────────────────────────────────────────

function getAdminTestDate() {
  return { testDate: getConfig('TEST_DATE') || '', isAdmin: isAdminEmail(getEmail_()) };
}

function saveAdminTestDate(dateStr) {
  if (!isAdminEmail(getEmail_())) return { ok: false, error: 'Geen beheerdersrechten.' };
  setConfig('TEST_DATE', dateStr || '');
  return { ok: true };
}
