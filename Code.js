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


// ─── TESTDATUM BEHEER (admin UI) ─────────────────────────────────────────────

function getAdminTestDate() {
  return { testDate: getConfig('TEST_DATE') || '', isAdmin: isAdminEmail(getEmail_()) };
}

function saveAdminTestDate(dateStr) {
  if (!isAdminEmail(getEmail_())) return { ok: false, error: 'Geen beheerdersrechten.' };
  setConfig('TEST_DATE', dateStr || '');
  return { ok: true };
}


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
      var naam = col0.replace(/^Totaal\s+/,'').trim();
      if (naam) result[jaar + '|' + huidigKw + '|' + huidigeSectie + '|' + naam] = status;
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

function maakKwartaaloverzicht() {
  var terugNaam;
  try { var _a = SpreadsheetApp.getActiveSheet(); terugNaam = _a ? _a.getName() : null; } catch(_) { terugNaam = null; }
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
  if (bestaand) {
    var hData = bestaand.getDataRange().getValues();
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


  var statusLookup = leesStatussenUitBestaandeSheet_(bestaand);

  // Lees vergrendelingstatus per kwartaal uit PropertiesService (blijft bewaard bij herbouw)
  var props = PropertiesService.getScriptProperties();
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
  var personenMap = bouwPersonenMap_(ss);
  var personenMapByEmail = bouwPersonenMapByEmail_(ss);

  // Verberg/wis oude personeelsrijen uit vorige jaren
  ruimPersoneelsOudRijenOp_(ss, jaar);

  // Personeelsmeldingen per kwartaal — enkel voor het weergegeven jaar
  var meldingPerKw = { 1:[], 2:[], 3:[], 4:[] };
  var persSheetM = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (persSheetM && persSheetM.getLastRow() >= 2) {
    var persDataM = persSheetM.getDataRange().getValues();
    for (var mi = 1; mi < persDataM.length; mi++) {
      var mStatus = (persDataM[mi][10]||'').toString().trim().toLowerCase();
      if (mStatus !== 'nieuw' && mStatus !== 'gewijzigd') continue;
      var geldigheid = (persDataM[mi][8]||'').toString().trim();
      var meldDetail = (persDataM[mi][9]||'').toString().trim();
      var mVanaf = geldigheid.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (!mVanaf) mVanaf = meldDetail.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (!mVanaf || parseInt(mVanaf[3]) !== jaar) continue;
      var meldDatum = new Date(parseInt(mVanaf[3]), parseInt(mVanaf[2])-1, parseInt(mVanaf[1]));
      var mKw = getQuarterFromDate(meldDatum);
      var melEmail   = (persDataM[mi][7]||persDataM[mi][5]||'').toString().toLowerCase().trim();
      var normEmail  = melEmail.replace(/[^a-z0-9@._-]/g, '_');
      var handledKw  = parseInt(props.getProperty('meldingHandledKw_' + normEmail + '_' + jaar)) || 0;
      var melObj = { domein: (persDataM[mi][0]||'').toString().trim(), naam: (persDataM[mi][1]||'').toString().trim(), email: melEmail, type: mStatus, handledKw: handledKw };
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
    var isGesloten = isHuidigJaar && isLocked(kw, jaar);
    var kwLabel    = kwNamen[kw];
    if (isHuidig)   kwLabel += '   ★ huidig kwartaal';
    if (isGesloten) kwLabel += '   🔒 Afgesloten op ' + formatDatumWeergave_(getQuarterDeadline(kw, jaar));

    var headerRij = rij;
    var kwHeader  = sheet.getRange(rij, 1, 1, 7);
    kwHeader.merge()
      .setValue(kwLabel)
      .setBackground(isHuidig ? '#f59e0b' : isGesloten ? '#64748b' : '#ffffff')
      .setFontColor(isHuidig ? '#7c2d12' : isGesloten ? '#ffffff' : '#94a3b8')
      .setFontSize(isHuidig ? 11 : 10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    kwHeader.setBorder(true, null, null, null, null, null, isHuidig ? '#ea580c' : '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(headerRij, 8)
      .setValue(isHuidig)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())
      .setBackground(isHuidig ? '#ede9fe' : isGesloten ? '#f1f5f9' : '#ffffff')
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
    var isInitieelGeopend = isHuidigJaar ? isHuidig : (kw === 1);
    if (isInitieelGeopend) {
      sheet.getRange(headerRij, 12)
        .setValue('← Helemaal klaar met kwartaal? Klik checkbox, graag!')
        .setFontSize(9).setFontStyle('italic').setFontColor('#1e40af')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    } else {
      sheet.getRange(headerRij, 12).clearContent();
    }
    sheet.setRowHeight(rij, isHuidig ? 42 : 26); rij++;

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
    kwartaalRijen.push({ kw: kw, dataStart: kwDataStart, dataEinde: rij - 1, isHuidig: isHuidig });
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
      if (!k.isHuidig) sheet.hideRows(k.dataStart, k.dataEinde - k.dataStart + 1);
    } else {
      sheet.getRange(hdrRij, 8).setValue(k.kw === 1);
      if (k.kw !== 1) sheet.hideRows(k.dataStart, k.dataEinde - k.dataStart + 1);
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

function verversKwartaaloverzichtAlsBestaat_() {
  try {
    var ss = getSS_();
    var bestaatSheet = !!ss.getSheetByName('Kwartaaloverzicht');
    // Debounce: eventuele lopende geplande refresh annuleren
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'verversKwartaaloverzichtGetriggerd') ScriptApp.deleteTrigger(t);
    });
    // Als het tabblad weg is: zo snel mogelijk herbouwen (5s); anders normale vertraging (45s)
    var vertraging = bestaatSheet ? 20000 : 5000;
    ScriptApp.newTrigger('verversKwartaaloverzichtGetriggerd')
      .timeBased().after(vertraging).create();
  } catch(e) { Logger.log('⚠️ Refresh inplannen mislukt: ' + e); }
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

function ruimPersoneelsOudRijenOp_(ss, jaar) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
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
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', totKm:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].totKm += parseFloat(r[7]||0);
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:g.iban,
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
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', aantalRitten:0, totBedrag:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].aantalRitten++;
    map[sleutel].totBedrag += parseFloat(r[9]||0);
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:g.iban,
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
      map[sleutel] = { naam:displayNaam, domein:p.domein||'', rijksreg:p.rijksreg||'', iban:p.iban||'', totKm:0, aantalVerpl:0, totParkeer:0 };
      volgorde.push(sleutel);
    }
    map[sleutel].totKm      += parseFloat(r[8]||0);
    map[sleutel].totParkeer += parseFloat(r[9]||0);
    map[sleutel].aantalVerpl++;
  });
  return volgorde.sort(function(a,b){ return map[a].naam.localeCompare(map[b].naam, 'nl'); }).map(function(s) {
    var g = map[s];
    var parkeerTekst = g.totParkeer > 0 ? '€ ' + g.totParkeer.toFixed(2) : '—';
    return { naam:g.domein, adres:g.naam, pcGem:g.rijksreg, iban:parkeerTekst,
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
  headers.forEach(function(h, k) {
    sheet.getRange(rij, k+1).setValue(h).setFontWeight('bold').setFontSize(8).setFontColor('#475569').setBackground('#e2e8f0').setVerticalAlignment('middle');
  });
  sheet.getRange(rij, 1, 1, 7).setBackground('#e2e8f0');
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
    groepen.forEach(function(g) {
      var tot5 = g.ritten.reduce(function(s,r){ return s+(r.col5||0); }, 0);
      var tot6 = heeftCol6 ? g.ritten.reduce(function(s,r){ return s+(r.col6||0); }, 0) : null;
      grandTot5 += tot5; if (tot6 !== null) grandTot6 += tot6;

      sheet.getRange(rij, 1, 1, 7).setBackground('#ffffff');
      sheet.getRange(rij, 1).setValue(g.naam).setFontWeight('normal').setFontSize(9).setFontColor('#334155');
      sheet.getRange(rij, 2).setValue(g.adres).setBackground('#ffffff').setFontWeight('bold').setFontSize(9).setFontColor('#26295a');
      sheet.getRange(rij, 3).setValue(g.pcGem).setBackground('#ffffff').setFontSize(8).setFontColor('#94a3b8');
      sheet.getRange(rij, 4).setValue(g.ritten[0].col4).setFontSize(9).setFontColor('#334155');
      sheet.getRange(rij, 5).setValue(tot5).setNumberFormat(g.fmt5).setFontWeight('bold').setFontSize(9).setFontColor('#26295a');
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
      sheet.getRange(rij, 7).setValue(g.iban).setFontSize(9).setFontWeight('bold').setFontColor('#26295a');
      sheet.setRowHeight(rij, 22);

      var sleutel = jaar+'|'+kw+'|'+sectie+'|'+g.naam;
      var bestaandeStatus = (statusLookup && statusLookup[sleutel]) ? statusLookup[sleutel] : STATUS_INGEDIEND;
      sheet.getRange(rij, 11).setValue(bestaandeStatus).setDataValidation(statusValidatie)
        .setBackground(statusKleur_(bestaandeStatus)).setFontSize(9).setFontColor('#334155')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      rij++;

      sheet.getRange(rij, 1, 1, 7).setBackground('#ffffff'); sheet.setRowHeight(rij, 6); rij++;
    });
  }
  return rij;
}
