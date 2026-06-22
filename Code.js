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



// ─── STATUS-WAARDEN ──────────────────────────────────────────────────────────
var STATUS_INGEDIEND = 'Ingediend';
var STATUS_CONTROLE  = 'Controle uitgevoerd';
var STATUS_BETALING  = 'Betaling verwerkt';

function statusKleur_(status) {
  if (status === STATUS_BETALING) return '#bbf7d0';
  if (status === STATUS_CONTROLE) return '#fef9c3';
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
    .addItem('Automatische triggers installeren', 'installeerBenodigdeTriggers')
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


function sheetNaamVoorCategorie_(categorie) {
  return { fiets: CONFIG.SHEETS.FIETSVERGOEDING, woonwerk: CONFIG.SHEETS.WOON_WERK, dienst: CONFIG.SHEETS.DIENSTVERPLAATSING }[categorie];
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






// ─── TESTDATUM BEHEER (admin UI) ─────────────────────────────────────────────

function getAdminTestDate() {
  return { testDate: getConfig('TEST_DATE') || '', isAdmin: isAdminEmail(getEmail_()) };
}

function saveAdminTestDate(dateStr) {
  if (!isAdminEmail(getEmail_())) return { ok: false, error: 'Geen beheerdersrechten.' };
  setConfig('TEST_DATE', dateStr || '');
  return { ok: true };
}
