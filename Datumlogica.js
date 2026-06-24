// ─── DATUM / KWARTAAL LOGICA ─────────────────────────────────────────────────

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
    return new Date(year, 6, 3, 23, 59, 59);  // Uiterste indieningstermijn Q2 = altijd 3 juli
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
