function bouwEmailDomeinMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var email  = (data[i][7]||'').toString().toLowerCase().trim();
    var domein = (data[i][0]||'').toString().trim();
    if (!email || !domein) continue;
    map[email] = domein;
    var p = email.split('@');
    if (p.length === 2) {
      if (p[1] === 'academie-ieper.be') map[p[0] + '@ieper.be']         = domein;
      else if (p[1] === 'ieper.be')     map[p[0] + '@academie-ieper.be'] = domein;
    }
  }
  return map;
}

function bouwEmailNaamMap_(ss) {
  var map   = {};
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var e = (data[i][7]||'').toString().toLowerCase().trim();
    var n = (data[i][1]||'').toString().trim();
    if (e && n) map[e] = n; // laatste rij wint → meest recente naam
  }
  return map;
}

function vindEmailVoorNaam_(ss, naam) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return { email: '', naam: '' };
  var data     = sheet.getDataRange().getValues();
  var gevonden = '';
  var gevondenNaam = '';
  for (var i = 1; i < data.length; i++) {
    // isZelfdePersoon_ vergelijkt woord-ongeordend, dus "Voornaam Familienaam" matcht ook
    // met kolom B ("Familienaam Voornaam") — geen apart omgekeerde-volgorde-vergelijking nodig.
    if (isZelfdePersoon_((data[i][1]||'').toString(), naam)) {
      gevonden     = (data[i][7]||'').toString().trim(); // letterlijk uit kolom H — nooit opgebouwd uit de naam
      gevondenNaam = (data[i][1]||'').toString().trim(); // officiële schrijfwijze, letterlijk uit kolom B
      // geen break: laatste rij wint, zelfde conventie als bouwEmailNaamMap_ hierboven
    }
  }
  return { email: gevonden, naam: gevondenNaam };
}

function sorteerSheet_(sheet, persMap) {
  var lastRow = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  if (lastRow < 2 || numCols < 1) return;

  var range     = sheet.getRange(2, 1, lastRow - 1, numCols);
  var waarden   = range.getValues();
  var formules  = range.getFormulas();
  var notities  = range.getNotes();
  var formaten  = range.getNumberFormats();
  var validaties    = range.getDataValidations();
  var achtergronden = range.getBackgrounds();
  var letterkleurs  = range.getFontColors();

  // Scheid vervolgrijen (woonwerk multi-bewijs, ID begint met '_') van hoofdrijen
  var vervolMap = {}, groepen = {}, volgorde = [];
  for (var i = 0; i < waarden.length; i++) {
    var rid   = (waarden[i][0]||'').toString();
    var email = (waarden[i][1]||'').toString().toLowerCase().trim();
    if (rid.charAt(0) === '_') {
      var m = rid.match(/^_(.+)_(\d+)$/);
      if (m) {
        if (!vervolMap[m[1]]) vervolMap[m[1]] = [];
        vervolMap[m[1]].push({ w: waarden[i], f: formules[i] || [], n: notities[i] || [], fmt: formaten[i] || [], val: validaties[i] || [], bg: achtergronden[i] || [], lk: letterkleurs[i] || [], nr: parseInt(m[2]) });
      }
      continue;
    }
    if (!email) {
      var naamSleutel = '\x00' + ((waarden[i][2]||'').toString().trim().toLowerCase() || String(i));
      if (!groepen[naamSleutel]) { groepen[naamSleutel] = []; volgorde.push(naamSleutel); }
      groepen[naamSleutel].push({ w: waarden[i], f: formules[i] || [], n: notities[i] || [], fmt: formaten[i] || [], val: validaties[i] || [], bg: achtergronden[i] || [], lk: letterkleurs[i] || [] });
      continue;
    }
    if (!groepen[email]) { groepen[email] = []; volgorde.push(email); }
    groepen[email].push({ w: waarden[i], f: formules[i] || [], n: notities[i] || [], fmt: formaten[i] || [], val: validaties[i] || [], bg: achtergronden[i] || [], lk: letterkleurs[i] || [] });
  }

  volgorde.sort(function(a, b) {
    var nA = ((groepen[a][0].w[2])||'').toString().trim() || (persMap[a] || a);
    var nB = ((groepen[b][0].w[2])||'').toString().trim() || (persMap[b] || b);
    return nA.toLowerCase() < nB.toLowerCase() ? -1 : nA.toLowerCase() > nB.toLowerCase() ? 1 : 0;
  });

  volgorde.forEach(function(e) {
    groepen[e].sort(function(a, b) {
      var jaarA = parseInt(a.w[3]) || 0, jaarB = parseInt(b.w[3]) || 0;
      if (jaarA !== jaarB) return jaarA - jaarB;
      return (parseInt(a.w[4]) || 0) - (parseInt(b.w[4]) || 0);
    });
  });

  var nieuweW = [], nieuweF = [], nieuweN = [], nieuweM = [], nieuweV = [], nieuweAg = [], nieuweLk = [];
  volgorde.forEach(function(e) {
    groepen[e].forEach(function(r) {
      nieuweW.push(r.w); nieuweF.push(r.f); nieuweN.push(r.n); nieuweM.push(r.fmt); nieuweV.push(r.val); nieuweAg.push(r.bg); nieuweLk.push(r.lk);
      // Vervolgrijen direct na hun hoofdrij plaatsen
      var pid = (r.w[0]||'').toString();
      if (vervolMap[pid]) {
        vervolMap[pid].sort(function(a, b) { return a.nr - b.nr; });
        vervolMap[pid].forEach(function(vr) { nieuweW.push(vr.w); nieuweF.push(vr.f); nieuweN.push(vr.n); nieuweM.push(vr.fmt); nieuweV.push(vr.val); nieuweAg.push(vr.bg); nieuweLk.push(vr.lk); });
      }
    });
  });
  if (nieuweW.length === 0) return;

  // Fiets en Woon-Werk: kolom 5 (0-gebaseerd) = maand-tekst ("april 2026").
  // Google Sheets kan die string auto-parsen als datum. Zet Date-objecten terug naar
  // tekst en forceer @-opmaak zodat dit na de sort niet opnieuw kan gebeuren.
  var isMaandTab = (sheet.getName() === CONFIG.SHEETS.FIETSVERGOEDING || sheet.getName() === CONFIG.SHEETS.WOON_WERK);
  if (isMaandTab) {
    for (var mi2 = 0; mi2 < nieuweW.length; mi2++) {
      if (nieuweW[mi2][5] instanceof Date) {
        nieuweW[mi2] = nieuweW[mi2].slice();
        nieuweW[mi2][5] = MAANDNAMEN_SERVER_[nieuweW[mi2][5].getMonth()] + ' ' + nieuweW[mi2][5].getFullYear();
      }
      nieuweM[mi2] = nieuweM[mi2].slice();
      nieuweM[mi2][5] = '@';
    }
  }

  range.clearContent();
  range.clearNote();
  sheet.getRange(2, 1, nieuweM.length, numCols).setNumberFormats(nieuweM);
  sheet.getRange(2, 1, nieuweV.length, numCols).setDataValidations(nieuweV);
  sheet.getRange(2, 1, nieuweAg.length, numCols).setBackgrounds(nieuweAg);
  sheet.getRange(2, 1, nieuweLk.length, numCols).setFontColors(nieuweLk);
  sheet.getRange(2, 1, nieuweW.length, numCols).setValues(nieuweW);
  sheet.getRange(2, 1, nieuweN.length, numCols).setNotes(nieuweN);
  for (var r = 0; r < nieuweF.length; r++) {
    for (var k = 0; k < nieuweF[r].length; k++) {
      if (nieuweF[r][k]) sheet.getRange(r + 2, k + 1).setFormula(nieuweF[r][k]);
    }
  }
  var schoonLastRow = sheet.getLastRow();
  if (schoonLastRow > nieuweW.length + 1) {
    sheet.deleteRows(nieuweW.length + 2, schoonLastRow - nieuweW.length - 1);
  }
}


function sorteerPersoneelsSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERSONEELSGEGEVENS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  var range   = sheet.getRange(2, 1, lastRow - 1, numCols);
  var waarden = range.getValues();

  var groepen = {}, volgorde = [], rijZonderEmail = [];
  for (var i = 0; i < waarden.length; i++) {
    var email = (waarden[i][7]||'').toString().toLowerCase().trim();
    if (!email) { rijZonderEmail.push(waarden[i]); continue; }
    if (!groepen[email]) { groepen[email] = { naam: '', rijen: [] }; volgorde.push(email); }
    groepen[email].rijen.push(waarden[i]);
    var geldig = (waarden[i][8]||'').toString().trim().toLowerCase();
    if (!geldig.startsWith('tot')) groepen[email].naam = (waarden[i][1]||'').toString().trim();
  }
  volgorde.forEach(function(e) {
    if (!groepen[e].naam) {
      var r = groepen[e].rijen;
      groepen[e].naam = (r[r.length - 1][1]||'').toString().trim();
    }
  });
  volgorde.sort(function(a, b) {
    var nA = groepen[a].naam.toLowerCase();
    var nB = groepen[b].naam.toLowerCase();
    return nA < nB ? -1 : nA > nB ? 1 : 0;
  });

  var nieuweW = [];
  volgorde.forEach(function(e) { groepen[e].rijen.forEach(function(r) { nieuweW.push(r); }); });
  rijZonderEmail.forEach(function(r) { nieuweW.push(r); });
  if (nieuweW.length === 0) return;

  range.clearContent();
  sheet.getRange(2, 1, nieuweW.length, numCols).setValues(nieuweW);
  var slr = sheet.getLastRow();
  if (slr > nieuweW.length + 1) sheet.deleteRows(nieuweW.length + 2, slr - nieuweW.length - 1);
}
