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
