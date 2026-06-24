// ─── TABBLAD AANPAK ──────────────────────────────────────────────────────────

function maakAanpakTabblad() {
  var ss = getSS_();

  ['Aanpak Nico', 'Aanpak Inse', 'Aanpak Myriam & Nele'].forEach(function(nm) {
    var s = ss.getSheetByName(nm); if (s) ss.deleteSheet(s);
  });

  var sheet = ss.getSheetByName('Aanpak');
  if (!sheet) { sheet = ss.insertSheet('Aanpak', 0); }
  else { sheet.clear(); sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); sheet.clearNotes(); try { sheet.showRows(1, sheet.getMaxRows()); } catch(_) {} }
  sheet.setTabColor('#26295a');

  // Herstel 2-kolomsopmaak (col A: nummer/checkbox 36px, col B: inhoud 660px)
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch(_) {}
  while (sheet.getMaxColumns() < 2) sheet.insertColumnAfter(sheet.getMaxColumns());
  sheet.setColumnWidth(1, 36);
  sheet.setColumnWidth(2, 660);
  if (sheet.getMaxColumns() > 2) sheet.hideColumns(3, sheet.getMaxColumns() - 2);

  var BG = '#f8fafc', BLAUW = '#26295a', TEKST = '#334155';
  var SUBBG = '#f1f5f9', SUBTEKST = '#475569';
  var KNOP_UIT = '#e2e8f0', KNOP_UIT_FC = '#64748b';
  var rij = 1;

  // Rij 1: Titelrij
  sheet.getRange(rij, 1).setBackground(BLAUW);
  sheet.getRange(rij, 2).setValue('Aanpak — Woon-werkverkeer')
    .setBackground(BLAUW).setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(rij, 44); rij++;

  // Rij 2: Ondertitel
  sheet.getRange(rij, 1).setBackground(BG);
  sheet.getRange(rij, 2).setValue('Klik de checkbox naast jouw naam om de instructies te tonen. Na plusminus 4 seconden kun je de instructie lezen.')
    .setBackground(BG).setFontColor('#64748b')
    .setFontSize(9).setFontStyle('italic')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(rij, 22); rij++;

  SpreadsheetApp.flush();

  var tabConfig = [
    { naam: 'Stad Ieper / Cheyana', kleur: '#f59c47' },
    { naam: 'Nico & Inse',          kleur: '#7bc6d3' },
    { naam: 'Myriam & Nele',        kleur: '#d599c4' },
  ];

  var goedkeurItemsNicoInse = [
    'Opgelet: Keur eerst handmatig af als dat nodig is (N + enter = afgekeurd: kleur rood of met het pijltje).',
    'Nico: Klik de checkbox bovenaan de Goedgekeurd?-kolom om in bulk goed te keuren (kleur groen), met uitzondering van het personeel van Beeld.',
    'Inse: Klik de checkbox bovenaan de Goedgekeurd?-kolom om in bulk goed te keuren (kleur groen), enkel voor jezelf én personeel Beeld.',
    'Onderaan verschijnt de melding: "... rijen goedgekeurd voor Kx 20XX".',
    'De checkbox bovenaan springt terug leeg, klaar voor de volgende keer.',
  ];
  var stopItems = [
    'Voor 1ste kwartaal: vanaf 11 januari',
    'Voor 2de kwartaal: vanaf 11 april',
    'Voor 3de kwartaal: vanaf 3 juli (omwille van verlof)',
    'Voor 4de kwartaal: vanaf 11 oktober',
  ];

  var secties = [
    {
      naam: 'Stad Ieper / Cheyana',
      aanhef: 'Beste Cheyana (Stad Ieper), dit ter info:',
      items: [
       { tekst: 'Zoals afgesproken: Er staat een stop op het invullen van de webapp door de leerkrachten:', sub: stopItems },
        { tekst: 'Je krijgt bericht van Myriam Demeester als zij klaar is met controle beleid en personeel Woord, Muziek en Kunstkuur. Je krijgt bericht van Nele Moerman als zij klaar is met controle personeel Beeld.', boldParts: ['Myriam Demeester', 'Nele Moerman'] },
        { tekst: 'Klik het juiste kwartaal aan om in te werken.' },
        { tekst: 'Voer het juiste tarief van fietsvergoeding en dienstverplaatsing per kwartaal manueel aan in kolom H. Als Nico er staat, noteer je daar het hogere tarief fietsvergoeding dat afwijkt van het standaardtarief.' },
        { tekst: 'Standaard staat in kolom K "Ingediend". De status van de verwerking verander je door een klik op het pijltje.' },
        { tekst: 'Bewijsjes kun je in de andere tabbladen hooveren/klikken.' },
        { tekst: 'Is er een nieuw personeelslid? Dan krijg je daar via het Kwartaaloverzicht melding van.' },
        { tekst: 'Is er een wijziging van de gegevens van een bestaand personeelslid? Dan krijg je daar via het Kwartaaloverzicht melding van. In het tabblad Personeelsgegevens zie je welke wijziging er werd doorgevoerd in kolommen I en J. Best "Wijziging doorgevoerd" klikken. Zoniet blijft deze melding zich elk kwartaal herhalen.' },
        { tekst: 'Helemaal klaar met dit kwartaal? Klik in kolom K op de grote checkbox naast de titelbalk van het kwartaal. Hiermee verberg je alle ingevoerde rijen in de andere tabbladen. Dit zorgt voor overzicht voor iedereen. Ter info: wil je ze terughalen? Dan klik je gewoon terug op de grote checkbox.' },
        { tekst: 'Per jaar wordt het tabblad Personeelsgegevens automatisch bijgestuurd en worden de oudere rijen verborgen.' },
        { tekst: 'Als je extra tabblad Kwartaaloverzicht_TEMP ziet verschijnen (en terug verdwijnen), dan is dat omdat er een bijsturing gebeurd is én het tabblad Kwartaaloverzicht op de achtergrond aan het "refreshen" is.' },
      ],
      sluiting: 'Mvg - team de Academie'
    },
    {
      naam: 'Nico & Inse',
      aanhef: 'Beste Nico & Inse,',
      items: [
        { tekst: "Jullie kunnen tussendoor goed/afkeuren, los van Myriam's of Nele's controles op juistheid." },
        { tekst: 'Je opent elk van de tabbladen:', sub: goedkeurItemsNicoInse },
        { tekst: 'Er staat een stop op het invullen van de webapp door de leerkrachten:', sub: stopItems },
        { tekst: 'Na definitieve controle: Graag Myriam / Nele (Beeld) verwittigen, zodat zij weten dat ze definitief kunnen bevestigen aan Cheyana dat zij met de verwerking kan beginnen.' },
        { tekst: 'Uitbetalingen voor het vorige kwartaal worden voorzien eind de eerste maand na afsluiten van het kwartaal (m.u.v. juli = eerder omwille van verlof).' },
        { tekst: 'Als je extra tabblad Kwartaaloverzicht_TEMP ziet verschijnen (en terug verdwijnen), dan is dat omdat er een bijsturing gebeurd is én het tabblad Kwartaaloverzicht op de achtergrond aan het "refreshen" is.' },
      ],
      sluiting: 'Dankjewel, Nico & Inse!'
    },
    {
      naam: 'Myriam & Nele',
      aanhef: 'Beste Myriam & Nele,',
      items: [
        { tekst: 'Jullie kunnen tussendoor controles uitvoeren, ongeacht goed/afkeuring van Nico/Inse. Per rij hebben jullie hiervoor een checkbox om aan te vinken, zodat jullie op die manier "bijhouden" wat jullie al controleerden.' },
        { tekst: 'Nele: de controleboxen voor jouw rijen (Beeld) zijn voor het gemak in groene kleur gezet.' },
        { tekst: 'Opgelet: bij manueel aanpassen van het bedrag wordt alles prima bijgestuurd. Manueel invoeren van extra rijen in een tabblad "Fietsvergoeding, Woon-Werk (trein-De Lijn), Dienstvergoeding" vereist invullen van minstens de kolommen Naam t.e.m. Bestand. Opgelet voor de exacte schrijfwijze van de naam ! (om alles juist te laten overgaan naar het Kwartaaloverzicht). Niet schrikken: de rij springt automatisch in de alfabetische volgorde.',
          hoogte: 90,
          richParts: [
            { tekst: 'Opgelet', bold: true, color: '#dc2626' },
            { tekst: 'Niet schrikken', bold: true, color: '#dc2626' },
            { tekst: 'exacte schrijfwijze van de naam !', bold: true, color: null },
          ]
        },
        { tekst: 'Er staat een stop op het invullen van de webapp door de leerkrachten:', sub: stopItems },
        { tekst: 'Na definitieve controle: Graag Cheyana.Schoer@ieper.be verwittigen, zodat zij vanaf 15 januari, 15 april, 5 juli en 15 oktober aan de slag kan. Je zal normaal via Nico/Inse bericht gekregen hebben dat ze klaar zijn met goedkeuring en/of je zal het zelf zien in de kolom links van die van jullie.' },
        { tekst: 'Uitbetalingen voor het vorige kwartaal worden voorzien eind de eerste maand na afsluiten van het kwartaal (m.u.v. juli = eerder omwille van verlof).' },
        { tekst: 'Als je extra tabblad Kwartaaloverzicht_TEMP ziet verschijnen (en terug verdwijnen), dan is dat omdat er een bijsturing gebeurd is én het tabblad Kwartaaloverzicht op de achtergrond aan het "refreshen" is.' },
      ],
      sluiting: 'Dankjewel, Myriam & Nele!'
    }
  ];

  // Rijen 3-6: Knoppen (alle 4 altijd zichtbaar bovenaan)
  var meta = [];
  tabConfig.forEach(function(tab, ti) {
    var aan = false;
    var bg  = KNOP_UIT;
    var fc  = KNOP_UIT_FC;
    sheet.getRange(rij, 1).insertCheckboxes().setValue(aan).setBackground(bg).setFontColor(fc);
    sheet.getRange(rij, 2).setValue(tab.naam)
      .setBackground(bg).setFontColor(fc)
      .setFontSize(11).setFontWeight(aan ? 'bold' : 'normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(rij, 28);
    meta.push({ naam: tab.naam, kleur: tab.kleur, buttonRij: rij, start: 0, eind: 0 });
    rij++;
  });

  // Scheiding knoppen / content
  sheet.getRange(rij, 1, 1, 2).setBackground('#e2e8f0'); sheet.setRowHeight(rij, 6); rij++;

  SpreadsheetApp.flush();

  secties.forEach(function(sectie, si) {
    var tabConf = tabConfig[si];

    // Aanhef
    var sectieStart = rij;
    sheet.getRange(rij, 1).setBackground(BG);
    sheet.getRange(rij, 2).setValue(sectie.aanhef)
      .setBackground(BG).setFontColor(tabConf.kleur)
      .setFontSize(10).setFontWeight('bold').setFontStyle('italic')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(rij, 32); rij++;

    // Items
    var itemVals = [], colABG = [], colBBG = [], colAFC = [], colBFC = [], rijHoogtes = [];
    var itemStartRij = rij;
    var boldRijen = []; // { offset, tekst, parts } voor items met boldParts
    var flatOffset = 0;

    sectie.items.forEach(function(item, idx) {
      itemVals.push([(idx + 1) + '.', item.tekst]);
      colABG.push([BG]);    colBBG.push([BG]);
      colAFC.push([BLAUW]); colBFC.push([TEKST]);
      rijHoogtes.push(item.hoogte || (item.sub ? 30 : 38));
      if (item.boldParts || item.richParts) boldRijen.push({ offset: flatOffset, tekst: item.tekst, parts: item.boldParts, richParts: item.richParts });
      flatOffset++;
      if (item.sub) {
        item.sub.forEach(function(subTekst) {
          itemVals.push(['', String.fromCharCode(8594) + ' ' + subTekst]);
          colABG.push([SUBBG]);  colBBG.push([SUBBG]);
          colAFC.push([BLAUW]);  colBFC.push([SUBTEKST]);
          rijHoogtes.push(30);
          flatOffset++;
        });
        itemVals.push(['', '']);
        colABG.push([BG]); colBBG.push([BG]);
        colAFC.push(['#000000']); colBFC.push(['#000000']);
        rijHoogtes.push(8);
        flatOffset++;
      }
    });

    var nItems = itemVals.length;
    if (nItems > 0) {
      sheet.getRange(itemStartRij, 1, nItems, 2).setValues(itemVals);
      sheet.getRange(itemStartRij, 1, nItems, 1).setBackgrounds(colABG).setFontColors(colAFC)
        .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('top');
      sheet.getRange(itemStartRij, 2, nItems, 1).setBackgrounds(colBBG).setFontColors(colBFC)
        .setFontSize(9).setFontWeight('normal').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      for (var h = 0; h < nItems; h++) sheet.setRowHeight(itemStartRij + h, rijHoogtes[h]);
      // Gedeeltelijk vet voor items met boldParts
      var normStijl = SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(TEKST).setBold(false).build();
      var vetStijl  = SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(TEKST).setBold(true).build();
      boldRijen.forEach(function(br) {
        var rtb = SpreadsheetApp.newRichTextValue().setText(br.tekst).setTextStyle(0, br.tekst.length, normStijl);
        if (br.richParts) {
          br.richParts.forEach(function(rp) {
            var kleur = rp.color || TEKST;
            var stijl = SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(kleur).setBold(rp.bold !== false).build();
            var pos = br.tekst.indexOf(rp.tekst);
            while (pos !== -1) { rtb.setTextStyle(pos, pos + rp.tekst.length, stijl); pos = br.tekst.indexOf(rp.tekst, pos + 1); }
          });
        }
        if (br.parts) {
          br.parts.forEach(function(part) {
            var pos = br.tekst.indexOf(part);
            while (pos !== -1) { rtb.setTextStyle(pos, pos + part.length, vetStijl); pos = br.tekst.indexOf(part, pos + 1); }
          });
        }
        sheet.getRange(itemStartRij + br.offset, 2).setRichTextValue(rtb.build());
      });
    }
    rij = itemStartRij + nItems;

    // Sluiting
    sheet.getRange(rij, 1, 1, 2).setBackground(BG); sheet.setRowHeight(rij, 10); rij++;
    sheet.getRange(rij, 1).setBackground(BG);
    sheet.getRange(rij, 2).setValue(sectie.sluiting)
      .setBackground(BG).setFontColor(tabConf.kleur)
      .setFontSize(9).setFontWeight('bold').setFontStyle('italic')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(rij, 28); rij++;

    var sectieEind = rij - 1;
    meta[si].start = sectieStart;
    meta[si].eind  = sectieEind;

    // Alle secties standaard verborgen
    sheet.hideRows(sectieStart, sectieEind - sectieStart + 1);

    SpreadsheetApp.flush();
  });

  // Meta opslaan als note op A1
  sheet.getRange(1, 1).setNote(JSON.stringify(meta));

  sheet.getRange(1, 1, rij - 1, 2)
    .setBorder(true, true, true, true, false, false, BLAUW, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var maxRij = sheet.getMaxRows();
  if (maxRij >= rij) sheet.deleteRows(rij, maxRij - rij + 1);

  Logger.log('Aanpak-tabblad aangemaakt.');
  try { SpreadsheetApp.getUi().alert('Aanpak-tabblad aangemaakt.'); } catch(_) {}
}

function toonAanpakSectie_(sheet, keuze, meta) {
  if (!meta) {
    var note = sheet.getRange(1, 1).getNote();
    if (!note) return;
    try { meta = JSON.parse(note); } catch(_) { return; }
  }
  var KNOP_UIT = '#e2e8f0', KNOP_UIT_FC = '#64748b', KNOP_AAN_FC = '#26295a';
  var firstRow = meta[0].buttonRij;
  var n = meta.length;
  var vals = [], bgs2 = [], fcs2 = [], weights = [];
  var selStart = 0, selEind = 0;
  meta.forEach(function(tab) {
    var aan = (tab.naam === keuze);
    var bg  = aan ? tab.kleur : KNOP_UIT;
    var fc  = aan ? KNOP_AAN_FC : KNOP_UIT_FC;
    vals.push([aan]);
    bgs2.push([bg, bg]);
    fcs2.push([fc, fc]);
    weights.push([aan ? 'bold' : 'normal']);
    if (aan) { selStart = tab.start; selEind = tab.eind; }
  });
  // Knoppen: 4 API-calls (vals, backgrounds A+B samen, fontcolors A+B samen, weights)
  sheet.getRange(firstRow, 1, n, 1).setValues(vals);
  sheet.getRange(firstRow, 1, n, 2).setBackgrounds(bgs2).setFontColors(fcs2);
  sheet.getRange(firstRow, 2, n, 1).setFontWeights(weights);
  // Content: 2 API-calls (alles verbergen, dan enkel de gekozen sectie tonen)
  var allStart = meta[0].start;
  var allEind  = meta[n - 1].eind;
  if (allEind >= allStart) {
    sheet.hideRows(allStart, allEind - allStart + 1);
    if (selStart > 0 && selEind >= selStart) sheet.showRows(selStart, selEind - selStart + 1);
  }
}
