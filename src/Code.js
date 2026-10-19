/**
 * ZOOM – Adatbázis (csak iroda)
 * Csak OLVAS a forrástáblákból, és egy külön, privát munkafüzetbe ír.
 */
var CFG = {
  KFT_ID: '1AJZGLcOlHerO0gV3rY3jaKa2hfl-gQcVzTMjJ1nmmBk', KFT_GID: 575676776,
  EV_ID: '1bU6oi6bL325iEe6Z0ktzU_sYIPcQ7efqV2DSME9YyvE', EV_GID: 673679712,
  TERV_ID: '1glT71zO4viHp8yVaJFXRODREg-B9XLMCXIA6YFaTaho', NAPLO_GID: 1465640060, FUGGO_GID: 48947702,
  DB_NAME: 'ZOOM Adatbázis (csak iroda)'
};

var TAN_FEJ = ['Név', 'Cég', 'Képzés', 'Státusz', 'Teendő', 'Fizetendő (Ft)', 'Befizetett (Ft)', 'Hátralék (Ft)', 'Oktató', 'Vezetett óra', 'Befizetett óra', 'Fedezet (óra)',
  'Telefon', 'E-mail', 'Kezdés', 'Szerződés kelte', 'Aláírva', 'Szül. idő', 'Azonosító', 'KAV azonosító', 'Csoport', 'Kedvezmény (Ft)', 'Ár (Ft)', 'Fizetés módja',
  'Vizsgadíj befizetve (Ft)', 'Függő számla (Ft, ellenőrizendő)', 'Kategória (oktatói tábla)', 'Óradíj (Ft)', 'Forrás', 'Megjegyzés',
  'Főcsoport', 'Alcsoport', 'Költségviselő', 'Megrendelő cég', 'Pótóra fizetője',
  'GKI vizsgák kész', 'GKI kártya igényelve', 'GKI kártya megérkezett', 'GKI kártya postázása',
  'Elméleti képzés formája', 'Céges befizetés (Ft)', 'Céges befizetés dátuma'];
var BEF_FEJ = ['Időpont', 'Számlaszám', 'Sorozat', 'Forrás', 'Név a számlán', 'Összeg (Ft)', 'Vizsgadíj (Ft)', 'Tanuló azonosító',
  'Tanuló (adatbázis)', 'Oktató', 'Állapot', 'Megjegyzés'];

function szinkron() {
  var t0 = new Date();
  var db = getDb_();
  var enr = readEnroll_(CFG.KFT_ID, CFG.KFT_GID, 'Kft').concat(readEnroll_(CFG.EV_ID, CFG.EV_GID, 'E.V.'));
  var CSMAP = csoportTerkep_();
  enr.forEach(function (e) {
    var m2 = CSMAP[String(e.kepzes).trim()];
    e.fcs = m2 ? m2[0] : ''; e.alcs = m2 ? m2[1] : '';
  });
  var ex = handleDeletions_(db);
  enr = enr.filter(function (e) { return !ex[e.id]; });
  var terv = SpreadsheetApp.openById(CFG.TERV_ID);
  var leg = readLegacy_(terv);
  var pays = readPayments_(terv);

  var students = [];
  var legUsed = {};
  enr.forEach(function (e) {
    e.key = key_(e.nev);
    var m = leg.filter(function (l) { return l.key === e.key; });
    if (m.length) { e.leg = m[0]; m.forEach(function (l) { legUsed[l.uid] = e.id; }); if (m.length > 1) e.megj.push('Több oktatói tábla sor egyezik a névvel'); }
    students.push(e);
  });
  students.forEach(function (s) { if (!s.forras) s.forras = 'Beiratkozás'; });
  var byKey = {};
  students.forEach(function (s) { (byKey[s.key] = byKey[s.key] || []).push(s); });
  var legIdByUid = {};
  students.forEach(function (s) { if (s.leg) legIdByUid[s.leg.uid] = s.id; });
  leg.forEach(function (l) { if (legUsed[l.uid]) legIdByUid[l.uid] = legUsed[l.uid]; });

  pays.forEach(function (p) {
    p.sid = ''; p.snev = '';
    if (p.oktato) {
      var cand = leg.filter(function (l) { return l.key === p.key && l.oktato === p.oktato; });
      if (cand.length) { p.sid = legIdByUid[cand[0].uid]; }
    }
    if (!p.sid) {
      var c = byKey[p.key] || [];
      if (c.length === 1) p.sid = c[0].id;
      else if (c.length > 1) p.megj = (p.megj ? p.megj + ' | ' : '') + 'Több egyező nevű tanuló (' + c.length + ')';
    }
  });
  var sidIdx = {}; students.forEach(function (s) { sidIdx[s.id] = s; });
  pays.forEach(function (p) { if (p.sid && sidIdx[p.sid]) p.snev = sidIdx[p.sid].nev; });
  var nemParos = pays.filter(function (p) { return !p.sid; });
  pays = pays.filter(function (p) { return p.sid; });

  var sum = {};
  pays.forEach(function (p) {
    if (!p.sid) return;
    var a = sum[p.sid] = sum[p.sid] || { naplo: 0, fuggo: 0, vizsga: 0 };
    if (p.tipus === 'naplo') a.naplo += p.osszeg; else a.fuggo += p.osszeg;
    a.vizsga += p.vizsga;
  });

  var rows = students.map(function (s) {
    var a = sum[s.id] || { naplo: 0, fuggo: 0, vizsga: 0 };
    var l = s.leg;
    var fizetendo = (s.ar !== '' && s.ar !== null) ? s.ar - (s.ked || 0) : '';
    var befiz = (l ? l.befiz : a.naplo + a.fuggo) + (s.cegBef || 0);
    var hatralek = fizetendo === '' ? '' : fizetendo - befiz;
    var befOra = l ? l.befOra : '';
    var fedezet = l ? round2_(l.befOra - l.vezetett) : '';
    return [s.nev, s.ceg, s.kepzes, s.statusz, s.teendo, fizetendo, befiz, hatralek, l ? l.oktato : '', l ? l.vezetett : '', befOra, fedezet,
      s.tel, s.email, s.kezd, s.szerz, s.alairva, s.szul, s.id, s.kav, s.csoport, s.ked === 0 ? '' : s.ked, s.ar, s.fizmod,
      a.vizsga || '', a.fuggo || '', l ? l.kat : '', l ? l.dij : '', s.forras, (s.megj || []).join(' | '),
      s.fcs || '', s.alcs || '', s.kviselo || '', s.mceg || '', s.potora || '',
      s.gkiKesz || '', s.gkiIg || '', s.gkiErk || '', s.gkiPosta || '', s.elmForma || '', s.cegBef || '', s.cegBefD || ''];
  });
  var oldT = db.getSheetByName('Tanulók');
  if (oldT && oldT.getRange(1, 1).getValue() === 'Töröl') oldT.deleteColumn(1);
  var shT = writeSheet_(db, 'Tanulók', TAN_FEJ, rows);
  var nR = Math.max(rows.length, 1);
  [6, 7, 8, 22, 23, 25, 26, 28, 41].forEach(function (c) { shT.getRange(2, c, nR, 1).setNumberFormat('#,##0'); });
  [15, 16, 18, 36, 37, 38, 42].forEach(function (c) { shT.getRange(2, c, nR, 1).setNumberFormat('yyyy.mm.dd.'); });
  var W = [210, 70, 270, 90, 250, 110, 110, 110, 170, 95, 95, 95, 140, 210, 100, 110, 80, 100, 90, 110, 200, 110, 110, 200, 120, 130, 130, 90, 110, 260,
    120, 180, 110, 230, 120, 120, 130, 140, 140, 180, 140, 140];
  W.forEach(function (w, i) { shT.setColumnWidth(i + 1, w); });
  var rg = function (c) { return shT.getRange(2, c, nR, 1); };
  var R = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  shT.setConditionalFormatRules([
    R().whenTextEqualTo('Kft').setBackground('#dbeafe').setFontColor('#1e40af').setRanges([rg(2)]).build(),
    R().whenTextEqualTo('E.V.').setBackground('#ffedd5').setFontColor('#9a3412').setRanges([rg(2)]).build(),
    R().whenTextEqualTo('Aktív').setBackground('#dcfce7').setFontColor('#166534').setRanges([rg(4)]).build(),
    R().whenTextEqualTo('Lezárva').setBackground('#e5e7eb').setFontColor('#374151').setRanges([rg(4)]).build(),
    R().whenCellNotEmpty().setBackground('#fef9c3').setFontColor('#854d0e').setRanges([rg(5)]).build(),
    R().whenNumberGreaterThan(0).setBackground('#ffedd5').setFontColor('#9a3412').setRanges([rg(8)]).build(),
    R().whenNumberLessThan(0).setBackground('#fee2e2').setFontColor('#991b1b').setRanges([rg(12)]).build()
  ]);
  shT.getRange(2, 2, nR, 1).setHorizontalAlignment('center');
  shT.getRange(2, 4, nR, 1).setHorizontalAlignment('center');
  shT.setFrozenColumns(1);
  try { if (shT.getColumnGroupDepth(19) === 0) { shT.getRange(1, 19, 1, 12).shiftColumnGroupDepth(1); } shT.collapseAllColumnGroups(); } catch (e) {}
  shT.insertColumnBefore(1);
  shT.getRange(1, 1).setValue('Töröl');
  if (rows.length) shT.getRange(2, 1, rows.length, 1).insertCheckboxes();
  shT.getRange(2, 1, nR, 1).setHorizontalAlignment('center').setBackground('#ffffff');
  shT.setColumnWidth(1, 60);
  shT.setFrozenColumns(2);
  if (shT.getFilter()) shT.getFilter().remove();
  shT.getRange(1, 1, Math.max(rows.length + 1, 2), shT.getLastColumn()).createFilter();

  var prow = pays.map(function (p) {
    return [p.ido, p.szamla, p.sorozat, p.forras, p.nev, p.osszeg, p.vizsga, p.sid, p.snev, p.oktato, p.allapot, p.megj];
  });
  prow.sort(function (x, y) { return (y[0] instanceof Date ? y[0].getTime() : 0) - (x[0] instanceof Date ? x[0].getTime() : 0); });
  writeSheet_(db, 'Befizetések', BEF_FEJ, prow);
  var nP = Math.max(prow.length, 1);
  db.getSheetByName('Befizetések').getRange(2, 6, nP, 2).setNumberFormat('#,##0');
  db.getSheetByName('Befizetések').getRange(2, 1, nP, 1).setNumberFormat('yyyy.mm.dd. hh:mm');

  var ok = {};
  leg.filter(function (l) { return legUsed[l.uid]; }).forEach(function (l) {
    var o = ok[l.oktato] = ok[l.oktato] || { db: 0, vez: 0, bef: 0, befFt: 0, hianyDb: 0, hianyOra: 0 };
    o.db++; o.vez += l.vezetett; o.bef += l.befOra; o.befFt += l.befiz;
    var f = l.befOra - l.vezetett;
    if (f < 0) { o.hianyDb++; o.hianyOra += -f; }
  });
  var orow = Object.keys(ok).sort().map(function (k) {
    var o = ok[k]; return [k, o.db, round2_(o.vez), round2_(o.bef), round2_(o.bef - o.vez), o.hianyDb, round2_(o.hianyOra), o.befFt];
  });
  writeSheet_(db, 'Oktatók', ['Oktató', 'Tanulók száma', 'Vezetett óra', 'Befizetett óra', 'Fedezet (óra)', 'Fedezet nélküli tanulók', 'Hiányzó fedezet (óra)', 'Befizetések összesen (Ft)'], orow);
  db.getSheetByName('Oktatók').getRange(2, 8, Math.max(orow.length, 1), 1).setNumberFormat('#,##0');

  var k = {
    total: students.length,
    kft: students.filter(function (s) { return s.ceg === 'Kft'; }).length,
    ev: students.filter(function (s) { return s.ceg === 'E.V.'; }).length,
    aktiv: students.filter(function (s) { return s.statusz === 'Aktív'; }).length,
    lezart: students.filter(function (s) { return s.statusz === 'Lezárva'; }).length,
    fedezetNelkul: rows.filter(function (r) { return r[11] !== '' && r[11] < 0; }).length,
    hatralek: rows.reduce(function (a, r) { return a + (typeof r[7] === 'number' && r[7] > 0 && String(r[29]).indexOf('Teszt') < 0 ? r[7] : 0); }, 0),
    befDb: pays.length,
    befOsszeg: pays.reduce(function (a, p) { return a + p.osszeg; }, 0),
    nemParos: nemParos.length,
    utolso: Utilities.formatDate(new Date(), 'Europe/Budapest', 'yyyy.MM.dd. HH:mm')
  };
  dashboard_(db, k);
  ensureGuide_(db);
  var nap = db.getSheetByName('Frissítés napló') || db.insertSheet('Frissítés napló');
  if (nap.getLastRow() === 0) nap.appendRow(['Időpont', 'Tanulók', 'Oktatói sorok', 'Befizetések', 'Futási idő (mp)', 'Megjegyzés']);
  nap.appendRow([new Date(), students.length, leg.length, pays.length, Math.round((new Date() - t0) / 1000), '']);
  var sorted = ['Összesítő', 'Tanulók', 'Befizetések', 'Oktatók', 'Kizárt', 'Útmutató', 'Frissítés napló'];
  sorted.forEach(function (n, i) { var s = db.getSheetByName(n); if (s) { db.setActiveSheet(s); db.moveActiveSheet(i + 1); } });
  var d = db.getSheetByName('Munkalap1') || db.getSheetByName('Sheet1'); if (d) db.deleteSheet(d);
  saveState_(db, students);
  ensureEditTrigger_(db);
  Logger.log('Kész. DB: ' + db.getUrl() + ' | tanulók: ' + students.length + ' | oktatói sor: ' + leg.length + ' | befizetés: ' + pays.length);
}

function napiIdozito() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'szinkron') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('szinkron').timeBased().everyDays(1).atHour(7).inTimezone('Europe/Budapest').create();
}

function readEnroll_(id, gid, ceg) {
  var sh = sheetByGid_(SpreadsheetApp.openById(id), gid);
  var v = sh.getDataRange().getValues();
  var h = {};
  v[0].forEach(function (x, i) { h[String(x).trim()] = i; });
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    var g = function (n) { var j = h[n]; return j === undefined ? '' : r[j]; };
    var nev = String(g('Név')).trim();
    if (!nev || !g('Sorszám')) continue;
    if (/^minta\b/i.test(nev)) continue;
    var megj = [];
    if (/^(próba|proba|teszt)/i.test(nev)) megj.push('Teszt adat?');
    var lez = g('Lezárva (dátum)');
    var ar = num_(g('összeg')), ked = num_(g('Kedvezmény (Ft)'));
    out.push({
      id: String(g('Sorszám')).trim(), ceg: ceg, nev: nev, szul: date_(g('Születési idő')), kepzes: g('Képzés'), csoport: g('csoport'),
      tel: g('Telefon'), email: g('E-mail'), szerz: date_(g('Szerződés kelte')), kezd: date_(g('Kezdés')), alairva: g('Aláírva') ? 'igen' : 'nem',
      teendo: g('TEENDŐ'), statusz: lez ? 'Lezárva' : 'Aktív', kav: g('KAV tanulóazonosító'), ar: ar || '', ked: ked, fizmod: g('Fizetés módja'),
      megj: megj,
      kviselo: g('Költségviselő'), mceg: g('Cég'), potora: g('Pótóra fizetője'),
      gkiKesz: date_(g('GKI vizsgák kész')), gkiIg: date_(g('GKI kártya igényelve')),
      gkiErk: date_(g('GKI kártya megérkezett')), gkiPosta: g('GKI kártya postázása'),
      elmForma: g('Elméleti képzés formája'),
      cegBef: num_(g('Céges befizetés (Ft)')), cegBefD: date_(g('Céges befizetés dátuma'))
    });
  }
  return out;
}

function readLegacy_(terv) {
  var out = [];
  terv.getSheets().forEach(function (sh) {
    var v = sh.getDataRange().getValues();
    var oktato = '';
    for (var i = 0; i < Math.min(4, v.length); i++) {
      var mm = String(v[i][0] || '').match(/Gyakorlati oktató neve:\s*([\s\S]+)/);
      if (mm) { oktato = mm[1].replace(/Sorszám/g, '').replace(/\s+/g, ' ').trim(); break; }
    }
    if (!oktato) return;
    for (var j = 1; j < v.length; j++) {
      var r = v[j];
      if (!(typeof r[0] === 'number' || /^\d+$/.test(String(r[0]).trim()))) continue;
      var nev = String(r[3] || '').trim();
      if (!nev || /vizsg|összes|átlag|idei|kiállított/i.test(nev)) continue;
      if (r[1] === '' && num_(r[5]) === 0) continue;
      if (/^\d+$/.test(nev)) continue;
      var k = key_(nev);
      out.push({
        oktato: oktato, tab: sh.getName(), sor: j + 1, nev: nev, key: k, kat: r[1], athozott: num_(r[2]), dij: num_(r[5]), befiz: num_(r[7]),
        egyenleg: num_(r[8]), vezetett: num_(r[9]), befOra: round2_(num_(r[10])),
        uid: 'R-' + hash_(k + '|' + oktato)
      });
    }
  });
  return out;
}

function readPayments_(terv) {
  var out = [], seen = {};
  var nv = sheetByGid_(terv, CFG.NAPLO_GID).getDataRange().getValues();
  for (var i = 1; i < nv.length; i++) {
    var r = nv[i];
    if (!r[3] && !r[2]) continue;
    var p = { tipus: 'naplo', ido: r[0], oktato: String(r[1] || '').trim(), nev: String(r[2] || '').trim(), szamla: String(r[3] || '').trim(),
      osszeg: num_(r[4]), vizsga: num_(r[8]), forras: r[9] || '', allapot: 'Könyvelve az oktatói táblába', megj: '' };
    p.key = key_(p.nev); p.sorozat = prefix_(p.szamla);
    seen[p.szamla + '|' + p.key + '|' + p.osszeg] = 1;
    out.push(p);
  }
  var fv = sheetByGid_(terv, CFG.FUGGO_GID).getDataRange().getValues();
  for (var j = 1; j < fv.length; j++) {
    var f = fv[j];
    if (!f[2] && !f[3]) continue;
    var nev = String(f[7] || f[3] || '').trim();
    var q = { tipus: 'fuggo', ido: f[0], oktato: String(f[6] || '').trim(), nev: nev, szamla: String(f[2] || '').trim(), osszeg: num_(f[4]), vizsga: num_(f[5]),
      forras: f[1] || '', allapot: f[8] || '', megj: f[9] || '' };
    q.key = key_(q.nev); q.sorozat = prefix_(q.szamla);
    if (seen[q.szamla + '|' + q.key + '|' + q.osszeg]) continue;
    out.push(q);
  }
  out.forEach(function (p) { p.ido = dateTime_(p.ido); });
  return out;
}

function getDb_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DB_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss = SpreadsheetApp.create(CFG.DB_NAME);
  props.setProperty('DB_ID', ss.getId());
  return ss;
}
function sheetByGid_(ss, gid) {
  var s = ss.getSheets().filter(function (x) { return x.getSheetId() === gid; })[0];
  if (!s) throw new Error('Nincs ilyen lap (gid ' + gid + '): ' + ss.getName());
  return s;
}
function writeSheet_(db, name, head, rows) {
  var sh = db.getSheetByName(name) || db.insertSheet(name);
  sh.clear(); sh.setConditionalFormatRules([]);
  sh.getBandings().forEach(function (b) { b.remove(); });
  if (sh.getFilter()) sh.getFilter().remove();
  var w = head.length, n = rows.length;
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  if (sh.getMaxRows() < n + 2) sh.insertRowsAfter(sh.getMaxRows(), n + 2 - sh.getMaxRows());
  sh.getRange(1, 1, n + 1, w).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, w).setValues([head]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff').setWrap(true).setHorizontalAlignment('left');
  if (n) {
    sh.getRange(2, 1, n, w).setValues(rows);
    sh.getRange(2, 1, n, w).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    sh.setRowHeights(2, n, 30);
  }
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(n + 1, 2), w).createFilter();
  sh.setHiddenGridlines(true);
  sh.setTabColor('#0f172a');
  for (var c = 1; c <= w; c++) sh.setColumnWidth(c, Math.max(100, Math.min(220, String(head[c - 1]).length * 9 + 30)));
  return sh;
}
function dashboard_(db, k) {
  var sh = db.getSheetByName('Összesítő') || db.insertSheet('Összesítő');
  sh.clear(); sh.setConditionalFormatRules([]);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.setHiddenGridlines(true); sh.setTabColor('#2563eb');
  if (sh.getMaxColumns() < 9) sh.insertColumnsAfter(sh.getMaxColumns(), 9 - sh.getMaxColumns());
  sh.getRange(1, 1, 16, 9).setBackground('#f1f5f9').setFontFamily('Roboto').setVerticalAlignment('middle');
  [20, 230, 20, 230, 20, 230, 20, 230, 20].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  [16, 44, 24, 16, 26, 52, 24, 16, 26, 52, 24].forEach(function (h, i) { sh.setRowHeight(i + 1, h); });
  sh.getRange(2, 2).setValue('ZOOM Adatbázis').setFontSize(22).setFontWeight('bold').setFontColor('#0f172a');
  sh.getRange(3, 2).setValue('Csak iroda · beiratkozott tanulók mindkét cégtől (Kft és E.V.)').setFontSize(10).setFontColor('#64748b');
  var card = function (row, col, label, value, note, fmt) {
    var box = sh.getRange(row, col, 3, 1);
    box.setBackground('#ffffff').setBorder(true, true, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, col).setValue(label).setFontSize(10).setFontWeight('bold').setFontColor('#64748b').setHorizontalAlignment('left');
    var v = sh.getRange(row + 1, col).setValue(value).setFontSize(26).setFontWeight('bold').setFontColor('#0f172a').setHorizontalAlignment('left');
    if (fmt) v.setNumberFormat(fmt);
    sh.getRange(row + 2, col).setValue(note).setFontSize(9).setFontColor('#94a3b8').setHorizontalAlignment('left');
  };
  card(5, 2, 'TANULÓK ÖSSZESEN', k.total, 'Kft: ' + k.kft + ' · E.V.: ' + k.ev, '#,##0');
  card(5, 4, 'AKTÍV TANULÓK', k.aktiv, 'Lezárt: ' + k.lezart, '#,##0');
  card(5, 6, 'FEDEZET NÉLKÜLI TANULÓK', k.fedezetNelkul, 'több órát vezetett, mint amennyit befizetett', '#,##0');
  card(5, 8, 'HÁTRALÉK ÖSSZESEN (FT)', k.hatralek, 'a beiratkozási adatok alapján, teszt sorok nélkül', '#,##0');
  card(9, 2, 'BEFIZETÉSEK (DB)', k.befDb, 'beiratkozott tanulókhoz kötve', '#,##0');
  card(9, 4, 'BEFIZETÉSEK ÖSSZEGE (FT)', k.befOsszeg, 'számlázz.hu alapján', '#,##0');
  card(9, 6, 'NEM PÁROSÍTOTT SZÁMLA', k.nemParos, 'nem beiratkozott tanulóé, nincs tárolva', '#,##0');
  card(9, 8, 'UTOLSÓ FRISSÍTÉS', k.utolso, 'naponta 6 és 7 óra között', null);
  sh.getRange(10, 8).setFontSize(18);
}
function onTanuloEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== 'Tanulók' || e.range.getColumn() !== 1 || e.range.getRow() < 2) return;
    if (String(e.value).toUpperCase() !== 'TRUE') return;
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var idCol = head.indexOf('Azonosító') + 1, nameCol = head.indexOf('Név') + 1;
    var row = e.range.getRow();
    var id = sh.getRange(row, idCol).getValue(), nev = sh.getRange(row, nameCol).getValue();
    var kiz = e.source.getSheetByName('Kizárt');
    if (kiz && id) kiz.appendRow([id, nev, new Date()]);
    var st = e.source.getSheetByName('_állapot');
    if (st && st.getLastRow() > 0) {
      var ids = st.getRange(1, 1, st.getLastRow(), 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(id)) { st.deleteRow(i + 1); break; } }
    }
    sh.deleteRow(row);
  } catch (err) { Logger.log('onTanuloEdit: ' + err); }
}
function ensureEditTrigger_(db) {
  try {
    var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'onTanuloEdit'; });
    if (!has) ScriptApp.newTrigger('onTanuloEdit').forSpreadsheet(db).onEdit().create();
  } catch (err) { Logger.log('Trigger hiba: ' + err); }
}
function handleDeletions_(db) {
  var last = loadState_(db);
  var kiz = db.getSheetByName('Kizárt') || db.insertSheet('Kizárt');
  if (kiz.getLastRow() === 0) {
    kiz.getRange(1, 1, 1, 3).setValues([['Azonosító', 'Név', 'Kizárva']]).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    kiz.setFrozenRows(1); kiz.setColumnWidth(1, 120); kiz.setColumnWidth(2, 260); kiz.setColumnWidth(3, 160); kiz.setTabColor('#94a3b8');
  }
  var ex = {};
  if (kiz.getLastRow() > 1) kiz.getRange(2, 1, kiz.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) ex[String(r[0])] = 1; });
  var shT = db.getSheetByName('Tanulók');
  if (shT && shT.getLastRow() > 1 && Object.keys(last).length) {
    var cur = {};
    shT.getRange(2, shT.getRange(1, 1, 1, shT.getLastColumn()).getValues()[0].indexOf('Azonosító') + 1, shT.getLastRow() - 1, 1).getValues().forEach(function (r) { cur[String(r[0])] = 1; });
    Object.keys(last).forEach(function (id) {
      if (!cur[id] && !ex[id]) { kiz.appendRow([id, last[id], new Date()]); ex[id] = 1; }
    });
  }
  return ex;
}
function loadState_(db) {
  var sh = db.getSheetByName('_állapot'), m = {};
  if (sh && sh.getLastRow() > 0) sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) { if (r[0]) m[String(r[0])] = r[1]; });
  return m;
}
function saveState_(db, students) {
  var sh = db.getSheetByName('_állapot') || db.insertSheet('_állapot');
  sh.clear();
  var v = students.map(function (s) { return [s.id, s.nev]; });
  if (v.length) sh.getRange(1, 1, v.length, 2).setValues(v);
  sh.hideSheet();
}
function ensureGuide_(db) {
  var sh = db.getSheetByName('Útmutató') || db.insertSheet('Útmutató');
  sh.clear();
  var t = [
    ['ZOOM Adatbázis – csak iroda'],
    [''],
    ['Ez a munkafüzet automatikusan frissül (naponta reggel). A lapokat NE szerkeszd kézzel, mert a következő frissítés felülírja őket.'],
    ['Források (csak olvasás): Kft és E.V. beiratkozási tábla (Tanulók lap), egységes tábla (oktatói fülek, Napló, Függőben). Adrienn kézi táblájához a program nem nyúl.'],
    [''],
    ['Lapok:'],
    ['Összesítő – néhány fő szám egy helyen.'],
    ['Tanulók – egy sor = egy tanuló (mindkét cég). Beiratkozási adatok + befizetés + oktatói tábla adatai (vezetett óra, befizetett óra, fedezet).'],
    ['Befizetések – minden számlázz.hu-s befizetés (könyvelt és függőben lévő), tanulóhoz kötve, ha a név egyértelmű.'],
    ['Oktatók – oktatónként: tanulók, vezetett és befizetett óra, fedezet nélküli tanulók.'],
    [''],
    ['Fontos megjegyzések:'],
    ['Fedezet (óra) = befizetett óra – vezetett óra. Negatív (piros) = a tanuló többet vezetett, mint amennyit befizetett.'],
    ['Fizetendő = Ár – Kedvezmény (ellenőrizendő, hogy az "összeg" a beiratkozási táblában kedvezmény előtti ár-e).'],
    ['Hátralék = Fizetendő – Befizetett összesen (csak a beiratkozási táblából jövő tanulóknál van Ár).'],
    ['Oktatói táblában szereplő tanulónál a Befizetett összesen az oktatói tábla száma. A Függő számla (nincs könyvelve) csak tájékoztató, mert lehet, hogy már benne van az oktatói tábla összegében.'],
    ['A személyi igazolvány, lakcímkártya, vezetői engedély számát és az anyja nevét a program szándékosan NEM másolja át (adatvédelem).'],
    ['Törlés: pipáld be a Tanulók lap első oszlopában (Töröl) a tanuló sorát, és a sor azonnal eltűnik. A tanuló a Kizárt lapra kerül, és többet nem jön vissza. Visszaállítás: a Kizárt lapon töröld a sorát, a következő frissítéskor visszakerül. A forrás táblát ez nem érinti.'],
    ['Csak a beiratkozási táblákban szereplő (beiratkozott) tanulók kerülnek ide. Régi tanulók és a hozzájuk tartozó befizetések nincsenek tárolva.'],
    ['A "Teszt adat?" jelzés próba/teszt nevű sorokat jelöl.']
  ];
  sh.getRange(1, 1, t.length, 1).setValues(t);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1, t.length, 1).setWrap(true);
}
function key_(s) {
  s = String(s || '').toLowerCase().replace(/[0-9.]+/g, ' ').replace(/[^a-záéíóöőúüű\s]/g, ' ');
  var t = s.split(/\s+/).filter(function (x) { return x && x !== 'dr'; });
  t.sort();
  return t.join(' ');
}
function num_(v) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return 0;
  var n = parseFloat(String(v).replace(/[^0-9,.\-]/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function round2_(n) { return Math.round(n * 100) / 100; }
function date_(v) {
  if (v instanceof Date) return v;
  var m = String(v || '').match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}
function dateTime_(v) {
  if (v instanceof Date) return v;
  var m = String(v || '').match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) : String(v || '');
}
function prefix_(s) { var m = String(s).match(/^([A-Za-zÁÉÍÓÖŐÚÜŰ]+)/); return m ? m[1] : ''; }
function hash_(s) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return d.slice(0, 4).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}


/* ================= ADATBÁZIS ÉPÍTÉS – 1. lépés (2026.10.15.) =================
   Mit csinál:
   1) A Kategóriák fül VÉGÉRE tesz két új oszlopot: "Csoport" és "Alcsoport",
      és besorolja a meglévő képzéseket. Meglévő oszlopot nem mozgat, nem ír át.
   2) Az E.V. listába felveszi a 4 hiányzó szaktanfolyamot (ár/óraszám üresen marad).
   Újra futtatható: ha már megvannak az oszlopok/sorok, nem duplázza. */
var EP_KFT = '1AJZGLcOlHerO0gV3rY3jaKa2hfl-gQcVzTMjJ1nmmBk';
var EP_EV  = '1bU6oi6bL325iEe6Z0ktzU_sYIPcQ7efqV2DSME9YyvE';

var EP_MAP = {
  'B – személygépkocsi': ['Jogosítvány', 'Személygépkocsi'],
  'B – automata váltós': ['Jogosítvány', 'Személygépkocsi'],
  'B – instant / extra gyors': ['Jogosítvány', 'Személygépkocsi'],
  'B – KRÉTA (középiskolai KRESZ után)': ['Jogosítvány', 'Személygépkocsi'],
  'B – KRÉTA automata (középiskolai KRESZ után)': ['Jogosítvány', 'Személygépkocsi'],
  'AM – segédmotoros kerékpár': ['Jogosítvány', 'Motor'],
  'A1 – könnyű motor': ['Jogosítvány', 'Motor'],
  'A1 – B kategóriás engedéllyel (A1B)': ['Jogosítvány', 'Motor'],
  'A2 – közepes motor': ['Jogosítvány', 'Motor'],
  'A2 – A1 után, 2 éven belül': ['Jogosítvány', 'Motor'],
  'A2 – A1 után, 2 éven túl': ['Jogosítvány', 'Motor'],
  'A – közvetlen (24 éves kortól)': ['Jogosítvány', 'Motor'],
  'A – A2 után, 2 éven belül': ['Jogosítvány', 'Motor'],
  'A – A2 után, 2 éven túl': ['Jogosítvány', 'Motor'],
  'A – A1 után, 2 éven belül': ['Jogosítvány', 'Motor'],
  'A – A1 után, 2 éven túl': ['Jogosítvány', 'Motor'],
  'BE – utánfutó (B+E)': ['Jogosítvány', 'Pótkocsi'],
  'C – tehergépkocsi': ['Jogosítvány', 'Tehergépkocsi'],
  'C – tehergépkocsi (D kategóriával)': ['Jogosítvány', 'Tehergépkocsi'],
  'CE – nehézpótkocsi (C+E)': ['Jogosítvány', 'Tehergépkocsi'],
  'C+CE – párhuzamos (kombinált)': ['Jogosítvány', 'Tehergépkocsi'],
  'D – autóbusz (B kategóriával)': ['Jogosítvány', 'Autóbusz'],
  'D – autóbusz (C kategóriával)': ['Jogosítvány', 'Autóbusz'],
  'GKI alapképzés – áruszállítás (C)': ['GKI', 'Alapképzés'],
  'GKI alapképzés – személyszállítás (D)': ['GKI', 'Alapképzés'],
  'GKI alapképzés – C és D összevont (CDA)': ['GKI', 'Alapképzés'],
  'GKI továbbképzés vizsgával (C vagy D)': ['GKI', 'Továbbképzés vizsgával'],
  'GKI továbbképzés vizsgával – C és D összevont (CDT)': ['GKI', 'Továbbképzés vizsgával'],
  'GKI vizsga nélküli továbbképzés (1 kategória)': ['GKI', 'Vizsga nélküli továbbképzés'],
  'GKI vizsga nélküli továbbképzés (C+D összevont)': ['GKI', 'Vizsga nélküli továbbképzés'],
  'Taxis szolgáltatói képesítés': ['Szaktanfolyam', 'Taxi'],
  'Személygépkocsis személyszállító szolgáltató': ['Szaktanfolyam', 'Taxi'],
  'Elsősegély tanfolyam': ['Szaktanfolyam', 'Elsősegély'],
  'Vezetéstechnikai tréning': ['Szaktanfolyam', 'Vezetéstechnika']
};

/* új sorok: [A képzés neve, B sablon, C képző, D kategória] */
var EP_UJ = [
  ['Taxis szolgáltatói képesítés', 'Szaktanfolyam', 'EV', 'TAXI'],
  ['Személygépkocsis személyszállító szolgáltató', 'Szaktanfolyam', 'EV', 'SZSZ'],
  ['Elsősegély tanfolyam', 'Szaktanfolyam', 'EV', 'ESEGELY'],
  ['Vezetéstechnikai tréning', 'Szaktanfolyam', 'EV', 'VEZTECH']
];

function epites1() {
  var sor = [];
  sor.push(epKategoriak_(EP_KFT, 'Kft', []));
  sor.push(epKategoriak_(EP_EV, 'E.V.', EP_UJ));
  var out = sor.join('\n');
  Logger.log(out);
  return out;
}

function epKategoriak_(id, cimke, ujSorok) {
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName('Kategóriák');
  if (!sh) return cimke + ': nincs "Kategóriák" fül – kihagyva';

  /* --- 1. új sorok beszúrása a megjegyzés-sor elé --- */
  var ujDb = 0;
  if (ujSorok && ujSorok.length) {
    var aOszlop = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
    var meglevo = {}, megjSor = 0, utolsoAdat = 1;
    for (var i = 0; i < aOszlop.length; i++) {
      var v = String(aOszlop[i][0] || '').trim();
      if (!v) continue;
      if (v.indexOf('Kék betű') === 0) { megjSor = i + 1; continue; }
      meglevo[v] = true;
      if (i + 1 > utolsoAdat) utolsoAdat = i + 1;
    }
    var beszur = ujSorok.filter(function (r) { return !meglevo[r[0]]; });
    if (beszur.length) {
      var hova = megjSor > 0 ? megjSor : utolsoAdat + 1;
      sh.insertRowsBefore(hova, beszur.length);
      sh.getRange(hova, 1, beszur.length, 4).setValues(beszur)
        .setBackground('#FFF2CC').setFontColor('#1155CC');
      sh.getRange(hova, 5, beszur.length, 1)
        .setNote('Új szaktanfolyam – az árakat, óraszámokat és a szerződés-szövegeket még ki kell tölteni.');
      ujDb = beszur.length;
    }
  }

  /* --- 2. Csoport / Alcsoport oszlop a lap végére --- */
  var utolsoOszlop = sh.getLastColumn();
  var fejlec = sh.getRange(1, 1, 1, utolsoOszlop).getValues()[0];
  var cCs = fejlec.indexOf('Csoport') + 1;
  var cAl = fejlec.indexOf('Alcsoport') + 1;
  if (!cCs) { cCs = ++utolsoOszlop; }
  if (!cAl) { cAl = ++utolsoOszlop; }
  var fej = sh.getRange(1, cCs, 1, 2);
  fej.setValues([['Csoport', 'Alcsoport']]);
  fej.setFontWeight('bold').setBackground('#EFEFEF').setWrap(true);
  sh.setColumnWidth(cCs, 130); sh.setColumnWidth(cAl, 190);

  var utolsoSor = sh.getLastRow();
  if (utolsoSor < 2) return cimke + ': üres a lista';
  var nevek = sh.getRange(2, 1, utolsoSor - 1, 1).getValues();
  var ki = [], talalt = 0, hianyzo = [];
  for (var r = 0; r < nevek.length; r++) {
    var nev = String(nevek[r][0] || '').trim();
    if (!nev || nev.indexOf('Kék betű') === 0) { ki.push(['', '']); continue; }
    var m = EP_MAP[nev];
    if (m) { ki.push([m[0], m[1]]); talalt++; }
    else { ki.push(['?', '?']); hianyzo.push(nev); }
  }
  sh.getRange(2, cCs, ki.length, 2).setValues(ki);

  return cimke + ': ' + talalt + ' képzés besorolva, ' + ujDb + ' új sor felvéve'
    + (hianyzo.length ? ', BESOROLATLAN: ' + hianyzo.join(' | ') : '');
}


/* ================= ADATBÁZIS ÉPÍTÉS – 2. lépés (2026.10.15.) =================
   1) "Cégek" törzsfül mindkét munkafüzetbe (Kati és Virág tölti).
   2) A Tanulók fül VÉGÉRE új oszlopok, legördülő választékkal:
      Költségviselő | Cég | Pótóra fizetője |
      GKI vizsgák kész | GKI kártya igényelve | GKI kártya megérkezett | GKI kártya postázása
   Meglévő oszlopot nem mozgat, nem ír át. Újra futtatható. */

var EP_CEG_FEJ = ['Cég neve','Rövid név','Székhely','Adószám','Cégjegyzékszám',
  'Kapcsolattartó','Beosztás','Telefon','E-mail','Számlázási cím','Számlázási e-mail',
  'Fizetési határidő (nap)','Pótóra fizetője (alap)','GKI kártya postázása (alap)',
  'Keretszerződés száma','Keretszerződés kelte','Aktív','Megjegyzés'];

/* csak a valóban meglévő partnerek */
var EP_CEG_KEZDO = [
  ['HBPO Professional Services Kft.','HBPO','','','','Németh Noémi','','','',
   '','',30,'Cég','Cég telephelye','','','igen','2026. okt. – 20 fő targonca (3324)'],
  ['T-Busz Tatabányai Közlekedési Kft.','T-Busz','','','','','','','',
   '','',30,'Cég','Cég telephelye','','','igen','D kategória + GKI keretszerződés']
];

var EP_TAN_UJ = [
  ['Költségviselő', ['Magán','Céges'], 120,
   'Ki fizeti a képzést. Céges esetén a Cég oszlopot is töltsd ki.'],
  ['Cég', 'CEGEK', 230,
   'A Cégek fülről választható. Új céget ott kell felvenni.'],
  ['Pótóra fizetője', ['Tanuló','Cég'], 130,
   'A szolgáltatási szerződés szerint ki fizeti a pótórát és a pótvizsgát.'],
  ['GKI vizsgák kész', null, 130,
   'Az a nap, amikor a tanuló minden vizsgájával végzett. Ettől igényelhető a kártya.'],
  ['GKI kártya igényelve', null, 140,
   'Amikor a képző megigényelte a kártyát a KAV-nál. EZT KÉRI VISSZA A KAV.'],
  ['GKI kártya megérkezett', null, 150,
   'Amikor a kész kártya megérkezett.'],
  ['GKI kártya postázása', ['Tanuló lakcíme','Képző irodája','Cég telephelye'], 150,
   'Hová postázza a KAV a kész kártyát. A szerződésben is rögzíteni kell.']
];

function epites2() {
  var sor = [];
  sor.push(epCegek_(EP_KFT, 'Kft'));
  sor.push(epCegek_(EP_EV, 'E.V.'));
  sor.push(epTanuloOszlopok_(EP_KFT, 'Kft'));
  sor.push(epTanuloOszlopok_(EP_EV, 'E.V.'));
  var out = sor.join('\n');
  Logger.log(out);
  return out;
}

function epCegek_(id, cimke) {
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName('Cégek');
  var uj = false;
  if (!sh) { sh = ss.insertSheet('Cégek'); uj = true; }

  sh.getRange(1, 1, 1, EP_CEG_FEJ.length).setValues([EP_CEG_FEJ])
    .setFontWeight('bold').setBackground('#1E73B8').setFontColor('#FFFFFF')
    .setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);
  var sz = [170,90,200,110,130,140,110,120,170,200,170,90,110,140,130,120,70,230];
  for (var i = 0; i < sz.length; i++) sh.setColumnWidth(i + 1, sz[i]);

  if (uj) {
    sh.getRange(2, 1, EP_CEG_KEZDO.length, EP_CEG_FEJ.length).setValues(EP_CEG_KEZDO);
    sh.getRange('A1').setNote('Itt kell felvenni minden céget, amelyik a tanulók képzését fizeti. '
      + 'A tanulói lapon a Cég oszlop ebből a listából választ.');
  }
  var sorok = Math.max(sh.getMaxRows(), 200);
  if (sh.getMaxRows() < sorok) sh.insertRowsAfter(sh.getMaxRows(), sorok - sh.getMaxRows());
  epLista_(sh.getRange(2, 13, sorok - 1, 1), ['Tanuló','Cég']);
  epLista_(sh.getRange(2, 14, sorok - 1, 1), ['Tanuló lakcíme','Képző irodája','Cég telephelye']);
  epLista_(sh.getRange(2, 17, sorok - 1, 1), ['igen','nem']);
  return cimke + ' – Cégek fül: ' + (uj ? 'létrehozva' : 'frissítve');
}

function epLista_(range, ertekek) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(ertekek, true).setAllowInvalid(false).build());
}

function epTanuloOszlopok_(id, cimke) {
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName('Tanulók');
  if (!sh) return cimke + ': nincs "Tanulók" fül';
  var ceg = ss.getSheetByName('Cégek');

  var utolso = sh.getLastColumn();
  var fejlec = sh.getRange(1, 1, 1, utolso).getValues()[0];
  var hozzaadva = 0;

  for (var i = 0; i < EP_TAN_UJ.length; i++) {
    var nev = EP_TAN_UJ[i][0], valaszt = EP_TAN_UJ[i][1],
        szel = EP_TAN_UJ[i][2], megj = EP_TAN_UJ[i][3];
    var hol = fejlec.indexOf(nev) + 1;
    if (!hol) {
      utolso++;
      if (sh.getMaxColumns() < utolso) sh.insertColumnsAfter(sh.getMaxColumns(), utolso - sh.getMaxColumns());
      hol = utolso;
      hozzaadva++;
      fejlec.push(nev);
    }
    var f = sh.getRange(1, hol);
    f.setValue(nev).setFontWeight('bold').setWrap(true)
     .setBackground(nev.indexOf('GKI') === 0 ? '#FBEDE2' : '#E6EFF8')
     .setNote(megj);
    sh.setColumnWidth(hol, szel);

    var sorok = sh.getMaxRows() - 1;
    var tart = sh.getRange(2, hol, sorok, 1);
    if (valaszt === 'CEGEK' && ceg) {
      tart.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(ceg.getRange('A2:A200'), true).setAllowInvalid(true).build());
    } else if (Array.isArray(valaszt)) {
      epLista_(tart, valaszt);
    } else {
      tart.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireDate().setAllowInvalid(false).build());
      tart.setNumberFormat('yyyy. mm. dd.');
    }
  }
  return cimke + ' – Tanulók fül: ' + hozzaadva + ' új oszlop (összesen ' + EP_TAN_UJ.length + ' van a helyén)';
}


/* ---- Képzés → Főcsoport / Alcsoport térkép (a Kategóriák és Gépcsoportok fülekről) ---- */
function csoportTerkep_() {
  var map = {};
  [CFG.KFT_ID, CFG.EV_ID].forEach(function (id) {
    var sh = SpreadsheetApp.openById(id).getSheetByName('Kategóriák');
    if (!sh) return;
    var v = sh.getDataRange().getValues(), h = {};
    v[0].forEach(function (x, i) { h[String(x).trim()] = i; });
    if (h['Csoport'] === undefined) return;
    for (var i = 1; i < v.length; i++) {
      var nev = String(v[i][0] || '').trim();
      if (!nev || nev.indexOf('Kék betű') === 0) continue;
      var cs = String(v[i][h['Csoport']] || '').trim();
      if (cs) map[nev] = [cs, String(v[i][h['Alcsoport']] || '').trim()];
    }
  });
  /* gépkezelő: a Gépcsoportok fülről, a gép típusa az alcsoport */
  var gsh = SpreadsheetApp.openById(CFG.EV_ID).getSheetByName('Gépcsoportok');
  if (gsh) {
    var g = gsh.getDataRange().getValues(), gh = {};
    g[0].forEach(function (x, i) { gh[String(x).trim()] = i; });
    var cN = gh['Választólista'], cT = gh['Típus (kalkulátor)'], cG = gh['Gépcsoport'], cN2 = gh['Választólista (típus szerint)'];
    for (var j = 1; j < g.length; j++) {
      var tip = cT === undefined ? '' : String(g[j][cT] || '').trim();
      [cN, cG, cN2].forEach(function (c) {
        if (c === undefined) return;
        var n = String(g[j][c] || '').trim();
        if (n && !map[n]) map[n] = ['Gépkezelő', tip || 'Egyéb gép'];
      });
    }
  }
  /* Két főcsoport: Kategóriás képzés (ami jogosítvány) és Szaktanfolyam (minden más).
     A GKI és a Gépkezelő a Szaktanfolyam alcsoportja lesz. */
  Object.keys(map).forEach(function (n) {
    var cs = map[n][0], al = map[n][1];
    if (cs === 'Jogosítvány' || cs === 'Kategóriás képzés') {
      map[n] = ['Kategóriás képzés', al || 'Egyéb kategória'];
    } else if (cs === 'Szaktanfolyam') {
      map[n] = ['Szaktanfolyam', al || 'Egyéb szaktanfolyam'];
    } else {
      map[n] = ['Szaktanfolyam', cs];
    }
  });
  return map;
}

/* a nyilvántartásban szereplő, de a törzsben nem található képzésnevek listája */
function csoportEllenorzes() {
  var map = csoportTerkep_(), hiany = {};
  [[CFG.KFT_ID, CFG.KFT_GID], [CFG.EV_ID, CFG.EV_GID]].forEach(function (p) {
    var sh = sheetByGid_(SpreadsheetApp.openById(p[0]), p[1]);
    var v = sh.getDataRange().getValues(), h = {};
    v[0].forEach(function (x, i) { h[String(x).trim()] = i; });
    for (var i = 1; i < v.length; i++) {
      var nev = String(v[i][h['Képzés']] || '').trim();
      if (nev && !map[nev]) hiany[nev] = (hiany[nev] || 0) + 1;
    }
  });
  var k = Object.keys(hiany);
  var out = k.length
    ? 'BESOROLATLAN KÉPZÉSNEVEK:\n' + k.map(function (n) { return '  ' + n + '  (' + hiany[n] + ' tanuló)'; }).join('\n')
    : 'Minden képzésnév be van sorolva.';
  Logger.log(out);
  return out;
}


/* ================= ADATBÁZIS ÉPÍTÉS – 3. lépés: KIMUTATÁS =================
   Létrehozza a "Kimutatás", "GKI kártyák" és a rejtett "_segéd" lapot
   a ZOOM Adatbázis (csak iroda) munkafüzetben. A Tanulók lapra épül,
   képletekkel – azonnal frissül, ha a szűrőt átállítod. */

var HO_NEVEK = ['Január','Február','Március','Április','Május','Június',
  'Július','Augusztus','Szeptember','Október','November','December'];
var IDOSZAKOK = ['Teljes év','I. félév (jan–jún)','II. félév (júl–dec)',
  '1. negyedév','2. negyedév','3. negyedév','4. negyedév'].concat(HO_NEVEK);
var DATUM_ALAP = ['Beiratkozás (kezdés)','Szerződés kelte','GKI kártya igénylése','GKI kártya megérkezése'];
var FIZ_ALLAPOT = ['Mind','Teljesen fizetett','Részben fizetett','Nincs befizetés'];

function f_(s) { return s.replace(/~/g, String.fromCharCode(39)).replace(/@/g, String.fromCharCode(34)); }

function epites3() {
  var db = SpreadsheetApp.openById('1WnMwIimpAt5eUYmKTliEIoPLsH6Z4TqQ8mgzXjYACmM');
  var seg = epSeged_(db);
  epKimutatas_(db, seg);
  epGkiLap_(db);
  epNezetek_(db);
  db.setActiveSheet(db.getSheetByName('Kimutatás'));
  var out = 'Kimutatás, GKI kártyák és _segéd lap kész.';
  Logger.log(out); return out;
}

function epSeged_(db) {
  var sh = db.getSheetByName('_segéd') || db.insertSheet('_segéd');
  sh.clear();
  sh.getRange(1, 1, 12, 1).setValues(HO_NEVEK.map(function (h) { return [h]; }));

  var K = 'Kimutatás!';
  sh.getRange('B1').setFormula(f_('=IFERROR(SWITCH(' + K + 'D5,@Teljes év@,1,@I. félév (jan–jún)@,1,@II. félév (júl–dec)@,7,@1. negyedév@,1,@2. negyedév@,4,@3. negyedév@,7,@4. negyedév@,10,MATCH(' + K + 'D5,$A$1:$A$12,0)),1)'));
  sh.getRange('B2').setFormula(f_('=IFERROR(SWITCH(' + K + 'D5,@Teljes év@,12,@I. félév (jan–jún)@,6,@II. félév (júl–dec)@,12,@1. negyedév@,3,@2. negyedév@,6,@3. negyedév@,9,@4. negyedév@,12,MATCH(' + K + 'D5,$A$1:$A$12,0)),12)'));
  sh.getRange('B3').setFormula(f_('=IF(' + K + 'C5=@Mind@,DATE(2000,1,1),DATE(VALUE(' + K + 'C5),B1,1))'));
  sh.getRange('B4').setFormula(f_('=IF(' + K + 'C5=@Mind@,DATE(2100,12,31),EOMONTH(DATE(VALUE(' + K + 'C5),B2,1),0))'));
  sh.getRange('B6').setFormula(f_('=SWITCH(' + K + 'B5,@Beiratkozás (kezdés)@,@P@,@Szerződés kelte@,@Q@,@GKI kártya igénylése@,@AL@,@GKI kártya megérkezése@,@AM@)'));
  sh.getRange('B7').setFormula(f_(
    '=@where B is not null@'
    + '&@ and @&B6&@ >= date ~@&TEXT(B3,@yyyy-mm-dd@)&@~@'
    + '&@ and @&B6&@ <= date ~@&TEXT(B4,@yyyy-mm-dd@)&@~@'
    + '&IF(' + K + 'E5=@Mind@,@@,@ and C = ~@&' + K + 'E5&@~@)'
    + '&IF(' + K + 'F5=@Mind@,@@,@ and AF = ~@&' + K + 'F5&@~@)'
    + '&IF(' + K + 'G5=@Mind@,@@,@ and AG = ~@&' + K + 'G5&@~@)'
    + '&IF(' + K + 'H5=@Mind@,@@,@ and D = ~@&' + K + 'H5&@~@)'
    + '&IF(' + K + 'I5=@Mind@,@@,@ and AH = ~@&' + K + 'I5&@~@)'
    + '&IF(' + K + 'J5=@Mind@,@@,@ and AI = ~@&' + K + 'J5&@~@)'
    + '&IF(' + K + 'K5=@Teljesen fizetett@,@ and I <= 0@,IF(' + K + 'K5=@Nincs befizetés@,@ and H = 0@,IF(' + K + 'K5=@Részben fizetett@,@ and H > 0 and I > 0@,@@)))'
    + '&IF(' + K + 'L5=@Mind@,@@,@ and AO = ~@&' + K + 'L5&@~@)'));
  sh.getRange('A14').setValue('Ide ne írj semmit – ez a Kimutatás segédlapja.');

  var T = 'Tanulók!';
  var listak = [
    ['D', 'C'],   // Képző
    ['E', 'AF'],  // Főcsoport
    ['F', 'AG'],  // Alcsoport
    ['G', 'D'],   // Képzés
    ['I', 'AI']   // Cég
  ];
  listak.forEach(function (p) {
    sh.getRange(p[0] + '1').setValue('Mind');
    sh.getRange(p[0] + '2').setFormula(f_('=IFERROR(SORT(UNIQUE(FILTER(' + T + p[1] + '2:' + p[1] + ',' + T + p[1] + '2:' + p[1] + '<>@@))),@@)'));
  });
  sh.getRange('H1:H3').setValues([['Mind'], ['Magán'], ['Céges']]);
  sh.getRange('J1:J' + FIZ_ALLAPOT.length).setValues(FIZ_ALLAPOT.map(function (x) { return [x]; }));
  sh.getRange('K1:K' + DATUM_ALAP.length).setValues(DATUM_ALAP.map(function (x) { return [x]; }));
  sh.getRange('L1:L4').setValues([['Mind'], ['2026'], ['2025'], ['2024']]);
  sh.getRange('M1:M' + IDOSZAKOK.length).setValues(IDOSZAKOK.map(function (x) { return [x]; }));
  sh.getRange('N1:N5').setValues([['Mind'], ['E-learning'], ['Tantermi'], ['Vegyes'], ['Nem releváns']]);
  sh.hideSheet();
  return sh;
}

function epLista2_(cella, range) {
  cella.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true).setAllowInvalid(false).build());
}

function epKimutatas_(db, seg) {
  var sh = db.getSheetByName('Kimutatás');
  if (!sh) { sh = db.insertSheet('Kimutatás', 0); } else { sh.clear(); sh.clearConditionalFormatRules(); }
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 20) sh.insertColumnsAfter(sh.getMaxColumns(), 20 - sh.getMaxColumns());

  sh.getRange('B2').setValue('ZOOM – Kimutatás')
    .setFontSize(18).setFontWeight('bold').setFontColor('#1E73B8');
  sh.getRange('B3').setValue('Állítsd be a szűrőket, a számok azonnal frissülnek. A Tanulók lapról dolgozik.')
    .setFontColor('#4A5765');

  var cimkek = ['Dátum alapja', 'Év', 'Időszak', 'Képző', 'Főcsoport', 'Alcsoport',
    'Képzés', 'Költségviselő', 'Cég', 'Fizetési állapot', 'Elméleti képzés'];
  sh.getRange(4, 2, 1, cimkek.length).setValues([cimkek])
    .setFontWeight('bold').setFontSize(9).setFontColor('#78848F').setWrap(true);
  sh.getRange(5, 2, 1, cimkek.length)
    .setValues([['Beiratkozás (kezdés)', '2026', 'Teljes év', 'Mind', 'Mind', 'Mind', 'Mind', 'Mind', 'Mind', 'Mind', 'Mind']])
    .setBackground('#FFFFFF').setBorder(true, true, true, true, true, true, '#DCE3EB', null)
    .setFontWeight('bold');
  epLista2_(sh.getRange('B5'), seg.getRange('K1:K9'));
  epLista2_(sh.getRange('C5'), seg.getRange('L1:L9'));
  epLista2_(sh.getRange('D5'), seg.getRange('M1:M25'));
  epLista2_(sh.getRange('E5'), seg.getRange('D1:D50'));
  epLista2_(sh.getRange('F5'), seg.getRange('E1:E50'));
  epLista2_(sh.getRange('G5'), seg.getRange('F1:F80'));
  epLista2_(sh.getRange('H5'), seg.getRange('G1:G90'));
  epLista2_(sh.getRange('I5'), seg.getRange('H1:H3'));
  epLista2_(sh.getRange('J5'), seg.getRange('I1:I200'));
  epLista2_(sh.getRange('K5'), seg.getRange('J1:J9'));
  epLista2_(sh.getRange('L5'), seg.getRange('N1:N5'));

  sh.getRange('B6').setFormula(f_('=@Időszak: @&TEXT(_segéd!B3,@yyyy. mm. dd.@)&@ – @&TEXT(_segéd!B4,@yyyy. mm. dd.@)&@   (@&B5&@ szerint)@'))
    .setFontColor('#78848F').setFontSize(9);

  var kpiC = ['Beiratkozó', 'Szerződött díj (Ft)', 'Befizetett (Ft)', 'Hátralék (Ft)'];
  sh.getRange(8, 2, 1, 4).setValues([kpiC]).setFontWeight('bold').setFontSize(9)
    .setFontColor('#FFFFFF').setBackground('#1E73B8').setHorizontalAlignment('center');
  var T = 'Tanulók!$A$2:$AN', W = '&_segéd!$B$7';
  sh.getRange('B9').setFormula(f_('=IFERROR(ROWS(QUERY(' + T + ',@select B @' + W + ',0)),0)'));
  sh.getRange('C9').setFormula(f_('=IFERROR(SUM(QUERY(' + T + ',@select G @' + W + ',0)),0)'));
  sh.getRange('D9').setFormula(f_('=IFERROR(SUM(QUERY(' + T + ',@select H @' + W + ',0)),0)'));
  sh.getRange('E9').setFormula(f_('=IFERROR(SUM(QUERY(' + T + ',@select I @' + W + ',0)),0)'));
  sh.getRange('B9:E9').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#DCE3EB', null);
  sh.getRange('C9:E9').setNumberFormat('#,##0');
  sh.getRange('E9').setFontColor('#B4342A');

  sh.getRange('B11').setValue('BONTÁS – képző, csoport, alcsoport, képzés')
    .setFontWeight('bold').setFontColor('#1E73B8');
  var bf = ['Képző', 'Főcsoport', 'Alcsoport', 'Képzés', 'Fő', 'Szerződött (Ft)', 'Befizetett (Ft)', 'Hátralék (Ft)'];
  sh.getRange(12, 2, 1, bf.length).setValues([bf]).setFontWeight('bold').setFontSize(9)
    .setBackground('#E9EFF5').setBorder(null, null, true, null, null, null, '#1E73B8', null);
  sh.getRange('B13').setFormula(f_('=IFERROR(QUERY(' + T + ',@select C, AF, AG, D, count(B), sum(G), sum(H), sum(I) @' + W + '&@ group by C, AF, AG, D order by C, AF, AG, D@,0),@Erre a szűrésre nincs találat.@)'));
  sh.getRange('F13:I400').setNumberFormat('#,##0');

  sh.getRange('K11').setValue('CÉGES MEGRENDELŐK').setFontWeight('bold').setFontColor('#D4620E');
  var cf = ['Cég', 'Fő', 'Szerződött (Ft)', 'Befizetett (Ft)', 'Hátralék (Ft)'];
  sh.getRange(12, 11, 1, cf.length).setValues([cf]).setFontWeight('bold').setFontSize(9)
    .setBackground('#FBEDE2').setBorder(null, null, true, null, null, null, '#D4620E', null);
  sh.getRange('K13').setFormula(f_('=IFERROR(QUERY(' + T + ',@select AI, count(B), sum(G), sum(H), sum(I) @' + W + '&@ and AI is not null group by AI order by count(B) desc@,0),@Nincs céges megrendelő ebben a szűrésben.@)'));
  sh.getRange('L13:O400').setNumberFormat('#,##0');

  [40, 170, 60, 140, 110, 130, 160, 230, 130, 110, 150, 230, 60, 120, 120, 120]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(9);
  return sh;
}

function epGkiLap_(db) {
  var sh = db.getSheetByName('GKI kártyák') || db.insertSheet('GKI kártyák');
  sh.clear(); sh.setHiddenGridlines(true);
  sh.getRange('A1').setValue('GKI kártyák – igénylés és nyilvántartás')
    .setFontSize(15).setFontWeight('bold').setFontColor('#D4620E');
  var fej = ['Tanuló', 'Képző', 'Képzés', 'Megrendelő cég', 'Vizsgák kész',
    'Kártya igényelve', 'Kártya megérkezett', 'Postázás', 'Állapot'];
  sh.getRange(2, 1, 1, fej.length).setValues([fej]).setFontWeight('bold').setFontSize(9)
    .setFontColor('#FFFFFF').setBackground('#D4620E');
  sh.getRange('A3').setFormula(f_('=IFERROR(QUERY(Tanulók!$A$2:$AN,@select B, C, D, AI, AK, AL, AM, AN where D contains ~GKI~ or AK is not null or AL is not null or AM is not null order by AL desc@,0),@Még nincs GKI-s tanuló.@)'));
  sh.getRange('I3').setFormula(f_('=ARRAYFORMULA(IF(A3:A=@@,@@,IF(G3:G<>@@,@Kártya megérkezett@,IF(F3:F<>@@,@Igényelve, kártyára vár@,IF(E3:E<>@@,@IGÉNYELNI KELL – minden vizsga megvan@,@Vizsgák folyamatban@)))))'));
  [170, 60, 240, 210, 110, 120, 130, 140, 240].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('E3:G400').setNumberFormat('yyyy. mm. dd.');
  sh.setFrozenRows(2);
  var r1 = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('IGÉNYELNI')
    .setBackground('#FAE9E7').setFontColor('#B4342A').setBold(true).setRanges([sh.getRange('I3:I400')]).build();
  var r2 = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Kártya megérkezett')
    .setBackground('#E4F3EB').setFontColor('#0E8A52').setRanges([sh.getRange('I3:I400')]).build();
  sh.setConditionalFormatRules([r1, r2]);
  return sh;
}


/* ================= 4. lépés: MENTETT NÉZETEK =================
   A Kimutatás L5 cellájában választható mentett nézet.
   A nézetek a "Nézetek" lapon vannak – közösek, mindenki látja őket. */

var NEZET_FEJ = ['Nézet neve', 'Dátum alapja', 'Év', 'Időszak', 'Képző', 'Főcsoport',
  'Alcsoport', 'Képzés', 'Költségviselő', 'Cég', 'Fizetési állapot', 'Elméleti képzés', 'Mire jó'];

var NEZET_KEZDO = [
  ['Teljes idei év','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Mind','Mind',
   'Minden idei beiratkozó.'],
  ['Idei GKI – minden','Beiratkozás (kezdés)','2026','Teljes év','Mind','GKI','Mind','Mind','Mind','Mind','Mind','Mind',
   'Alapképzés, továbbképzés és vizsga nélküli együtt.'],
  ['GKI kártya igénylések – idei','GKI kártya igénylése','2026','Teljes év','Mind','GKI','Mind','Mind','Mind','Mind','Mind','Mind',
   'Mikor melyik kártyát igényeltük – ezt kéri vissza a KAV.'],
  ['Céges tanulók – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Céges','Mind','Mind','Mind',
   'Akiket a munkáltató fizet.'],
  ['Hátralékosok – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Részben fizetett','Mind',
   'Akik fizettek valamennyit, de még tartoznak.'],
  ['Nem fizettek – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Nincs befizetés','Mind',
   'Akiktől még egy forint sem érkezett.'],
  ['Gépkezelő – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Gépkezelő','Mind','Mind','Mind','Mind','Mind','Mind',
   'Targonca, emelőgép, földmunkagép.'],
  ['Kft (B kategória) – idei','Beiratkozás (kezdés)','2026','Teljes év','Kft','Mind','Mind','Mind','Mind','Mind','Mind','Mind',
   'Csak a Kft képzései.'],
  ['E-learning elmélet – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Mind','E-learning',
   'Akik e-learninggel tanulják az elméletet.'],
  ['Tantermi elmélet – idei','Beiratkozás (kezdés)','2026','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Mind','Tantermi',
   'Akik tanteremben tanulják az elméletet.'],
  ['Tavalyi teljes év','Beiratkozás (kezdés)','2025','Teljes év','Mind','Mind','Mind','Mind','Mind','Mind','Mind','Mind',
   'Összehasonlításhoz.']
];

function epites4() {
  var db = SpreadsheetApp.openById('1WnMwIimpAt5eUYmKTliEIoPLsH6Z4TqQ8mgzXjYACmM');
  var out = epNezetek_(db);
  Logger.log(out); return out;
}

function epNezetek_(db) {
  var n = db.getSheetByName('Nézetek'), uj = false;
  if (!n) { n = db.insertSheet('Nézetek'); uj = true; }
  n.getRange(1, 1, 1, NEZET_FEJ.length).setValues([NEZET_FEJ])
    .setFontWeight('bold').setBackground('#0E8A52').setFontColor('#FFFFFF').setWrap(true);
  n.setFrozenRows(1);
  [200, 170, 70, 150, 80, 120, 180, 230, 110, 180, 140, 130, 320]
    .forEach(function (w, i) { n.setColumnWidth(i + 1, w); });
  if (uj) {
    n.getRange(2, 1, NEZET_KEZDO.length, NEZET_FEJ.length).setValues(NEZET_KEZDO);
    n.getRange('A1').setNote('Saját nézetet a Kimutatás lapon a ZOOM menü → '
      + '„Mostani nézet mentése…” ponttal lehet felvenni. Kézzel is írhatsz ide sort.');
  }

  var sh = db.getSheetByName('Kimutatás');
  if (sh) {
    sh.getRange('M4').setValue('Mentett nézet')
      .setFontWeight('bold').setFontSize(9).setFontColor('#0E8A52').setWrap(true);
    sh.getRange('M5').setBackground('#E4F3EB').setFontWeight('bold')
      .setBorder(true, true, true, true, true, true, '#0E8A52', null)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(n.getRange('A2:A100'), true).setAllowInvalid(true).build());
    sh.getRange('M6').setValue('← válassz, és a szűrők beállnak')
      .setFontSize(9).setFontColor('#78848F');
    sh.setColumnWidth(13, 230);
  }
  return 'Nézetek lap kész (' + (n.getLastRow() - 1) + ' nézet), a Kimutatás L5 cellájából választható.';
}

/* a Kimutatás L5 cellájának módosítására betölti a nézetet */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== 'Kimutatás') return;
    if (e.range.getA1Notation() !== 'M5') return;
    var nev = String(e.range.getValue() || '').trim();
    if (!nev) return;
    var n = sh.getParent().getSheetByName('Nézetek');
    if (!n) return;
    var v = n.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][0]).trim() === nev) {
        sh.getRange(5, 2, 1, 11).setValues([v[i].slice(1, 12)]);
        return;
      }
    }
  } catch (err) { /* csendben */ }
}

function nezetMentes() {
  var ui = SpreadsheetApp.getUi(), db = SpreadsheetApp.getActive();
  var sh = db.getSheetByName('Kimutatás'), n = db.getSheetByName('Nézetek');
  if (!sh || !n) { ui.alert('Nincs Kimutatás vagy Nézetek lap.'); return; }
  var r = ui.prompt('Nézet mentése', 'Add meg a nézet nevét:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var nev = String(r.getResponseText() || '').trim();
  if (!nev) { ui.alert('Név nélkül nem tudom elmenteni.'); return; }
  var ertek = sh.getRange(5, 2, 1, 11).getValues()[0];
  var v = n.getDataRange().getValues(), sor = 0;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === nev) sor = i + 1;
  var ujE = !sor;
  if (!sor) sor = n.getLastRow() + 1;
  n.getRange(sor, 1, 1, 12).setValues([[nev].concat(ertek)]);
  if (ujE) n.getRange(sor, 13).setValue('Mentve: ' + Utilities.formatDate(new Date(), 'Europe/Budapest', 'yyyy. MM. dd. HH:mm'));
  ui.alert(ujE ? 'Elmentve új nézetként: ' + nev : 'Felülírva: ' + nev);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ZOOM')
    .addItem('📊 Élő kimutatás megnyitása', 'eloKimutatas')
    .addSeparator()
    .addItem('Mostani nézet mentése…', 'nezetMentes')
    .addItem('Adatok frissítése most', 'szinkron')
    .addItem('Besorolatlan képzésnevek', 'csoportEllenorzesUi')
    .addToUi();
}

/* A grafikus kimutatás megnyitása a táblázatból, friss adattal. */
function eloKimutatas() {
  var html = HtmlService.createHtmlOutputFromFile('oldal').getContent();
  html = html.split('%%ADATOK%%').join(JSON.stringify(zoomAdatok_()));
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1900).setHeight(1200),
    'ZOOM – élő kimutatás'
  );
}

function csoportEllenorzesUi() {
  SpreadsheetApp.getUi().alert(csoportEllenorzes());
}


/* ================= ÚTMUTATÓ KÉSZÍTÉSE (Google Doc a Drive-ba) ================= */
var UTM_TARTALOM = [
['c','ZOOM Autósiskola – adatbázis útmutató'],
['al','Teljesen kezdőknek. Nem kell semmihez érteni, csak elolvasni és követni.'],

['h1','1. Mi ez az egész, és miből áll?'],
['p','Három Google-táblázat dolgozik együtt. Mindhárom a böngészőben nyílik meg, semmit nem kell telepíteni.'],
['t',['Táblázat neve','Mi van benne','Ki használja']],
['t',['ZOOM_Kft_B_nyilvantartas_2','A Kft. B kategóriás tanulói, szerződései.','Kati, Virág, Laci']],
['t',['ZOOM_EV_nyilvantartas_2','Az E.V. tanulói: motor, C, CE, D, GKI, taxi, gépkezelő.','Kati, Virág, Laci']],
['t',['ZOOM Adatbázis (csak iroda)','Ide folyik össze minden. Itt vannak a kimutatások és a pénzügy.','Csak Laci és Adrienn']],
['p','A harmadik táblázatba SENKI nem ír kézzel. Oda egy program másolja át az adatokat minden éjjel 3 órakor. Ha most azonnal frissíteni akarod, a ZOOM menüben van rá gomb.'],

['h1','2. A legfontosabb szabály'],
['box','Mindent a két nyilvántartásban (Kft, E.V.) kell beírni. A harmadik táblázat csak megmutatja az adatokat, oda írni felesleges – a következő éjjel úgyis felülíródik.'],

['h1','3. Új tanuló felvétele – magánszemély'],
['n','Nyisd meg a megfelelő táblázatot: B kategória a Kft-ben, minden más az E.V.-ben.'],
['n','A Kezdőlapon kattints a zöld „ÚJ TANULÓ FELVÉTELE” gombra. Megnyílik egy űrlap.'],
['n','Töltsd ki az űrlapot, és küldd be. A tanuló magától megjelenik a Tanulók lapon.'],
['n','A Tanulók lapon a „Költségviselő” oszlopba írd be: Magán.'],
['n','A szerződést a ZOOM menü → „Szerződés PDF” ponttal nyomtathatod.'],

['h1','4. Céges megrendelő – szolgáltatási szerződés'],
['p','Ez akkor kell, ha nem a tanuló fizet, hanem a munkáltatója. Például a HBPO 20 dolgozója targoncára.'],
['h2','4.1 Először vedd fel a céget'],
['n','Nyisd meg a „Cégek” lapot (a Kft-ben és az E.V.-ben is van ilyen).'],
['n','Írj egy új sort: cég neve, székhely, adószám, kapcsolattartó, telefon, e-mail, számlázási cím.'],
['n','A „Fizetési határidő (nap)” mezőbe általában 30 kerül.'],
['n','A „Pótóra fizetője (alap)” és a „GKI kártya postázása (alap)” azt mondja meg, mi legyen az alapértelmezés ennél a cégnél. Ezt elég egyszer beállítani.'],
['h2','4.2 Aztán töltsd ki a szerződést'],
['n','ZOOM menü → „Szolgáltatási szerződés – új (üres) lap”. Létrejön egy üres lap.'],
['n','Töltsd ki a SÁRGA cellákat: iktatószám (pl. SZ-2026-0001), szerződés kelte, megrendelő cég, képzés.'],
['n','A cég kiválasztásakor a kapcsolattartó, a számlázási e-mail és a fizetési határidő magától kitöltődik. A képzés kiválasztásakor jön a díj/fő.'],
['n','A lap alsó felén, a MELLÉKLET részbe írd be a dolgozókat: név, születési adatok, anyja neve, lakcím, telefon, e-mail, végzettség. 50 fő fér rá.'],
['n','Ha valakinél más a képzés vagy más fizeti a pótóráját, azt a sor végén felülírhatod.'],
['n','A létszám és a mindösszesen magától kiszámolódik.'],
['h2','4.3 Végül vidd be őket tanulóként'],
['n','ZOOM menü → „Szolgáltatási szerződés – dolgozók felvétele tanulóként”.'],
['n','Megkérdezi, hány fő és melyik céghez. Ha rendben, kattints az OK-ra.'],
['n','Mindenki bekerül a Tanulók lapra. Automatikusan beíródik: Költségviselő = Céges, Cég = a megrendelő, pótóra fizetője, képzés, szerződés kelte, és a megjegyzésbe az iktatószám.'],
['note','Aki már bent van (ugyanaz a név és képzés), azt kihagyja – nem lesz duplán. A mellékletben a „Felvéve” oszlopba beírja a dátumot, így látod, kit vitt már át.'],

['h1','5. GKI kártya – ez a legfontosabb új rész'],
['p','A GKI kártyát CSAK a képző szerv tudja megigényelni a KAV-nál, és csak akkor, ha a tanuló minden vizsgájával végzett. A KAV utólag gyakran visszakérdez, hogy mikor igényeltük – ezért ezt most rögzítjük.'],
['h2','5.1 Mit kell beírni a Tanulók lapon'],
['t',['Oszlop','Mikor töltsd ki']],
['t',['GKI vizsgák kész','Amikor a tanuló az utolsó vizsgáját is letette. Innentől igényelhető a kártya.']],
['t',['GKI kártya igényelve','Amikor elküldted a KAV-nak az igénylést. EZT KÉRDEZI VISSZA A KAV.']],
['t',['GKI kártya megérkezett','Amikor a kész kártya megjött.']],
['t',['GKI kártya postázása','Hová küldje a KAV: tanuló lakcíme / képző irodája / cég telephelye.']],
['h2','5.2 Hol látod, kinek kell kártyát igényelni'],
['p','A „ZOOM Adatbázis (csak iroda)” táblázatban van egy „GKI kártyák” lap. Ott minden GKI-s tanuló egy listában van, és az Állapot oszlop magától megmondja, mi a teendő:'],
['b','Vizsgák folyamatban – még nem lehet igényelni.'],
['b','IGÉNYELNI KELL – minden vizsga megvan (pirossal) – ezt kell ma elintézni.'],
['b','Igényelve, kártyára vár – elküldtük, várjuk.'],
['b','Kártya megérkezett (zölddel) – kész.'],
['note','A szerződésbe is bele kell írni, hogy a KAV hová postázza a kész kártyát. A szolgáltatási szerződésnél ez a „GKI kártya postázása” mező.'],

['h1','6. A Kimutatás – itt nézed meg a számokat'],
['p','Nyisd meg a „ZOOM Adatbázis (csak iroda)” táblázatot, és menj a „Kimutatás” lapra.'],
['h2','6.1 A szűrősor'],
['p','Felül tíz legördülő van. Mindegyiket átállíthatod, és a számok AZONNAL újraszámolnak. Nem kell gombot nyomni.'],
['t',['Szűrő','Mit csinál']],
['t',['Dátum alapja','Mi szerint számoljon: beiratkozás, szerződés kelte, GKI kártya igénylése vagy megérkezése.']],
['t',['Év','2026 vagy 2025.']],
['t',['Időszak','Teljes év, I. vagy II. félév, 1–4. negyedév, vagy egy konkrét hónap.']],
['t',['Képző','Kft vagy E.V.']],
['t',['Főcsoport','Kategóriás képzés vagy Szaktanfolyam.']],
['t',['Alcsoport','Pl. Motor, Tehergépkocsi, Targonca, Taxi.']],
['t',['Képzés','Egy konkrét képzés, pl. „B – instant / extra gyors”.']],
['t',['Költségviselő','Magán vagy Céges.']],
['t',['Cég','Egy konkrét megrendelő cég.']],
['t',['Fizetési állapot','Teljesen fizetett, részben fizetett, nincs befizetés.']],
['h2','6.2 Mit látsz alatta'],
['b','Négy nagy szám: beiratkozó, szerződött díj, befizetett, hátralék.'],
['b','Bontás: képző → főcsoport → alcsoport → képzés, mindegyiknél fő és forint.'],
['b','Jobb oldalt a céges megrendelők listája, hátralékkal.'],
['h2','6.3 Mentett nézetek'],
['p','A szűrősor jobb végén van egy zöld cella: „Mentett nézet”. Ha kiválasztasz egyet, a tíz szűrő magától beáll. Kilenc nézet eleve bent van, például „GKI kártya igénylések – idei” vagy „Hátralékosok – idei”.'],
['p','Saját nézet mentése: állítsd be a szűrőket, majd ZOOM menü → „Mostani nézet mentése…”, és adj neki nevet. Amit te elmentesz, azt Kati és Virág is látja.'],
['h2','6.4 Nyomtatás'],
['n','Állítsd be a szűrőket úgy, ahogy látni akarod.'],
['n','Fájl → Nyomtatás (vagy Ctrl+P).'],
['n','Választhatsz nyomtatót, vagy menthetsz PDF-be.'],

['h1','7. Ki mit lát'],
['t',['Ki','Mit']],
['t',['Laci','Mindent.']],
['t',['Adrienn','Mindent, ő a fő szerkesztő az oktatói tábláknál.']],
['t',['Kati és Virág','A két nyilvántartást és a Cégek lapot szerkeszthetik. A pénzügyi összesítőket nem látják.']],
['t',['Oktatók','Csak a saját táblájukat, és azt is csak olvasni. Egymásét nem látják.']],

['h1','8. Ha valami nem stimmel'],
['t',['Tünet','Mit csinálj']],
['t',['Nem látom a ZOOM menüt','Töltsd újra az oldalt (F5). Első alkalommal a Google engedélyt kér – ezt egyszer kell jóváhagyni.']],
['t',['A kimutatásban nincs benne az új tanuló','A program éjjel 3-kor frissít. Azonnali frissítés: ZOOM menü → „Adatok frissítése most”.']],
['t',['Egy képzés nem jelenik meg a bontásban','ZOOM menü → „Besorolatlan képzésnevek”. Megmondja, melyik képzés hiányzik a Kategóriák lap Csoport/Alcsoport oszlopából.']],
['t',['A cég nem választható ki','Előbb vedd fel a Cégek lapon. A legördülő onnan dolgozik.']],
['t',['#ERROR! vagy #N/A egy cellában','Általában azért, mert egy hivatkozott mező még üres. Töltsd ki a sárga cellákat.']],

['h1','9. Rövid szótár'],
['t',['Szó','Mit jelent']],
['t',['Költségviselő','Ki fizeti a képzést: a tanuló (magán) vagy a munkáltatója (céges).']],
['t',['Főcsoport','A két nagy csoport: Kategóriás képzés (ami jogosítvány) és Szaktanfolyam (minden más).']],
['t',['Alcsoport','A főcsoporton belüli bontás, pl. Motor, Targonca, Taxi.']],
['t',['Hátralék','Amennyit még nem fizettek be a szerződött díjból.']],
['t',['FAR','A felnőttképzési adatszolgáltatási rendszer. Minden képzést be kell jelenteni bele.']],
['t',['KAV','A vizsgaközpont. Tőlük jönnek a vizsgaértesítők, és ők adják ki a GKI kártyát.']],
['t',['Szolgáltatási szerződés','A céggel kötött szerződés, ha a munkáltató fizeti a dolgozói képzését.']]
];

function utmutatoKeszites() {
  var cim = 'ZOOM Adatbázis – útmutató (kezdőknek)';
  var regi = DriveApp.getFilesByName(cim);
  while (regi.hasNext()) { regi.next().setTrashed(true); }

  var doc = DocumentApp.create(cim);
  var b = doc.getBody();
  b.clear();
  b.setMarginTop(50).setMarginBottom(50).setMarginLeft(56).setMarginRight(56);

  var tablaSorok = null;
  function tablaLezar() {
    if (!tablaSorok) return;
    var t = b.appendTable(tablaSorok);
    t.setBorderColor('#C8D2DC');
    for (var c = 0; c < tablaSorok[0].length; c++) {
      var cell = t.getRow(0).getCell(c);
      cell.setBackgroundColor('#1E73B8');
      cell.editAsText().setForegroundColor('#FFFFFF').setBold(true);
    }
    for (var r = 0; r < t.getNumRows(); r++) {
      for (var c2 = 0; c2 < t.getRow(r).getNumCells(); c2++) {
        t.getRow(r).getCell(c2).editAsText().setFontSize(10);
      }
      if (r > 0 && r % 2 === 0) for (var c3 = 0; c3 < t.getRow(r).getNumCells(); c3++)
        t.getRow(r).getCell(c3).setBackgroundColor('#F4F7FA');
    }
    tablaSorok = null;
  }

  UTM_TARTALOM.forEach(function (e) {
    var tip = e[0], sz = e[1];
    if (tip !== 't') tablaLezar();
    if (tip === 'c') {
      var p = b.appendParagraph(sz);
      p.setHeading(DocumentApp.ParagraphHeading.TITLE);
      p.editAsText().setForegroundColor('#1E73B8');
    } else if (tip === 'al') {
      var p2 = b.appendParagraph(sz);
      p2.editAsText().setItalic(true).setForegroundColor('#555555').setFontSize(11);
      p2.setSpacingAfter(16);
    } else if (tip === 'h1') {
      var h = b.appendParagraph(sz);
      h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      h.editAsText().setForegroundColor('#1E73B8');
      h.setSpacingBefore(20).setSpacingAfter(6);
    } else if (tip === 'h2') {
      var h2 = b.appendParagraph(sz);
      h2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      h2.editAsText().setForegroundColor('#2E6DB4');
      h2.setSpacingBefore(12).setSpacingAfter(4);
    } else if (tip === 'p') {
      b.appendParagraph(sz).setSpacingAfter(8).editAsText().setFontSize(11);
    } else if (tip === 'b') {
      b.appendListItem(sz).setGlyphType(DocumentApp.GlyphType.BULLET)
        .editAsText().setFontSize(11);
    } else if (tip === 'n') {
      b.appendListItem(sz).setGlyphType(DocumentApp.GlyphType.NUMBER)
        .editAsText().setFontSize(11);
    } else if (tip === 'note') {
      var n = b.appendParagraph('Megjegyzés: ' + sz);
      n.editAsText().setItalic(true).setForegroundColor('#6B5600').setFontSize(10);
      n.setSpacingBefore(6).setSpacingAfter(10);
    } else if (tip === 'box') {
      var tb = b.appendTable([[sz]]);
      tb.getRow(0).getCell(0).setBackgroundColor('#FFF2CC');
      tb.getRow(0).getCell(0).editAsText().setBold(true).setFontSize(11);
      tb.setBorderColor('#E0B800');
    } else if (tip === 't') {
      tablaSorok = tablaSorok || [];
      tablaSorok.push(sz);
    }
  });
  tablaLezar();

  b.appendParagraph('').setSpacingBefore(18);
  var l = b.appendParagraph('Készült: ' + Utilities.formatDate(new Date(), 'Europe/Budapest', 'yyyy. MM. dd.')
    + ' · Ha valami nem világos, szólj, és kiegészítjük.');
  l.editAsText().setFontSize(9).setForegroundColor('#888888');

  doc.saveAndClose();

  /* az adatbázis mellé tesszük a Drive-on */
  try {
    var dbFajl = DriveApp.getFileById('1WnMwIimpAt5eUYmKTliEIoPLsH6Z4TqQ8mgzXjYACmM');
    var szulok = dbFajl.getParents();
    if (szulok.hasNext()) {
      var mappa = szulok.next();
      var f = DriveApp.getFileById(doc.getId());
      mappa.addFile(f);
      DriveApp.getRootFolder().removeFile(f);
    }
  } catch (err) { /* marad a Saját meghajtón */ }

  var url = doc.getUrl();
  Logger.log('Útmutató kész: ' + url);
  return url;
}


/* ======= 5. lépés: elméleti képzés formája (tantermi / e-learning) ======= */
function epites5() {
  var sor = [];
  [[EP_KFT, 'Kft'], [EP_EV, 'E.V.']].forEach(function (p) {
    var ss = SpreadsheetApp.openById(p[0]);
    var sh = ss.getSheetByName('Tanulók');
    if (!sh) { sor.push(p[1] + ': nincs Tanulók fül'); return; }
    var utolso = sh.getLastColumn();
    var fejlec = sh.getRange(1, 1, 1, utolso).getValues()[0];
    var hol = fejlec.indexOf('Elméleti képzés formája') + 1;
    var uj = false;
    if (!hol) {
      hol = utolso + 1; uj = true;
      if (sh.getMaxColumns() < hol) sh.insertColumnsAfter(sh.getMaxColumns(), hol - sh.getMaxColumns());
    }
    sh.getRange(1, hol).setValue('Elméleti képzés formája')
      .setFontWeight('bold').setWrap(true).setBackground('#E4F3EB')
      .setNote('Tanteremben vagy e-learningen tanulja az elméletet. '
             + 'A kimutatásban külön szűrhető. Vegyes = mindkettő volt.');
    sh.setColumnWidth(hol, 180);
    sh.getRange(2, hol, sh.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['E-learning', 'Tantermi', 'Vegyes', 'Nem releváns'], true)
        .setAllowInvalid(false).build());
    sor.push(p[1] + ': oszlop ' + (uj ? 'létrehozva' : 'már megvolt') + ' (' + hol + '. oszlop)');
  });
  /* a Nézetek lap újraépül, hogy belekerüljön az új szűrő */
  try {
    var db = SpreadsheetApp.openById('1WnMwIimpAt5eUYmKTliEIoPLsH6Z4TqQ8mgzXjYACmM');
    var n = db.getSheetByName('Nézetek');
    if (n) { db.deleteSheet(n); sor.push('Nézetek lap törölve – az epites3 újraépíti.'); }
  } catch (err) { sor.push('Nézetek: ' + err); }
  var out = sor.join('\n'); Logger.log(out); return out;
}


/* ======= 6. lépés: céges befizetés szétosztása – az adatbázis oldala ======= */
var CEGBEF_FEJ = ['Dátum', 'Megrendelő cég', 'Bizonylat / számla száma', 'Befizetett összeg (Ft)',
  'Szolgáltatási szerződés (iktatószám)', 'Szétosztva', 'Megjegyzés'];

function epites6() {
  var sor = [];
  [[EP_KFT, 'Kft'], [EP_EV, 'E.V.']].forEach(function (p) {
    var ss = SpreadsheetApp.openById(p[0]);

    /* a) két új oszlop a Tanulók lap végére */
    var sh = ss.getSheetByName('Tanulók');
    if (sh) {
      var utolso = sh.getLastColumn();
      var fejlec = sh.getRange(1, 1, 1, utolso).getValues()[0];
      [['Céges befizetés (Ft)', '#,##0', 150,
        'A cég befizetéséből erre a tanulóra jutó rész. A program tölti ki, kézzel ne írd át.'],
       ['Céges befizetés dátuma', 'yyyy. mm. dd.', 150,
        'Melyik céges befizetésből származik. A program tölti ki.']
      ].forEach(function (o) {
        var hol = fejlec.indexOf(o[0]) + 1;
        if (!hol) {
          utolso++; hol = utolso; fejlec.push(o[0]);
          if (sh.getMaxColumns() < hol) sh.insertColumnsAfter(sh.getMaxColumns(), hol - sh.getMaxColumns());
        }
        sh.getRange(1, hol).setValue(o[0]).setFontWeight('bold').setWrap(true)
          .setBackground('#EFEFEF').setNote(o[3]);
        sh.setColumnWidth(hol, o[2]);
        sh.getRange(2, hol, sh.getMaxRows() - 1, 1).setNumberFormat(o[1]);
      });
      sor.push(p[1] + ': Tanulók lap – céges befizetés oszlopok rendben');
    }

    /* b) "Céges befizetések" lap */
    var cb = ss.getSheetByName('Céges befizetések'), uj = false;
    if (!cb) { cb = ss.insertSheet('Céges befizetések'); uj = true; }
    cb.getRange(1, 1, 1, CEGBEF_FEJ.length).setValues([CEGBEF_FEJ])
      .setFontWeight('bold').setBackground('#0E8A52').setFontColor('#FFFFFF').setWrap(true);
    cb.setFrozenRows(1);
    [110, 230, 180, 160, 230, 190, 280].forEach(function (w, i) { cb.setColumnWidth(i + 1, w); });
    var sorokDb = Math.max(cb.getMaxRows(), 200);
    if (cb.getMaxRows() < sorokDb) cb.insertRowsAfter(cb.getMaxRows(), sorokDb - cb.getMaxRows());
    cb.getRange(2, 1, sorokDb - 1, 1).setNumberFormat('yyyy. mm. dd.');
    cb.getRange(2, 4, sorokDb - 1, 1).setNumberFormat('#,##0');
    var cegLap = ss.getSheetByName('Cégek');
    if (cegLap) cb.getRange(2, 2, sorokDb - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(cegLap.getRange('A2:A200'), true)
        .setAllowInvalid(false).build());
    cb.getRange(2, 6, sorokDb - 1, 1).setBackground('#EFEFEF');
    if (uj) cb.getRange('A1').setNote('Amikor egy cégtől átutalás érkezik, írj ide egy sort. '
      + 'Utána: ZOOM menü → „Céges befizetés szétosztása”. A program szétosztja a dolgozói között.');
    sor.push(p[1] + ': Céges befizetések lap ' + (uj ? 'létrehozva' : 'frissítve'));
  });
  var out = sor.join('\n'); Logger.log(out); return out;
}


/* ================= ÉLŐ KIMUTATÁS – weboldal =================
   A webcímet a Telepítés > Új telepítés > Webalkalmazás adja.
   Az oldal minden megnyitáskor frissen olvassa a Tanulók lapot. */
/* Akik megnyithatják az élő kimutatást. Bővítés: új sor a listába, utána új verzió kiadása. */
var ENGEDELYEZETT = [
  'halaszlaci78@gmail.com',   // László
  'zoomadripenzugy@gmail.com' // Adrienn
];

function doGet(){
  var ki = '';
  try { ki = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (!ki) { try { ki = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {} }
  var szabad = false;
  for (var i = 0; i < ENGEDELYEZETT.length; i++) {
    if (ENGEDELYEZETT[i].toLowerCase() === ki) { szabad = true; break; }
  }
  console.log('doGet: effektiv=' + (function(){try{return Session.getEffectiveUser().getEmail()}catch(e){return 'HIBA '+e}})() + ' | aktiv=' + (function(){try{return Session.getActiveUser().getEmail()}catch(e){return 'HIBA '+e}})() + ' | ki=' + ki + ' | szabad=' + szabad);
  if (!szabad) return nincsJog_(ki);
  var html = HtmlService.createHtmlOutputFromFile('oldal').getContent();
  html = html.split('%%ADATOK%%').join(JSON.stringify(zoomAdatok_()));
  return HtmlService.createHtmlOutput(html)
    .setTitle('ZOOM Adatbázis – élő')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function zoomAdatok_(){
  var ss = getDb_();
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Budapest';
  var sh = ss.getSheetByName('Tanulók');
  var out = [];
  var D = function(x){
    if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    x = String(x == null ? '' : x).trim();
    return x ? x.slice(0,10) : '';
  };
  var N = function(x){ var n = Number(x); return isNaN(n) ? 0 : n; };
  if (sh && sh.getLastRow() > 1){
    var v = sh.getRange(2, 1, sh.getLastRow()-1, 43).getValues();
    for (var i = 0; i < v.length; i++){
      var r = v[i];
      if (String(r[0] || '').trim()) continue;
      var nev = String(r[1] || '').trim();
      if (!nev) continue;
      var cs = String(r[31] || '').trim();
      var vkesz = D(r[36]), igeny = D(r[37]), kartya = D(r[38]);
      var kall = kartya ? 'Kártya megérkezett'
               : igeny  ? 'Igényelve, kártyára vár'
               : vkesz  ? 'Igényelhető – még nincs igényelve'
               : (String(r[32] || '').trim() === 'GKI' ? 'Vizsgák folyamatban' : '');
      var kvisel = String(r[33] || '').trim();
      var ceg = String(r[34] || '').trim();
      if (!ceg && kvisel === 'Céges') ceg = '(cég nincs megadva)';
      out.push({
        nev: nev,
        kepzo: String(r[2] || '').trim(),
        t: String(r[3] || '').trim(),
        cs: cs,
        al: String(r[32] || '').trim(),
        ar: N(r[6]),
        bef: N(r[7]),
        okt: String(r[9] || '').trim(),
        d: D(r[15]) || D(r[16]),
        szerz: D(r[16]),
        mod: String(r[24] || '').trim(),
        ceg: ceg,
        kvisel: kvisel,
        elm: String(r[40] || '').trim(),
        vkesz: vkesz,
        igeny: igeny,
        kartya: kartya,
        posta: String(r[39] || '').trim(),
        dbef: D(r[42]),
        kallapot: kall
      });
    }
  }
  return {
    ma: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    frissitve: Utilities.formatDate(new Date(), tz, 'yyyy. MM. dd. HH:mm'),
    tanulok: out,
    nezetek: zoomNezetek_(ss),
    katalogus: katalogus_()
  };
}

function zoomNezetek_(ss){
  try{
    var sh = ss.getSheetByName('Nézetek');
    if (!sh || sh.getLastRow() < 2) return [];
    var v = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < v.length; i++){
      var nev = String(v[i][0] || '').trim();
      if (!nev) continue;
      for (var j = 1; j < v[i].length; j++){
        var c = String(v[i][j] || '').trim();
        if (c.charAt(0) === '{'){
          try { out.push({ nev: nev, f: JSON.parse(c) }); } catch(e){}
          break;
        }
      }
    }
    return out;
  } catch(e){ return []; }
}

/* Gyors ellenőrzés a szerkesztőből: hány tanulót lát az oldal */
function eloOldalProba(){
  var a = zoomAdatok_();
  var cs = {};
  a.tanulok.forEach(function(r){ cs[r.cs || '(nincs csoport)'] = (cs[r.cs || '(nincs csoport)'] || 0) + 1; });
  var szerz = 0, bef = 0;
  a.tanulok.forEach(function(r){ szerz += r.ar; bef += r.bef; });
  console.log('Tanuló: ' + a.tanulok.length + ' | szerződött: ' + szerz + ' | befizetett: ' + bef);
  console.log('Csoportok: ' + JSON.stringify(cs));
  console.log('Nézetek: ' + a.nezetek.length);
  console.log('Első sor: ' + JSON.stringify(a.tanulok[0]));
  return a.tanulok.length;
}


/* A ZOOM menü felrakása a ZOOM Adatbázis (csak iroda) táblázatra – egyszer kell lefuttatni. */
function menuTelepites() {
  var id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) { console.log('Nincs DB_ID beállítva.'); return 'nincs DB_ID'; }
  var volt = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'onOpen' && String(t.getEventType()) === 'ON_OPEN';
  }).length;
  if (volt) { console.log('A menü már be van állítva.'); return 'mar volt'; }
  ScriptApp.newTrigger('onOpen').forSpreadsheet(id).onOpen().create();
  console.log('Kész: a ZOOM menü mostantól megjelenik a táblázat megnyitásakor.');
  return 'letrehozva';
}

/* A próba/teszt nevű sorok kizárása. Nem töröl a forrás táblákból,
   csak a Kizárt lapra teszi őket – a Kizárt lapról kivéve visszajönnek. */
function probaTorles() {
  var db = getDb_();
  var sh = db.getSheetByName('Tanulók');
  if (!sh || sh.getLastRow() < 2) { console.log('Nincs Tanulók lap.'); return; }
  var kiz = db.getSheetByName('Kizárt') || db.insertSheet('Kizárt');
  if (kiz.getLastRow() === 0) kiz.getRange(1, 1, 1, 3).setValues([['Azonosító', 'Név', 'Kizárva']]);
  var v = sh.getDataRange().getValues(), h = {};
  v[0].forEach(function (x, i) { h[String(x).trim()] = i; });
  var volt = {};
  if (kiz.getLastRow() > 1) {
    kiz.getRange(2, 1, kiz.getLastRow() - 1, 1).getValues()
       .forEach(function (r) { volt[String(r[0]).trim()] = 1; });
  }
  var add = [], nevek = [];
  for (var i = 1; i < v.length; i++) {
    var nev = String(v[i][h['Név']] || '').trim();
    var id  = String(v[i][h['Azonosító']] || '').trim();
    if (!nev || !id || volt[id]) continue;
    if (!/^\s*(pr[óo]ba|teszt)/i.test(nev)) continue;
    add.push([id, nev, new Date()]);
    nevek.push(nev + ' (' + id + ')');
  }
  if (add.length) kiz.getRange(kiz.getLastRow() + 1, 1, add.length, 3).setValues(add);
  console.log('Kizárva: ' + add.length + ' tétel — ' + (nevek.join(' | ') || 'nincs új'));
  return nevek;
}

/* Akinek nincs joga: semmilyen adatot nem lát. */
function nincsJog_(ki) {
  var o = '<!doctype html><html lang="hu"><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<title>Nincs hozzáférés</title><style>'
    + 'html,body{margin:0;height:100%;background:#EEF3F8;font-family:"Segoe UI",system-ui,sans-serif;color:#10202F}'
    + '.d{max-width:420px;margin:0 auto;padding:56px 24px;text-align:center}'
    + '.c{background:#fff;border:1px solid #D8E2EC;border-radius:16px;padding:30px 26px;'
    + 'box-shadow:0 10px 30px -18px rgba(16,32,47,.4)}'
    + 'h1{font-size:19px;margin:0 0 10px}p{color:#4B5E70;font-size:14px;line-height:1.6;margin:0 0 8px}'
    + 'code{background:#F6F9FC;border:1px solid #D8E2EC;border-radius:6px;padding:2px 7px;font-size:12.5px}'
    + '</style></head><body><div class="d"><div class="c">'
    + '<h1>Ehhez a kimutatáshoz nincs hozzáférésed</h1>'
    + '<p>Az oldalt csak a ZOOM Autósiskola irodája nyithatja meg.</p>'
    + (ki ? '<p>Beléptél mint <code>' + ki.replace(/[<>&]/g, '') + '</code>. Ha ez nem a munkahelyi fiókod, lépj át arra.</p>'
         : '<p>Úgy tűnik, nem vagy bejelentkezve Google-fiókkal.</p>')
    + '</div></div></body></html>';
  return HtmlService.createHtmlOutput(o).setTitle('Nincs hozzáférés');
}

/* A teljes képzéskínálat (főcsoport / alcsoport / képzés), hogy a kimutatás
   szűrői akkor is mutassák pl. a gépkezelőt, ha még nincs rá tanuló.
   Egy órára gyorsítótárazva, hogy ne lassítsa az oldalt. */
function katalogus_() {
  var c = CacheService.getScriptCache();
  try { var k = c.get('katalogus_v1'); if (k) return JSON.parse(k); } catch (e) {}
  var out = [];
  try {
    var map = csoportTerkep_();
    Object.keys(map).forEach(function (n) {
      out.push({ t: n, cs: map[n][0], al: map[n][1] });
    });
    c.put('katalogus_v1', JSON.stringify(out), 3600);
  } catch (e) {}
  return out;
}

/* Ellenőrzés: látja-e a program a bejelentkezett fiókot */
function engedelyProba() {
  console.log('Bejelentkezett: ' + Session.getActiveUser().getEmail() + ' | effektív: ' + Session.getEffectiveUser().getEmail());
}

