/* Kernlogik: Zeitrechnung, Excel-Export (.xlsx im Layout der Vorlage) und ZIP-Writer.
 * Reines JavaScript ohne Abhängigkeiten, läuft im Browser und in Node (zum Testen).
 *
 * Datenmodell eines Tages: { s: Startminute|null, e: Endminute|null, p: Pause in Minuten, n: Stichworte }
 * (Minuten seit Mitternacht, z. B. 7:30 = 450)
 */
(function (root) {
  'use strict';

  var MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
  var DAYS = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];

  function two(n) { return n < 10 ? '0' + n : String(n); }

  /** "7:30", "07:30", "7.30" -> Minuten, sonst null */
  function parseTime(s) {
    var m = /^\s*(\d{1,2})\s*[:.]\s*(\d{2})\s*$/.exec(s || '');
    if (!m) return null;
    var h = parseInt(m[1], 10), mi = parseInt(m[2], 10);
    if (h > 23 || mi > 59) return null;
    return h * 60 + mi;
  }
  function hhmm(min) { return two(Math.floor(min / 60)) + ':' + two(min % 60); }
  function dur(min) { min = Math.max(min, 0); return Math.floor(min / 60) + ':' + two(min % 60); }

  function gross(e) {
    if (!e || e.s == null || e.e == null) return 0;
    return Math.max(e.e - e.s, 0);
  }
  function net(e) { return Math.max(gross(e) - ((e && e.p) || 0), 0); }

  function daysInMonth(y, m) { return new Date(y, m, 0).getDate(); }
  /** 0 = Montag ... 6 = Sonntag */
  function weekday(y, m, d) { return (new Date(y, m - 1, d).getDay() + 6) % 7; }
  function isWeekend(y, m, d) { return weekday(y, m, d) >= 5; }
  function key(y, m, d) { return y + '-' + two(m) + '-' + two(d); }
  /** Excel-Seriennummer (Tage seit 1899-12-30) */
  function serial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }

  // ---------------------------------------------------------------- XLSX

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function frac(min) { return min === 0 ? '0' : String(min / 1440); }

  function str(ref, style, text) {
    return '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
  }
  function num(ref, style, v) { return '<c r="' + ref + '" s="' + style + '"><v>' + v + '</v></c>'; }
  function blank(ref, style) { return '<c r="' + ref + '" s="' + style + '"/>'; }
  function formula(ref, style, f, cached) {
    return '<c r="' + ref + '" s="' + style + '"><f>' + esc(f) + '</f><v>' + cached + '</v></c>';
  }

  // Stil-IDs aus styles.xml der Vorlage
  var S = {
    LABEL: 1, COMPANY: 2, EMP: 7, MONTH_NUM: 12, MONTH_NAME: 13, YEAR: 14,
    H_A: 15, H_B: 16, H_C: 17, H_DE: 18, H_F: 19, H_GHJ: 20, H_I: 21,
    FIRST_A: 22, FIRST_B: 23, FIRST_C: 24, TIME: 25, FIRST_DUR: 26, FIRST_PROJ: 27, NOTE: 28,
    A: 30, B: 31, C: 32, DUR: 33, PROJ: 34, PAUSE: 35,
    TOT_F: 46, TOT_G: 47, TOT_I: 48
  };

  function buildSheet(meta, days) {
    var n = days.length, firstRow = 6, lastRow = firstRow + n - 1, totalRow = lastRow + 2;
    var b = [];
    b.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    b.push('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    b.push('<sheetPr><pageSetUpPr/></sheetPr><sheetViews><sheetView workbookViewId="0"/></sheetViews>');
    b.push('<sheetFormatPr defaultColWidth="12.63" defaultRowHeight="15.0"/>');
    b.push('<cols>' +
      '<col min="1" max="1" width="7.25" customWidth="1"/>' +
      '<col min="2" max="2" width="7.38" customWidth="1"/>' +
      '<col min="3" max="3" width="24" customWidth="1"/>' +
      '<col min="4" max="5" width="9.0" customWidth="1"/>' +
      '<col min="6" max="6" width="16.25" customWidth="1"/>' +
      '<col min="7" max="7" width="22.5" customWidth="1"/>' +
      '<col min="8" max="8" width="8.25" customWidth="1"/>' +
      '<col min="9" max="9" width="10.25" customWidth="1"/>' +
      '<col min="10" max="10" width="34.13" customWidth="1"/>' +
      '<col min="11" max="15" width="2.63" customWidth="1"/></cols>');
    b.push('<sheetData>');
    b.push('<row r="1" ht="18.0" customHeight="1">' + str('A1', S.LABEL, 'Firma:') + str('C1', S.COMPANY, meta.company) + '</row>');
    b.push('<row r="2" ht="18.0" customHeight="1">' + str('A2', S.LABEL, 'Mitarbeiter:') + str('C2', S.EMP, meta.employee) + '</row>');
    b.push('<row r="3" ht="18.0" customHeight="1">' + str('A3', S.EMP, 'Monat:') + num('B3', S.MONTH_NUM, meta.month) +
      str('C3', S.MONTH_NAME, MONTHS[meta.month - 1]) + num('E3', S.YEAR, meta.year) + '</row>');
    b.push('<row r="4" ht="15.75" customHeight="1"/>');
    b.push('<row r="5">' + str('A5', S.H_A, 'Datum       ') + blank('B5', S.H_B) + str('C5', S.H_C, 'Tag') +
      str('D5', S.H_DE, 'Startzeit') + str('E5', S.H_DE, 'Endzeit') + str('F5', S.H_F, 'Dauer') +
      str('G5', S.H_GHJ, 'Dauer(rel.)') + str('H5', S.H_GHJ, 'Projekt') + str('I5', S.H_I, 'Pausen') +
      str('J5', S.H_GHJ, 'Stichworte') + '</row>');

    var firstSerial = serial(meta.year, meta.month, 1);
    var sumGross = 0, sumNet = 0, sumPause = 0;
    for (var i = 0; i < n; i++) {
      var e = days[i] || {}, r = firstRow + i, first = i === 0;
      var ok = e.s != null && e.e != null;
      var s = ok ? e.s : 0, en = ok ? e.e : 0;
      if (ok && en < s) { ok = false; s = 0; en = 0; }
      var g = en - s;
      var p = ok ? Math.min(Math.max(e.p || 0, 0), g) : 0;
      var nt = g - p;
      sumGross += g; sumPause += p; sumNet += nt;
      var row = '<row r="' + r + '">';
      row += first ? formula('A' + r, S.FIRST_A, 'DATE(E3,B3,1)', firstSerial)
                   : formula('A' + r, S.A, 'A' + (r - 1) + '+1', firstSerial + i);
      row += blank('B' + r, first ? S.FIRST_B : S.B) + blank('C' + r, first ? S.FIRST_C : S.C);
      row += num('D' + r, S.TIME, frac(s)) + num('E' + r, S.TIME, frac(en));
      row += formula('F' + r, first ? S.FIRST_DUR : S.DUR, 'SUM(E' + r + '-D' + r + ')', frac(g));
      row += formula('G' + r, first ? S.FIRST_DUR : S.DUR, 'SUM(F' + r + '-I' + r + ')', frac(nt));
      row += str('H' + r, first ? S.FIRST_PROJ : S.PROJ, meta.project);
      row += num('I' + r, first ? S.TIME : S.PAUSE, frac(p));
      row += e.n ? str('J' + r, S.NOTE, e.n) : blank('J' + r, S.NOTE);
      row += '</row>';
      b.push(row);
    }
    b.push('<row r="' + totalRow + '" ht="28.5" customHeight="1">' +
      formula('F' + totalRow, S.TOT_F, 'SUM(F' + firstRow + ':F' + lastRow + ')', frac(sumGross)) +
      formula('G' + totalRow, S.TOT_G, 'SUM(G' + firstRow + ':G' + lastRow + ')', frac(sumNet)) +
      formula('I' + totalRow, S.TOT_I, 'SUM(I' + firstRow + ':I' + lastRow + ')', frac(sumPause)) + '</row>');
    b.push('</sheetData>');
    b.push('<mergeCells count="' + (n + 3) + '"><mergeCell ref="A1:B1"/><mergeCell ref="A2:B2"/><mergeCell ref="C3:D3"/>');
    for (var rr = firstRow; rr <= lastRow; rr++) b.push('<mergeCell ref="A' + rr + ':C' + rr + '"/>');
    b.push('</mergeCells>');
    b.push('<pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.0" footer="0.0"/>');
    b.push('<pageSetup orientation="landscape"/></worksheet>');
    return b.join('');
  }

  // ---------------------------------------------------------------- ZIP (ohne Kompression)

  var crcTable = null;
  function crc32(buf) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        crcTable[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) crc = crcTable[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function utf8(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    return new Uint8Array(Buffer.from(s, 'utf8'));
  }

  /** files: [{name, data:Uint8Array}] -> Uint8Array (ZIP, Methode "stored") */
  function zip(files, when) {
    var d = when || new Date();
    var dosDate = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    var dosTime = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    var chunks = [], central = [], offset = 0, i, f;
    for (i = 0; i < files.length; i++) {
      f = files[i];
      var name = utf8(f.name), crc = crc32(f.data), size = f.data.length;
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true); lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, size, true); lh.setUint32(22, size, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), name, f.data);

      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, dosTime, true);
      ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true); ch.setUint32(20, size, true);
      ch.setUint32(24, size, true); ch.setUint16(28, name.length, true); ch.setUint16(30, 0, true);
      ch.setUint16(32, 0, true); ch.setUint16(34, 0, true); ch.setUint16(36, 0, true);
      ch.setUint32(38, 0, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + size;
    }
    var cdSize = 0;
    for (i = 0; i < central.length; i++) cdSize += central[i].length;
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true);
    end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
    var all = chunks.concat(central, [new Uint8Array(end.buffer)]);
    var total = 0;
    for (i = 0; i < all.length; i++) total += all[i].length;
    var out = new Uint8Array(total), pos = 0;
    for (i = 0; i < all.length; i++) { out.set(all[i], pos); pos += all[i].length; }
    return out;
  }

  /**
   * meta: {company, employee, project, month (1-12), year}
   * days: Array je Kalendertag des Monats (1..n) mit {s,e,p,n}
   * assets: {styles, theme} (Text aus der Vorlage)
   */
  function buildXlsx(meta, days, assets) {
    var sheetName = MONTHS[meta.month - 1] + '_' + meta.year;
    var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var files = [
      { name: '[Content_Types].xml', data: utf8(head +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>') },
      { name: '_rels/.rels', data: utf8(head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
      { name: 'xl/workbook.xml', data: utf8(head +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<workbookPr/><bookViews><workbookView/></bookViews>' +
        '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets>' +
        '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>') },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8(head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>') },
      { name: 'xl/styles.xml', data: utf8(assets.styles) },
      { name: 'xl/theme/theme1.xml', data: utf8(assets.theme) },
      { name: 'xl/worksheets/sheet1.xml', data: utf8(buildSheet(meta, days)) }
    ];
    return zip(files);
  }

  function namePart(s) {
    return String(s == null ? '' : s).trim().replace(/\s+/g, '_')
      .replace(/[\\\/:*?"<>|\u0000-\u001F\u007F]/g, '').replace(/_+/g, '_').replace(/^_|_$/g, '');
  }

  /** Vorgabe-Dateiname: Stunden_Projekt_Mitarbeiter_MM_JJJJ.xlsx (leere Teile entfallen) */
  function fileName(project, employee, month, year) {
    var p = namePart(project), e = namePart(employee);
    return 'Stunden' + (p ? '_' + p : '') + (e ? '_' + e : '') + '_' + two(month) + '_' + year + '.xlsx';
  }

  /** Eingabe bereinigen und ".xlsx" sicherstellen; null, wenn nichts übrig bleibt */
  function sanitizeFileName(input) {
    var t = String(input == null ? '' : input).trim();
    if (/\.xlsx$/i.test(t)) t = t.slice(0, -5);
    t = t.replace(/[\\\/:*?"<>|\u0000-\u001F\u007F]/g, '').trim().replace(/[. ]+$/, '');
    if (t.length > 120) t = t.slice(0, 120);
    return t ? t + '.xlsx' : null;
  }

  var api = {
    MONTHS: MONTHS, DAYS: DAYS, two: two, parseTime: parseTime, hhmm: hhmm, dur: dur,
    gross: gross, net: net, daysInMonth: daysInMonth, weekday: weekday, isWeekend: isWeekend,
    key: key, buildXlsx: buildXlsx, fileName: fileName, sanitizeFileName: sanitizeFileName, zip: zip, crc32: crc32
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.AZ = api;
})(typeof self !== 'undefined' ? self : this);
