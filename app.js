/* Oberfläche und Zustand der Arbeitszeit-Web-App. Daten liegen im localStorage des Geräts. */
(function () {
  'use strict';
  var AZ = window.AZ, ASSETS = window.XLSX_ASSETS;
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // ------------------------------------------------------------ Speicher
  var mem = {}; // Fallback, falls localStorage nicht verfügbar ist
  function load(k, def) {
    try {
      var v = localStorage.getItem(k);
      if (v == null) return k in mem ? mem[k] : def;
      return JSON.parse(v);
    } catch (e) { return k in mem ? mem[k] : def; }
  }
  function save(k, v) {
    mem[k] = v;
    try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* ignorieren */ }
  }

  var entries = load('az_entries', {});
  var settings = Object.assign({
    company: '', employee: '', project: '', // werden einmalig in den Einstellungen eingetragen (nur lokal gespeichert)
    defStart: 450, defEnd: 960, defPause: 30, theme: 'system', layout: 'auto'
  }, load('az_settings', {}));

  /** Design anwenden: 'light' / 'dark' setzen data-theme, 'system' folgt dem Gerät */
  function applyTheme() {
    var t = settings.theme, el = document.documentElement;
    if (t === 'light' || t === 'dark') el.setAttribute('data-theme', t); else el.removeAttribute('data-theme');
    var dark = t === 'dark' || (t !== 'light' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', dark ? '#101316' : '#1f6f8b');
  }
  applyTheme();
  if (window.matchMedia) {
    var mq = window.matchMedia('(prefers-color-scheme: dark)');
    var onMq = function () { applyTheme(); };
    if (mq.addEventListener) mq.addEventListener('change', onMq); else if (mq.addListener) mq.addListener(onMq);
  }
  // Laufende Erfassung: {date:'YYYY-MM-DD', pauseStart: ms|0, pauseAccum: ms} oder null
  var run = load('az_run', null);

  function saveEntries() { save('az_entries', entries); }
  function saveRun() { save('az_run', run); }

  // ------------------------------------------------------------ Hilfen
  var $ = function (id) { return document.getElementById(id); };
  function today() { var d = new Date(); return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() }; }
  function todayKey() { var t = today(); return AZ.key(t.y, t.m, t.d); }
  function nowMin() { var d = new Date(); return d.getHours() * 60 + d.getMinutes(); }
  function parseKey(k) { var p = k.split('-'); return { y: +p[0], m: +p[1], d: +p[2] }; }
  function isEmpty(e) { return !e || (e.s == null && e.e == null && !e.p && !e.n); }

  // ------------------------------------------------------------ Ansicht: Mobil / Desktop
  var ICONS = {
    auto: '<rect x="2" y="5" width="14" height="10" rx="1.5"/><rect x="13" y="9" width="8" height="12" rx="1.5"/>',
    mobile: '<rect x="7" y="2.5" width="10" height="19" rx="2"/><path d="M11 18.5h2"/>',
    desktop: '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/>'
  };
  var VIEW_LABEL = { auto: 'Auto', mobile: 'Mobil', desktop: 'Desktop' };
  var DESKTOP_MIN_WIDTH = 900; // ab dieser Fensterbreite (CSS-Pixel) schaltet "Auto" auf Desktop

  function effectiveView() {
    if (settings.layout === 'desktop') return 'desktop';
    if (settings.layout === 'mobile') return 'mobile';
    return window.innerWidth >= DESKTOP_MIN_WIDTH ? 'desktop' : 'mobile';
  }

  function isPhone() { return Math.min(window.screen.width || 9999, window.screen.height || 9999) < 700; }
  /** Erzwungene Desktop-Ansicht auf einem Handy: Die Seite wird breit gerechnet (1100 px) und verkleinert dargestellt. */
  function wantsWide(layout) { return layout === 'desktop' && isPhone(); }

  function applyView() {
    var lay = VIEW_LABEL[settings.layout] ? settings.layout : 'auto';
    var vp = document.getElementById('vp');
    if (vp) vp.setAttribute('content', wantsWide(lay) ? 'width=1100' : 'width=device-width, initial-scale=1, viewport-fit=cover');
    document.documentElement.setAttribute('data-view', effectiveView());
    $('viewLbl').textContent = VIEW_LABEL[lay];
    $('icoView').innerHTML = ICONS[lay];
    var tip = 'Ansicht: ' + VIEW_LABEL[lay] + ' – zum Umschalten tippen';
    $('bView').title = tip;
    $('bView').setAttribute('aria-label', tip);
  }
  applyView();

  var t0 = today();
  var view = { y: t0.y, m: t0.m };

  var toastTimer;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  /** Meldung mit Schaltflächen: buttons = [{label, ok, fn}] */
  function message(title, text, buttons) {
    $('msgTitle').textContent = title;
    $('msgText').textContent = text;
    var box = $('msgBtns'); box.innerHTML = '';
    buttons.forEach(function (b) {
      var el = document.createElement('button');
      el.type = 'button'; el.textContent = b.label;
      if (b.ok) el.className = 'ok';
      el.onclick = function () { $('dlgMsg').close(); if (b.fn) b.fn(); };
      box.appendChild(el);
    });
    $('dlgMsg').showModal();
  }

  function hms(ms) {
    var s = Math.max(Math.floor(ms / 1000), 0);
    return Math.floor(s / 3600) + ':' + AZ.two(Math.floor((s % 3600) / 60)) + ':' + AZ.two(s % 60);
  }

  // ------------------------------------------------------------ Start / Stopp / Pause
  function startWork() {
    var k = todayKey(), now = nowMin();
    var e = entries[k];
    if (!e) {
      e = { s: now, e: null, p: 0, n: '' };
    } else if (e.s == null) {
      e.s = now; e.e = null;
    } else if (e.e != null) {
      // Weiterarbeit am selben Tag: Lücke seit dem letzten Ende zählt als Pause
      var gap = now - e.e;
      if (gap > 0) e.p = (e.p || 0) + gap;
      e.e = null;
    }
    entries[k] = e; saveEntries();
    run = { date: k, pauseStart: 0, pauseAccum: 0 }; saveRun();
    var t = today(); view = { y: t.y, m: t.m };
    render(); scrollToToday();
  }

  function stopWork() {
    var k = run.date, e = entries[k] || { s: null, e: null, p: 0, n: '' };
    var accum = run.pauseAccum + (run.pauseStart ? Date.now() - run.pauseStart : 0);
    e.p = (e.p || 0) + Math.round(accum / 60000);
    var end = nowMin();
    if (todayKey() !== k) {
      end = 23 * 60 + 59;
      toast('Die Erfassung lief über Mitternacht – Ende auf 23:59 gesetzt, bitte prüfen.');
    }
    if (e.s == null) e.s = end;
    if (end < e.s) end = e.s;
    e.e = end;
    if (e.p > AZ.gross(e)) e.p = AZ.gross(e);
    entries[k] = e; saveEntries();
    run = null; saveRun();
    render();
  }

  function togglePause() {
    if (!run) return;
    if (run.pauseStart) { run.pauseAccum += Date.now() - run.pauseStart; run.pauseStart = 0; }
    else run.pauseStart = Date.now();
    saveRun(); tick();
  }

  function dayStartMs(k, min) {
    var p = parseKey(k);
    return new Date(p.y, p.m - 1, p.d, Math.floor(min / 60), min % 60, 0, 0).getTime();
  }

  function tick() {
    var st = $('status'), sub = $('sub'), bs = $('bStart'), bp = $('bPause');
    if (run) {
      var e = entries[run.date], paused = !!run.pauseStart, now = Date.now();
      if (paused) {
        var d = new Date(run.pauseStart);
        st.textContent = 'Pause seit ' + AZ.two(d.getHours()) + ':' + AZ.two(d.getMinutes());
      } else {
        st.textContent = 'Arbeitet seit ' + (e && e.s != null ? AZ.hhmm(e.s) : '?');
      }
      var startMs = e && e.s != null ? dayStartMs(run.date, e.s) : now;
      var pauseMs = run.pauseAccum + (paused ? now - run.pauseStart : 0) + ((e && e.p) || 0) * 60000;
      sub.textContent = 'Netto heute: ' + hms(now - startMs - pauseMs);
      bs.textContent = 'Stopp'; bs.className = 'big run';
      bp.textContent = paused ? 'Weiter' : 'Pause'; bp.className = 'big' + (paused ? ' resume' : '');
      bp.disabled = false;
    } else {
      bs.textContent = 'Start'; bs.className = 'big';
      bp.textContent = 'Pause'; bp.className = 'big'; bp.disabled = true;
      var te = entries[todayKey()];
      if (te && te.s != null && te.e != null) {
        st.textContent = 'Heute: ' + AZ.hhmm(te.s) + ' – ' + AZ.hhmm(te.e);
        sub.textContent = 'Netto ' + AZ.dur(AZ.net(te)) + ' h, Pause ' + (te.p || 0) + ' min';
      } else {
        st.textContent = 'Nicht gestartet';
        sub.textContent = 'Mit „Start“ die Arbeitszeit erfassen';
      }
    }
  }

  // ------------------------------------------------------------ Monatsansicht
  function render() {
    $('monthLabel').textContent = AZ.MONTHS[view.m - 1] + ' ' + view.y;
    var list = $('list'); list.innerHTML = '';
    var n = AZ.daysInMonth(view.y, view.m), tk = todayKey();
    var sumNet = 0, sumPause = 0, worked = 0;
    for (var d = 1; d <= n; d++) {
      var k = AZ.key(view.y, view.m, d), e = entries[k];
      var we = AZ.isWeekend(view.y, view.m, d);
      var isRun = run && run.date === k;
      var row = document.createElement('button');
      row.type = 'button';
      row.className = 'row' + (isRun ? ' run' : k === tk ? ' today' : we ? ' we' : '');
      row.dataset.key = k;
      var times, tcls = 't none', pauseTxt = '', hasStart = false, netTxt = '';
      if (e && e.s != null) {
        hasStart = true;
        if (e.e != null) {
          times = AZ.hhmm(e.s) + ' – ' + AZ.hhmm(e.e); tcls = 't';
          netTxt = AZ.dur(AZ.net(e)) + ' h';
          sumNet += AZ.net(e); sumPause += Math.min(e.p || 0, AZ.gross(e));
          if (AZ.net(e) > 0) worked++;
        } else if (isRun) { times = AZ.hhmm(e.s) + ' – läuft'; tcls = 't live'; }
        else { times = AZ.hhmm(e.s) + ' – ?'; tcls = 't open'; }
        pauseTxt = (e.p || 0) + ' min';
      } else times = '–';
      var note = e && e.n ? e.n : '';
      // Zellen: Datum | Zeit | Pause + Stichworte (auf dem Handy unter der Zeit, am Desktop eigene Spalten) | Netto
      row.innerHTML = '<div class="d"><span class="wd"></span><span class="dt"></span></div><div class="t"></div>' +
        '<div class="i"><span class="p"><span class="lbl">Pause </span><span class="pv"></span></span><span class="nt"></span></div><div class="n"></div>';
      row.querySelector('.wd').textContent = AZ.DAYS[AZ.weekday(view.y, view.m, d)];
      row.querySelector('.dt').textContent = AZ.two(d) + '.' + AZ.two(view.m) + '.';
      var tEl = row.querySelector('.t');
      tEl.className = tcls; tEl.textContent = times;
      row.querySelector('.pv').textContent = pauseTxt;
      row.querySelector('.p').hidden = !hasStart;
      row.querySelector('.nt').textContent = note;
      row.querySelector('.i').hidden = !hasStart && !note;
      row.querySelector('.n').textContent = netTxt;
      list.appendChild(row);
    }
    $('sum').innerHTML =
      '<span class="s"><small>Netto</small><b>' + AZ.dur(sumNet) + ' h</b></span>' +
      '<span class="s"><small>Pausen</small><b>' + AZ.dur(sumPause) + ' h</b></span>' +
      '<span class="s"><small>' + (worked === 1 ? 'Tag' : 'Tage') + '</small><b>' + worked + '</b></span>';
    tick();
  }

  function scrollToToday() {
    var t = today(), list = $('list');
    if (view.y === t.y && view.m === t.m) {
      var row = list.querySelector('[data-key="' + todayKey() + '"]');
      if (row) list.scrollTop = Math.max(row.offsetTop - list.offsetTop - row.offsetHeight * 2, 0);
    } else list.scrollTop = 0;
  }

  // ------------------------------------------------------------ Tag bearbeiten
  var editKey = null;

  function openDay(k) {
    editKey = k;
    var e = entries[k], p = parseKey(k);
    var runningDay = run && run.date === k;
    $('dayTitle').textContent = AZ.DAYS[AZ.weekday(p.y, p.m, p.d)] + ', ' + AZ.two(p.d) + '.' + AZ.two(p.m) + '.' + p.y;
    $('fOff').checked = !!(e && e.s == null && e.n);
    $('fStart').value = AZ.hhmm(e && e.s != null ? e.s : settings.defStart);
    $('fEnd').value = AZ.hhmm(e && e.e != null ? e.e : settings.defEnd);
    $('fPause').value = e && e.s != null ? (e.p || 0) : settings.defPause;
    $('fNote').value = e && e.n ? e.n : '';
    $('endWrap').hidden = !!runningDay;
    $('runHint').hidden = !runningDay;
    $('dayErr').textContent = '';
    syncOff();
    $('dlgDay').showModal();
  }
  function syncOff() { $('timeBox').hidden = $('fOff').checked; }

  function saveDay() {
    var k = editKey, runningDay = run && run.date === k;
    var note = $('fNote').value.trim();
    var e = entries[k] || { s: null, e: null, p: 0, n: '' };
    var err = $('dayErr');
    if ($('fOff').checked) {
      if (runningDay) { err.textContent = 'Die Erfassung läuft noch – bitte erst stoppen.'; return false; }
      e.s = null; e.e = null; e.p = 0; e.n = note;
    } else {
      var s = AZ.parseTime($('fStart').value), en = AZ.parseTime($('fEnd').value);
      var pause = parseInt($('fPause').value, 10);
      if (s == null) { err.textContent = 'Bitte eine Startzeit angeben.'; return false; }
      if (isNaN(pause) || pause < 0) { err.textContent = 'Bitte die Pause in Minuten angeben (0 für keine).'; return false; }
      if (!runningDay) {
        if (en == null) { err.textContent = 'Bitte eine Endzeit angeben.'; return false; }
        if (en <= s) { err.textContent = 'Das Ende muss nach dem Start liegen.'; return false; }
        if (pause > en - s) { err.textContent = 'Die Pause ist länger als die Arbeitszeit.'; return false; }
      }
      e.s = s; e.e = runningDay ? null : en; e.p = pause; e.n = note;
    }
    if (isEmpty(e)) delete entries[k]; else entries[k] = e;
    saveEntries(); $('dlgDay').close(); render();
    return true;
  }

  // ------------------------------------------------------------ Einstellungen
  function openSettings() {
    $('sCompany').value = settings.company;
    $('sEmployee').value = settings.employee;
    $('sProject').value = settings.project;
    $('sStart').value = AZ.hhmm(settings.defStart);
    $('sEnd').value = AZ.hhmm(settings.defEnd);
    $('sPause').value = settings.defPause;
    $('sTheme').value = settings.theme || 'system';
    $('sLayout').value = VIEW_LABEL[settings.layout] ? settings.layout : 'auto';
    $('setErr').textContent = '';
    $('dlgSet').showModal();
  }
  function saveSettings() {
    var s = AZ.parseTime($('sStart').value), e = AZ.parseTime($('sEnd').value), p = parseInt($('sPause').value, 10);
    var err = $('setErr');
    if (s == null || e == null) { err.textContent = 'Bitte Zeiten im Format HH:MM angeben, z. B. 07:30.'; return false; }
    if (e <= s) { err.textContent = 'Das Standard-Ende muss nach dem Standard-Start liegen.'; return false; }
    if (isNaN(p) || p < 0) { err.textContent = 'Bitte die Pause in Minuten angeben.'; return false; }
    var wasWide = wantsWide(settings.layout);
    settings = { company: $('sCompany').value.trim(), employee: $('sEmployee').value.trim(), project: $('sProject').value.trim(),
      defStart: s, defEnd: e, defPause: p, theme: $('sTheme').value, layout: $('sLayout').value };
    save('az_settings', settings);
    if (wasWide !== wantsWide(settings.layout)) { location.reload(); return true; }
    applyTheme();
    applyView();
    $('dlgSet').close(); toast('Gespeichert');
    return true;
  }

  // ------------------------------------------------------------ Werktage füllen
  function fillCandidates() {
    var res = [], t = today(), n = AZ.daysInMonth(view.y, view.m);
    for (var d = 1; d <= n; d++) {
      var k = AZ.key(view.y, view.m, d);
      if (k > todayKey()) break;
      if (AZ.isWeekend(view.y, view.m, d) || entries[k]) continue;
      if (k === todayKey() && (run || nowMin() < settings.defEnd)) continue;
      res.push(k);
    }
    return res;
  }
  function fillWeekdays() {
    var cand = fillCandidates();
    if (!cand.length) { toast('Keine leeren Werktage zum Füllen in diesem Monat.'); return; }
    message('Werktage füllen',
      cand.length + ' leere Werktage (Mo–Fr) mit der Standardzeit ' + AZ.hhmm(settings.defStart) + '–' + AZ.hhmm(settings.defEnd) +
      ' und ' + settings.defPause + ' min Pause füllen?\n\nUrlaub und Feiertage danach einfach einzeln antippen und ändern.',
      [{ label: 'Abbrechen' }, { label: 'Füllen', ok: true, fn: function () {
        cand.forEach(function (k) { entries[k] = { s: settings.defStart, e: settings.defEnd, p: settings.defPause, n: '' }; });
        saveEntries(); render();
      } }]);
  }

  // ------------------------------------------------------------ Export
  function exportMonth() {
    var n = AZ.daysInMonth(view.y, view.m), incomplete = 0;
    for (var d = 1; d <= n; d++) {
      var e = entries[AZ.key(view.y, view.m, d)];
      if (e && e.s != null && e.e == null) incomplete++;
    }
    if (incomplete) {
      message('Unvollständige Tage',
        incomplete + (incomplete === 1 ? ' Tag hat' : ' Tage haben') + ' einen Start, aber kein Ende (oder die Erfassung läuft noch). Diese Tage werden mit 0:00 exportiert.',
        [{ label: 'Abbrechen' }, { label: 'Trotzdem exportieren', ok: true, fn: askName }]);
    } else askName();
  }

  /** Vorgabename aus den Einstellungen: Stunden_Projekt_Mitarbeiter_MM_JJJJ (ohne .xlsx) */
  function defaultName() {
    return AZ.fileName(settings.project, settings.employee, view.m, view.y).replace(/\.xlsx$/, '');
  }

  /** Dialog, in dem der Dateiname vor dem Export geändert werden kann */
  function askName() {
    var el = $('nName');
    el.value = defaultName();
    $('nameErr').textContent = '';
    $('dlgName').showModal();
    el.focus();
    el.setSelectionRange(el.value.length, el.value.length);
  }

  function doExport(name) {
    var n = AZ.daysInMonth(view.y, view.m), days = [];
    for (var d = 1; d <= n; d++) days.push(entries[AZ.key(view.y, view.m, d)] || {});
    var meta = { company: settings.company, employee: settings.employee, project: settings.project, month: view.m, year: view.y };
    var bytes;
    try {
      bytes = AZ.buildXlsx(meta, days, ASSETS);
    } catch (err) {
      message('Export fehlgeschlagen', String(err && err.message || err), [{ label: 'OK', ok: true }]);
      return;
    }
    var blob = new Blob([bytes], { type: XLSX_MIME });
    var file;
    try { file = new File([blob], name, { type: XLSX_MIME }); } catch (e) { file = null; }

    function download() {
      var url = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 4000);
      toast('Datei „' + name + '“ wird gespeichert');
    }

    var canShare = !!(file && navigator.canShare && navigator.canShare({ files: [file] }) && navigator.share);
    if (canShare) {
      message('Export fertig', name + '\n\nTeilen öffnet das Menü, in dem du „In Dateien sichern“, Mail, WhatsApp o. Ä. wählen kannst.',
        [{ label: 'Herunterladen', fn: download },
         { label: 'Teilen / Speichern', ok: true, fn: function () {
            navigator.share({ files: [file], title: name }).catch(function (e) { if (e && e.name !== 'AbortError') download(); });
          } }]);
    } else download();
  }

  // ------------------------------------------------------------ Verdrahtung
  $('bStart').onclick = function () { run ? stopWork() : startWork(); };
  $('bPause').onclick = togglePause;
  $('bPrev').onclick = function () { view.m--; if (view.m < 1) { view.m = 12; view.y--; } render(); $('list').scrollTop = 0; };
  $('bNext').onclick = function () { view.m++; if (view.m > 12) { view.m = 1; view.y++; } render(); $('list').scrollTop = 0; };
  $('monthLabel').onclick = function () { var t = today(); view = { y: t.y, m: t.m }; render(); scrollToToday(); };
  $('list').onclick = function (ev) {
    var row = ev.target.closest && ev.target.closest('.row');
    if (row) openDay(row.dataset.key);
  };
  $('fOff').onchange = syncOff;
  $('bDefault').onclick = function () {
    $('fStart').value = AZ.hhmm(settings.defStart); $('fEnd').value = AZ.hhmm(settings.defEnd); $('fPause').value = settings.defPause;
  };
  $('fDay').onsubmit = function (ev) { ev.preventDefault(); saveDay(); };
  $('bDayCancel').onclick = function () { $('dlgDay').close(); };
  $('bClear').onclick = function () {
    if (run && run.date === editKey) { $('dayErr').textContent = 'Die Erfassung läuft noch – bitte erst stoppen.'; return; }
    delete entries[editKey]; saveEntries(); $('dlgDay').close(); render();
  };
  $('bSettings').onclick = openSettings;
  // Kleiner Button oben rechts: Auto -> Mobil -> Desktop -> Auto
  $('bView').onclick = function () {
    var order = ['auto', 'mobile', 'desktop'];
    var wasWide = wantsWide(settings.layout);
    settings.layout = order[(order.indexOf(settings.layout) + 1) % order.length];
    save('az_settings', settings);
    // Handy-Browser übernehmen eine geänderte Viewport-Breite nicht zuverlässig im laufenden Betrieb: dann neu laden
    if (wasWide !== wantsWide(settings.layout)) { location.reload(); return; }
    applyView();
    toast('Ansicht: ' + VIEW_LABEL[settings.layout]);
  };
  // In der Einstellung "Auto" folgt die Ansicht der Fensterbreite (auch beim Drehen oder Größe ändern)
  window.addEventListener('resize', function () {
    if (settings.layout !== 'mobile' && settings.layout !== 'desktop') applyView();
  });
  $('fSet').onsubmit = function (ev) { ev.preventDefault(); saveSettings(); };
  $('bSetCancel').onclick = function () { $('dlgSet').close(); };
  $('bFill').onclick = fillWeekdays;
  $('bExport').onclick = exportMonth;
  $('bNameReset').onclick = function () { var el = $('nName'); el.value = defaultName(); el.focus(); };
  $('bNameCancel').onclick = function () { $('dlgName').close(); };
  $('fName').onsubmit = function (ev) {
    ev.preventDefault();
    var name = AZ.sanitizeFileName($('nName').value);
    if (!name) { $('nameErr').textContent = 'Bitte einen Dateinamen eingeben.'; return; }
    $('dlgName').close();
    doExport(name);
  };

  setInterval(tick, 1000);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) { render(); } });

  render(); scrollToToday();

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
  }
})();
