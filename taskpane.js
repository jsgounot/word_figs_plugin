/*
 * SuppRef task pane — Office.js glue around SuppRefCore (core.js).
 * ES2017: async/await is fine; no optional chaining, nullish coalescing,
 * object spread or optional catch binding.
 */
(function () {
  'use strict';

  var C = window.SuppRefCore;
  var DEFAULTS_KEY = 'suppref-defaults';
  var SYNC_BATCH = 100;
  var VISIBLE_TOGGLES = 12;

  var state = {
    registry: null,
    selectedFig: null,      // figure id targeted by the Insert form
    panels: [],             // 1-based indices chosen in the Insert form
    editing: null,          // { id, kind } of the control being edited
    warnings: [],
    summary: '',
    maxUsedPanel: {},       // series id → highest panel index used by its citations
    addSeq: null,           // series chosen in the Add figure form
    draft: null,            // Settings view working copy: { series, config, current }
    busy: false,
    api14: false,
    showAllToggles: false,
    editingFigRow: null,    // figure id whose row is in rename mode
    selTimer: null
  };

  var $ = function (id) { return document.getElementById(id); };

  // ------------------------------------------------------------- bootstrap

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) {
      showMessage('SuppRef only runs inside Microsoft Word.', 'error');
      return;
    }
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      showMessage('This version of Word is too old for SuppRef (WordApi 1.3 required).', 'error');
      return;
    }
    state.api14 = Office.context.requirements.isSetSupported('WordApi', '1.4');
    applyTheme();
    loadRegistry();
    bindEvents();
    renderAll();
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    } catch (e) { /* event not available: the Load selection button still works */ }
    checkSelection(false);
  });

  function applyTheme() {
    try {
      var t = Office.context.officeTheme;
      if (!t || !t.bodyBackgroundColor) return;
      var hex = t.bodyBackgroundColor.replace('#', '');
      var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
      var lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      document.documentElement.setAttribute('data-theme', lum < 0.5 ? 'dark' : 'light');
    } catch (e) { /* fall back to prefers-color-scheme */ }
  }

  // -------------------------------------------------------------- registry

  function loadDefaults() {
    try {
      if (Office.context.roamingSettings) {
        var v = Office.context.roamingSettings.get(DEFAULTS_KEY);
        if (v) return v;
      }
    } catch (e) { /* not available in Word */ }
    try {
      var s = window.localStorage.getItem(DEFAULTS_KEY);
      return s ? JSON.parse(s) : null;
    } catch (e) { return null; }
  }

  function saveDefaults(cfg) {
    try {
      if (Office.context.roamingSettings) {
        Office.context.roamingSettings.set(DEFAULTS_KEY, cfg);
        Office.context.roamingSettings.saveAsync();
        return true;
      }
    } catch (e) { /* fall through */ }
    try {
      window.localStorage.setItem(DEFAULTS_KEY, JSON.stringify(cfg));
      return true;
    } catch (e) { return false; }
  }

  function loadRegistry() {
    var raw = Office.context.document.settings.get(C.SETTINGS_KEY);
    state.registry = C.normalizeRegistry(raw, raw ? null : loadDefaults());
  }

  function saveRegistry() {
    var settings = Office.context.document.settings;
    settings.set(C.SETTINGS_KEY, state.registry);
    return new Promise(function (resolve) {
      settings.saveAsync(function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          showMessage('Could not save the figure registry: ' + (res.error && res.error.message), 'error');
        }
        resolve();
      });
    });
  }

  function figures() { return state.registry.figures; }
  function seriesList() { return state.registry.series; }
  /** Document-wide settings (appearance, auto-refresh, track changes, warnings). */
  function cfg() { return state.registry.config; }
  /** Formatting of the series a figure belongs to. */
  function fmtFor(figId) { return C.formatOf(state.registry, figId); }
  function lastNumber(id) {
    var n = state.registry.lastRefresh.numbers[id];
    return typeof n === 'number' ? n : null;
  }

  // ------------------------------------------------------------- utilities

  function isMine(tag) { return typeof tag === 'string' && tag.indexOf(C.TAG_PREFIX) === 0; }

  function errText(e) {
    if (!e) return 'Unknown error';
    var msg = e.message || String(e);
    if (e.code) msg += ' (' + e.code + ')';
    return msg;
  }

  function setBusy(on) {
    state.busy = on;
    $('spinner').hidden = !on;
    Array.prototype.forEach.call(document.querySelectorAll('button, input, select'), function (el) {
      if (on) { el.dataset.wasDisabled = el.disabled ? '1' : ''; el.disabled = true; }
      else { el.disabled = el.dataset.wasDisabled === '1'; }
    });
  }

  /** Wrap a user action: one at a time, spinner, inline errors. */
  async function action(fn) {
    if (state.busy) return;
    clearMessage();
    setBusy(true);
    try { await fn(); }
    catch (e) { showMessage(errText(e), 'error'); }
    finally { setBusy(false); renderAll(); }
  }

  function showMessage(text, kind) {
    var m = $('message');
    m.textContent = text;
    m.className = 'message' + (kind ? ' ' + kind : '');
    m.hidden = false;
  }
  function clearMessage() { $('message').hidden = true; }

  /** Inline confirmation (alert/confirm are unsupported in add-ins). Resolves true/false. */
  function confirmInline(text, okLabel) {
    return new Promise(function (resolve) {
      var box = $('confirm');
      $('confirmText').textContent = text;
      $('confirmOk').textContent = okLabel;
      box.hidden = false;
      // The confirm box must stay usable while everything else is disabled.
      $('confirmOk').disabled = false;
      $('confirmCancel').disabled = false;
      var done = function (v) {
        box.hidden = true;
        $('confirmOk').onclick = null;
        $('confirmCancel').onclick = null;
        resolve(v);
      };
      $('confirmOk').onclick = function () { done(true); };
      $('confirmCancel').onclick = function () { done(false); };
      box.scrollIntoView({ block: 'nearest' });
    });
  }

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (k === 'text') e.textContent = attrs[k];
      else if (k === 'class') e.className = attrs[k];
      else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { if (c) e.appendChild(c); });
    return e;
  }

  // ------------------------------------------------------- document access

  /**
   * Load every SuppRef control in the body, in document order, and parse tags.
   * Returns { items: [{kind, figId, panels, flags, ref, cc}], malformed: [{cc, error}], all: [cc] }.
   */
  async function scanControls(ctx) {
    var ccs = ctx.document.body.contentControls;
    ccs.load('items/id,items/tag,items/text,items/title,items/appearance,items/cannotEdit,items/cannotDelete,items/font/bold');
    await ctx.sync();
    var mine = ccs.items.filter(function (cc) { return isMine(cc.tag); });
    mine = await ensureDocumentOrder(ctx, mine);
    var items = [], malformed = [];
    mine.forEach(function (cc) {
      var p = C.parseTag(cc.tag);
      if (p.error) malformed.push({ cc: cc, error: p.error });
      else items.push({ kind: p.kind, figId: p.figId, panels: p.panels, flags: p.flags, ref: cc.id, cc: cc });
    });
    return { items: items, malformed: malformed, all: mine };
  }

  var BEFORE = { Before: 1, AdjacentBefore: 1, OverlapsBefore: 1, Contains: 1, ContainsStart: 1 };

  /**
   * §6.3: body.contentControls is expected to be in document order. Check consecutive
   * pairs with compareLocationWith (one sync); if any pair is out of order, fall back
   * to ranking every control by how many others precede it.
   */
  async function ensureDocumentOrder(ctx, list) {
    if (list.length < 2) return list;
    var ranges = list.map(function (cc) { return cc.getRange('Whole'); });
    var cmp = [];
    for (var i = 1; i < ranges.length; i++) cmp.push(ranges[i - 1].compareLocationWith(ranges[i]));
    await ctx.sync();
    if (cmp.every(function (r) { return BEFORE[r.value]; })) return list;

    console.warn('SuppRef: contentControls not in document order; sorting with compareLocationWith');
    var before = list.map(function () { return 0; });
    var pending = [];
    for (var a = 0; a < list.length; a++) {
      for (var b = 0; b < list.length; b++) {
        if (a === b) continue;
        pending.push({ a: a, r: ranges[b].compareLocationWith(ranges[a]) });
        if (pending.length >= 2000) { await flush(); }
      }
    }
    await flush();
    async function flush() {
      await ctx.sync();
      pending.forEach(function (p) { if (BEFORE[p.r.value]) before[p.a]++; });
      pending = [];
    }
    return list.map(function (cc, idx) { return { cc: cc, k: before[idx] }; })
      .sort(function (x, y) { return x.k - y.k; })
      .map(function (x) { return x.cc; });
  }

  function titleFor(kind, figId) {
    var f = C.figureById(state.registry, figId);
    var label = f ? f.label : '(unregistered ' + figId + ')';
    return kind === 'l' ? C.legendTitle(label) : C.citationTitle(label);
  }

  function styleControl(cc, kind, figId) {
    cc.title = titleFor(kind, figId);
    cc.appearance = cfg().appearance;
    cc.font.bold = fmtFor(figId).bold;
    cc.cannotDelete = false;
    cc.cannotEdit = true;
  }

  // --------------------------------------------------------------- refresh

  /** §9.7. Not wrapped in action() so other commands can chain it. */
  async function doRefresh() {
    var reg = state.registry, config = cfg();
    var summary = null;
    await Word.run(async function (ctx) {
      var scan = await scanControls(ctx);
      var res = C.computeNumbering(reg, scan.items);
      var warnings = res.warnings.slice();
      scan.malformed.forEach(function (m) {
        warnings.push({ type: 'malformed', ref: m.cc.id, message: 'Malformed SuppRef control (' + m.error + '), left untouched' });
      });

      var changes = [], textChanged = 0, boldChanged = 0, citations = 0;
      state.maxUsedPanel = maxPanelsBySeries(scan.items);
      scan.items.forEach(function (it) {
        if (it.kind === 'c') citations++;
        var target = C.targetText(it, reg, res.numbers);
        var title = titleFor(it.kind, it.figId);
        var t = it.cc.text !== target;
        var bold = fmtFor(it.figId).bold;
        // font.bold is null when the control mixes bold and regular text.
        var b = it.cc.font.bold !== bold;
        if (t) textChanged++;
        if (b) boldChanged++;
        if (t || b || it.cc.title !== title || it.cc.appearance !== config.appearance || it.cc.cannotEdit !== true) {
          changes.push({ cc: it.cc, target: target, title: title, text: t, bold: b ? bold : null });
        }
      });

      var restoreMode = null;
      // Bold changes are formatting revisions, so they also need Track Changes suspended.
      if ((textChanged || boldChanged) && config.suspendTrackChanges) {
        if (state.api14) {
          ctx.document.load('changeTrackingMode');
          await ctx.sync();
          if (ctx.document.changeTrackingMode !== 'Off') {
            restoreMode = ctx.document.changeTrackingMode;
            ctx.document.changeTrackingMode = 'Off';
            await ctx.sync();
          }
        } else {
          warnings.push({ type: 'trackChanges', message: 'Track Changes may record renumbering (WordApi 1.4 not available in this Word).' });
        }
      }

      try {
        for (var i = 0; i < changes.length; i++) {
          var c = changes[i];
          if (c.text || c.bold !== null) c.cc.cannotEdit = false;
          if (c.text) c.cc.insertText(c.target, 'Replace');
          if (c.bold !== null) c.cc.font.bold = c.bold;
          c.cc.cannotEdit = true;
          if (c.cc.title !== c.title) c.cc.title = c.title;
          if (c.cc.appearance !== config.appearance) c.cc.appearance = config.appearance;
          if ((i + 1) % SYNC_BATCH === 0) await ctx.sync();
        }
        await ctx.sync();
      } finally {
        if (restoreMode) {
          ctx.document.changeTrackingMode = restoreMode;
          await ctx.sync();
        }
      }

      reg.lastRefresh = { at: new Date().toISOString(), numbers: res.numbers, counts: res.counts };
      state.warnings = warnings;
      summary = 'Refreshed: ' + citations + ' citation' + (citations === 1 ? '' : 's') + ', ' +
        reg.figures.length + ' figure' + (reg.figures.length === 1 ? '' : 's') + ', ' + textChanged + ' changed';
    });
    await saveRegistry();
    state.summary = summary;
  }

  /** Highest panel index cited per series, for the panelStart check in Settings. */
  function maxPanelsBySeries(items) {
    var out = {};
    items.forEach(function (it) {
      var f = C.figureById(state.registry, it.figId);
      if (it.kind !== 'c' || !f || !it.panels.length) return;
      out[f.seq] = Math.max(out[f.seq] || 0, it.panels[it.panels.length - 1]);
    });
    return out;
  }

  async function maybeAutoRefresh() {
    if (cfg().autoRefreshOnInsert) await doRefresh();
  }

  // ---------------------------------------------------------------- insert

  function insertKindText(kind, figId, panels, flags) {
    var n = lastNumber(figId), fmt = fmtFor(figId);
    return kind === 'l' ? C.renderLegend(n, fmt) : C.renderCitation(n, panels, flags, fmt);
  }

  function readInsertForm() {
    var figId = $('insFigure').value;
    if (!figId || !C.figureById(state.registry, figId)) throw new Error('Choose a figure first (add one under Figures).');
    var parsed = C.parsePanelText($('panelText').value, fmtFor(figId));
    if (parsed.error) throw new Error(parsed.error);
    var max = C.maxPanels(fmtFor(figId));
    if (parsed.panels.length && parsed.panels[parsed.panels.length - 1] > max) {
      throw new Error('Only ' + max + ' panels are available with the current panel settings.');
    }
    return { figId: figId, panels: parsed.panels, flags: $('numberOnly').checked ? 'n' : '' };
  }

  /** §9.4 / §9.6. kind 'c' or 'l'. */
  async function doInsert(kind) {
    var form = readInsertForm();
    var enteredEdit = false;
    await Word.run(async function (ctx) {
      var sel = ctx.document.getSelection();
      var parent = sel.parentContentControlOrNullObject;
      parent.load('id,tag');
      await ctx.sync();
      if (!parent.isNullObject && isMine(parent.tag)) {
        enteredEdit = true;
        return;
      }
      var tag = kind === 'l' ? C.buildLegendTag(form.figId) : C.buildCitationTag(form.figId, form.panels, form.flags);
      var range = sel.insertText(insertKindText(kind, form.figId, form.panels, form.flags), 'Replace');
      var cc = range.insertContentControl();
      cc.tag = tag;
      styleControl(cc, kind, form.figId);
      cc.getRange('After').select();
      await ctx.sync();
    });
    if (enteredEdit) {
      await checkSelection(true);
      showMessage('The cursor is inside a SuppRef control, so it was loaded for editing. Change it and press Update, or move the cursor out first.', 'warn');
      return;
    }
    await maybeAutoRefresh();
  }

  /** §9.5: rewrite the tag and text of the control being edited. */
  async function doUpdate() {
    var form = readInsertForm();
    var editing = state.editing;
    await Word.run(async function (ctx) {
      var cc = ctx.document.body.contentControls.getById(editing.id);
      cc.load('id');
      await ctx.sync();
      cc.tag = editing.kind === 'l' ? C.buildLegendTag(form.figId) : C.buildCitationTag(form.figId, form.panels, form.flags);
      cc.cannotEdit = false;
      cc.insertText(insertKindText(editing.kind, form.figId, form.panels, form.flags), 'Replace');
      styleControl(cc, editing.kind, form.figId);
      await ctx.sync();
    });
    exitEdit();
    await maybeAutoRefresh();
  }

  // ------------------------------------------------------ selection / edit

  function onSelectionChanged() {
    clearTimeout(state.selTimer);
    state.selTimer = setTimeout(function () { if (!state.busy) checkSelection(false); }, 250);
  }

  /** If the selection is inside a SuppRef control, load it into the form (edit mode). */
  async function checkSelection(explicit) {
    try {
      var found = null;
      await Word.run(async function (ctx) {
        var parent = ctx.document.getSelection().parentContentControlOrNullObject;
        parent.load('id,tag');
        await ctx.sync();
        if (!parent.isNullObject && isMine(parent.tag)) found = { id: parent.id, tag: parent.tag };
      });
      if (!found) {
        if (state.editing) exitEdit();
        else if (explicit) showMessage('The cursor is not inside a SuppRef citation or legend label.', 'warn');
        renderInsert();
        return;
      }
      var p = C.parseTag(found.tag);
      if (p.error) { showMessage('This SuppRef control has a malformed tag: ' + p.error, 'warn'); return; }
      if (!C.figureById(state.registry, p.figId)) {
        showMessage('This control points to an unregistered figure ("' + p.figId + '"). Refresh, then use Register orphans.', 'warn');
        return;
      }
      state.editing = { id: found.id, kind: p.kind };
      state.selectedFig = p.figId;
      state.panels = p.panels;
      $('numberOnly').checked = p.flags.indexOf('n') >= 0;
      $('panelText').value = C.formatPanelText(p.panels, fmtFor(p.figId));
      if (p.panels.length && p.panels[p.panels.length - 1] > VISIBLE_TOGGLES) state.showAllToggles = true;
      renderAll();
    } catch (e) {
      if (explicit) showMessage(errText(e), 'error');
    }
  }

  function exitEdit() {
    state.editing = null;
    renderInsert();
  }

  async function goTo(ref) {
    await Word.run(async function (ctx) {
      ctx.document.body.contentControls.getById(ref).select();
      await ctx.sync();
    });
  }

  // --------------------------------------------------------------- figures

  async function addFigure(label, description, seq) {
    var err = C.validateLabel(label, figures()) || C.validateDescription(description);
    if (err) throw new Error(err);
    if (!C.seriesById(state.registry, seq)) throw new Error('Choose a series for the figure.');
    var f = {
      id: C.genId(figures().map(function (x) { return x.id; })),
      seq: seq,
      label: label,
      description: description || '',
      manualOrder: C.nextManualOrder(figures())
    };
    figures().push(f);
    state.selectedFig = f.id;
    await saveRegistry();
    return f;
  }

  async function renameFigure(id, label, description, seq) {
    var err = C.validateLabel(label, figures(), id) || C.validateDescription(description);
    if (err) throw new Error(err);
    var f = C.figureById(state.registry, id);
    var moved = seq && seq !== f.seq && C.seriesById(state.registry, seq);
    f.label = label;
    f.description = description || '';
    if (moved) {
      f.seq = seq;
      f.manualOrder = C.nextManualOrder(figures());
    }
    state.editingFigRow = null;
    await saveRegistry();
    // Moving a figure to another series changes its number and its citations' prefix.
    if (moved) await doRefresh();
  }

  /** §9.3 */
  async function deleteFigure(id) {
    var f = C.figureById(state.registry, id);
    var count = 0;
    await Word.run(async function (ctx) {
      var scan = await scanControls(ctx);
      count = scan.items.filter(function (it) { return it.figId === id; }).length;
    });
    if (count) {
      var ok = await confirmInline('"' + f.label + '" has ' + count + ' citation/legend control' + (count === 1 ? '' : 's') +
        ' in the document. Deleting it converts them to plain text.', 'Delete the figure and unlink its citations');
      if (!ok) return;
      await Word.run(async function (ctx) {
        var scan = await scanControls(ctx);
        scan.items.forEach(function (it) {
          if (it.figId !== id) return;
          it.cc.cannotEdit = false;
          it.cc.cannotDelete = false;
          it.cc.delete(true);
        });
        await ctx.sync();
      });
    }
    state.registry.figures = figures().filter(function (x) { return x.id !== id; });
    if (state.selectedFig === id) state.selectedFig = null;
    await saveRegistry();
    if (count) await maybeAutoRefresh();
  }

  /** Manual ordering is per series: swap with the neighbour in the same series. */
  async function moveFigure(id, delta) {
    var seq = C.figureById(state.registry, id).seq;
    var list = C.byManualOrder(figures().filter(function (f) { return f.seq === seq; }));
    var i = list.findIndex(function (f) { return f.id === id; });
    var j = i + delta;
    if (j < 0 || j >= list.length) return;
    var slots = list.map(function (f) { return f.manualOrder; });
    list.splice(j, 0, list.splice(i, 1)[0]);
    list.forEach(function (f, k) { f.manualOrder = slots[k]; });
    await saveRegistry();
    await maybeAutoRefresh();
  }

  async function registerOrphans() {
    var ids = {};
    state.warnings.forEach(function (w) { if (w.type === 'orphan') ids[w.figId] = true; });
    Object.keys(ids).forEach(function (id) {
      if (C.figureById(state.registry, id)) return;
      var base = ('orphan-' + id).slice(0, 30), label = base, k = 2;
      while (C.validateLabel(label, figures())) label = (base.slice(0, 27) + '-' + k++);
      // Orphans join the supplementary series; the user can move them via ✎.
      figures().push({ id: id, seq: C.defaultSeries(state.registry).id, label: label, description: '',
        manualOrder: C.nextManualOrder(figures()) });
    });
    await saveRegistry();
    await doRefresh();
  }

  // ---------------------------------------------------------------- import

  var INSIDE = { Inside: 1, InsideStart: 1, InsideEnd: 1 };

  /** §9.9, once per series that has a placeholder word (Figure_droso_A, MainFig_overview…). */
  async function doImport() {
    var patterns = C.placeholderPatterns(seriesList());
    if (!patterns.length) throw new Error('No series has a placeholder word. Set one in Settings.');
    var converted = 0, created = 0, skipped = [];
    await Word.run(async function (ctx) {
      var body = ctx.document.body;
      var results = patterns.map(function (p) {
        var res = body.search(p.pattern, { matchWildcards: true });
        res.load('items/text');
        return res;
      });
      await ctx.sync();

      var all = [];
      results.forEach(function (res) { res.items.forEach(function (range) { all.push(range); }); });
      if (!all.length) return;

      // Compare every pair once. This drops matches nested in a longer one ("Figure_droso"
      // inside "Figure_droso_A", or duplicates) and yields document order, so replacements
      // can run back to front and new figures are registered in reading order.
      var n = all.length, before = [], drop = [], pending = [];
      for (var z = 0; z < n; z++) { before.push(0); drop.push(false); }
      var flush = async function () {
        await ctx.sync();
        pending.forEach(function (p) {
          var v = p.r.value;
          if (BEFORE[v]) before[p.b]++;
          if (INSIDE[v] || (v === 'Equal' && p.a > p.b)) drop[p.a] = true;
        });
        pending = [];
      };
      for (var a = 0; a < n; a++) {
        for (var b = 0; b < n; b++) {
          if (a === b) continue;
          pending.push({ a: a, b: b, r: all[a].compareLocationWith(all[b]) });
          if (pending.length >= 2000) await flush();
        }
      }
      if (pending.length) await flush();

      var live = [];
      for (var i = 0; i < n; i++) if (!drop[i]) live.push({ range: all[i], pos: before[i] });
      live.sort(function (x, y) { return x.pos - y.pos; });

      var byLabel = {};
      figures().forEach(function (f) { byLabel[f.label.toLowerCase()] = f; });
      var work = [];
      live.forEach(function (m) {
        var p = C.parsePlaceholder(m.range.text, seriesList());
        if (!p) return;
        if (C.validateLabel(p.label, []) !== null) { skipped.push(m.range.text + ' (label too long)'); return; }
        var f = byLabel[p.label.toLowerCase()];
        if (f && f.seq !== p.seq) {
          skipped.push(m.range.text + ' (label "' + f.label + '" already belongs to ' + seriesName(f.seq) + ')');
          return;
        }
        if (!f) {
          f = {
            id: C.genId(figures().map(function (x) { return x.id; })),
            seq: p.seq, label: p.label, description: '',
            manualOrder: C.nextManualOrder(figures())
          };
          figures().push(f);
          byLabel[p.label.toLowerCase()] = f;
          created++;
        }
        work.push({ range: m.range, fig: f, panels: p.panel ? [p.panel] : [] });
      });

      for (var k = work.length - 1, done = 1; k >= 0; k--, done++) {
        var w = work[k];
        var r = w.range.insertText(insertKindText('c', w.fig.id, w.panels, ''), 'Replace');
        var cc = r.insertContentControl();
        cc.tag = C.buildCitationTag(w.fig.id, w.panels, '');
        styleControl(cc, 'c', w.fig.id);
        converted++;
        if (done % SYNC_BATCH === 0) await ctx.sync();
      }
      await ctx.sync();
    });
    await saveRegistry();
    await doRefresh();
    var msg = 'Imported ' + converted + ' placeholder' + (converted === 1 ? '' : 's') + ', created ' + created + ' figure' + (created === 1 ? '' : 's') + '.';
    if (skipped.length) msg += '\nSkipped: ' + skipped.join('; ');
    showMessage(msg, converted ? '' : 'warn');
  }

  // ---------------------------------------------------------------- unlink

  /** §9.8 */
  async function doUnlinkAll() {
    var ok = await confirmInline('Unlink all converts every SuppRef citation and legend label into plain text. ' +
      'This cannot be undone by SuppRef — save a copy of the document first.', 'Unlink all');
    if (!ok) return;
    await doRefresh();
    var n = 0;
    await Word.run(async function (ctx) {
      var scan = await scanControls(ctx);
      for (var i = 0; i < scan.all.length; i++) {
        var cc = scan.all[i];
        cc.cannotEdit = false;
        cc.cannotDelete = false;
        cc.delete(true);
        n++;
        if ((i + 1) % SYNC_BATCH === 0) await ctx.sync();
      }
      await ctx.sync();
    });
    exitEdit();
    state.warnings = [];
    state.summary = 'Unlinked ' + n + ' control' + (n === 1 ? '' : 's') + '. The text is now plain text.';
  }

  // -------------------------------------------------------------- settings
  // The Settings view edits a draft copy of the series list and global settings;
  // nothing reaches the document until Save.

  var CFG_INT = { numberStart: 1, panelRangeMinRun: 1 };

  function draftSeries() {
    var d = state.draft;
    return d.series.filter(function (s) { return s.id === d.current; })[0];
  }

  /** Inputs of one scope: "series" (name, placeholder), "format" (default) or "global". */
  function formInputs(scope) {
    return Array.prototype.filter.call(document.querySelectorAll('#cfgForm [data-key]'), function (input) {
      return (input.dataset.scope || 'format') === scope;
    });
  }

  function fillSettingsForm() {
    var d = state.draft, s = draftSeries();
    var sel = $('cfgSeries');
    sel.textContent = '';
    d.series.forEach(function (x) { sel.appendChild(el('option', { value: x.id, text: x.name.trim() || '(unnamed)' })); });
    sel.value = d.current;
    formInputs('series').forEach(function (input) { input.value = s[input.dataset.key] || ''; });
    formInputs('format').forEach(function (input) {
      if (input.type === 'checkbox') input.checked = !!s.format[input.dataset.key];
      else input.value = s.format[input.dataset.key];
    });
    formInputs('global').forEach(function (input) {
      if (input.type === 'checkbox') input.checked = !!d.config[input.dataset.key];
      else input.value = d.config[input.dataset.key];
    });
  }

  /** Copy the form into the draft: the series being edited, plus global settings. */
  function readSettingsForm() {
    var d = state.draft, s = draftSeries();
    formInputs('series').forEach(function (input) {
      s[input.dataset.key] = input.dataset.key === 'placeholder' ? input.value.trim() : input.value;
    });
    formInputs('format').forEach(function (input) {
      var k = input.dataset.key;
      if (input.type === 'checkbox') s.format[k] = input.checked;
      else if (CFG_INT[k]) s.format[k] = /^\s*-?\d+\s*$/.test(input.value) ? parseInt(input.value, 10) : NaN;
      else s.format[k] = input.value;
    });
    formInputs('global').forEach(function (input) {
      d.config[input.dataset.key] = input.type === 'checkbox' ? input.checked : input.value;
    });
  }

  /** Every problem in the draft, across all series: [{ series, key, message }]. */
  function draftErrors() {
    var d = state.draft, out = [];
    d.series.forEach(function (s) {
      C.validateSeriesMeta(s, d.series).concat(C.validateFormat(s.format, state.maxUsedPanel[s.id])).forEach(function (e) {
        out.push({ series: s, key: e.key, message: e.message });
      });
    });
    C.validateGlobal(d.config).forEach(function (e) { out.push({ series: null, key: e.key, message: e.message }); });
    return out;
  }

  function updateSettingsPreview() {
    if (!state.draft) return;
    readSettingsForm();
    var s = draftSeries(), errs = draftErrors();
    var box = $('cfgErrors');
    box.hidden = !errs.length;
    box.textContent = errs.map(function (e) {
      return (e.series ? (e.series.name.trim() || '(unnamed)') + ' › ' : '') + e.key + ': ' + e.message;
    }).join('\n');

    var opt = $('cfgSeries').options[$('cfgSeries').selectedIndex];
    if (opt) opt.textContent = s.name.trim() || '(unnamed)';

    var ownFormatErrs = C.validateFormat(s.format, state.maxUsedPanel[s.id]);
    $('cfgPreview').style.fontWeight = s.format.bold === true ? '700' : '400';
    $('cfgPreview').textContent = ownFormatErrs.length ? '—' :
      C.renderCitation(3, [1, 2, 3], '', s.format) + '   ·   ' + C.renderLegend(3, s.format);

    var avail = $('panelsAvailable'), used = state.maxUsedPanel[s.id];
    if (C.isLetterStyle(s.format.panelStyle) && /^[A-Za-z]$/.test(String(s.format.panelStart))) {
      var m = C.maxPanels(s.format);
      avail.textContent = m + ' panels available (' + C.panelLabel(1, s.format) + '–' + C.panelLabel(m, s.format) + ')' +
        (used ? '; this series uses up to ' + used + '.' : '.');
    } else {
      avail.textContent = '';
    }
    $('cfgSave').disabled = !!errs.length;
    $('cfgDefaults').disabled = !!errs.length;
    $('cfgDelSeries').disabled = state.draft.series.length <= 1;
  }

  function openSettings() {
    clearMessage();
    var reg = state.registry;
    var sel = state.selectedFig && C.figureById(reg, state.selectedFig);
    state.draft = {
      series: JSON.parse(JSON.stringify(reg.series)),
      config: Object.assign({}, reg.config),
      current: sel ? sel.seq : reg.series[0].id
    };
    fillSettingsForm();
    updateSettingsPreview();
    $('mainView').hidden = true;
    $('settingsView').hidden = false;
    $('btnSettings').classList.add('active');
    $('btnSettings').textContent = 'Close settings';
  }

  function switchDraftSeries(id) {
    state.draft.current = id;
    fillSettingsForm();
    updateSettingsPreview();
  }

  function addDraftSeries() {
    readSettingsForm();
    var d = state.draft, base = draftSeries();
    var name = 'New series', k = 2;
    var taken = function (nm) { return d.series.some(function (s) { return s.name.trim().toLowerCase() === nm.toLowerCase(); }); };
    while (taken(name)) name = 'New series ' + k++;
    var s = C.makeSeries(C.genId(d.series.map(function (x) { return x.id; })), name, '', base.format);
    d.series.push(s);
    switchDraftSeries(s.id);
    var nameInput = document.querySelector('#cfgForm [data-key="name"]');
    nameInput.focus();
    nameInput.select();
  }

  function deleteDraftSeries() {
    readSettingsForm();
    var d = state.draft, s = draftSeries();
    if (d.series.length <= 1) return;
    var n = figures().filter(function (f) { return f.seq === s.id; }).length;
    if (n) {
      showMessage('"' + s.name.trim() + '" still has ' + n + ' figure' + (n === 1 ? '' : 's') +
        '. Move them to another series (✎ in the figure list) or delete them first.', 'warn');
      return;
    }
    clearMessage();
    d.series = d.series.filter(function (x) { return x !== s; });
    switchDraftSeries(d.series[0].id);
  }

  function closeSettings() {
    state.draft = null;
    $('settingsView').hidden = true;
    $('mainView').hidden = false;
    $('btnSettings').classList.remove('active');
    $('btnSettings').textContent = 'Settings';
    var fid = $('insFigure').value;
    $('panelText').value = C.formatPanelText(state.panels, fmtFor(fid));
    $('panelError').hidden = true;
    renderAll();
  }

  async function saveSettings() {
    // Rescan so the panelStart check uses the document's real maximum panel per series.
    await Word.run(async function (ctx) {
      var scan = await scanControls(ctx);
      state.maxUsedPanel = maxPanelsBySeries(scan.items);
    });
    readSettingsForm();
    if (draftErrors().length) { updateSettingsPreview(); throw new Error('Fix the settings listed in red first.'); }
    var d = state.draft;
    state.registry.series = d.series.map(function (s) { return C.makeSeries(s.id, s.name.trim(), s.placeholder, s.format); });
    state.registry.config = C.mergeGlobal(d.config);
    await saveRegistry();
    closeSettings();
    await doRefresh();
  }

  // ------------------------------------------------------------- rendering

  function renderAll() {
    renderFigures();
    renderInsert();
    renderStatus();
  }

  function seriesName(seq) {
    var s = C.seriesById(state.registry, seq);
    return s ? s.name : seq;
  }

  /** Figures of one series in display order: by current number, or by manual order. */
  function figuresOf(s) {
    var figs = figures().filter(function (f) { return f.seq === s.id; });
    if (s.format.orderMode === 'manual') return C.byManualOrder(figs);
    var nums = state.registry.lastRefresh.numbers;
    return figs.sort(function (a, b) {
      var x = typeof nums[a.id] === 'number' ? nums[a.id] : Infinity;
      var y = typeof nums[b.id] === 'number' ? nums[b.id] : Infinity;
      return x - y;
    });
  }

  function seriesOptions(select, selected) {
    select.textContent = '';
    seriesList().forEach(function (s) { select.appendChild(el('option', { value: s.id, text: s.name })); });
    if (selected) select.value = selected;
  }

  function renderFigures() {
    var reg = state.registry;
    var list = $('figList');
    list.textContent = '';
    $('figEmpty').hidden = figures().length > 0;
    var refreshed = !!reg.lastRefresh.at;
    var counts = reg.lastRefresh.counts || {};

    seriesList().forEach(function (s) {
      var figs = figuresOf(s);
      if (!figs.length) return;
      var manual = s.format.orderMode === 'manual';
      list.appendChild(el('li', { class: 'group', text: s.name }));

      figs.forEach(function (f, idx) {
        var n = lastNumber(f.id);
        var count = counts[f.id] || 0;
        var li = el('li', { class: f.id === state.selectedFig ? 'selected' : '', title: f.description || f.label });
        li.addEventListener('click', function () { state.selectedFig = f.id; renderAll(); });

        if (state.editingFigRow === f.id) {
          var stop = function (e) { e.stopPropagation(); };
          var lab = el('input', { value: f.label, maxlength: '30' });
          var desc = el('input', { value: f.description, maxlength: '200', placeholder: 'Description' });
          var ser = el('select', { title: 'Series' });
          seriesOptions(ser, f.seq);
          [lab, desc, ser].forEach(function (x) { x.addEventListener('click', stop); });
          li.appendChild(el('div', { class: 'editrow' }, [ser, lab, desc, el('div', { class: 'row' }, [
            el('button', { class: 'primary', text: 'Save', onclick: function (e) {
              e.stopPropagation();
              action(function () { return renameFigure(f.id, lab.value.trim(), desc.value.trim(), ser.value); });
            } }),
            el('button', { text: 'Cancel', onclick: function (e) { e.stopPropagation(); state.editingFigRow = null; renderFigures(); } })
          ])]));
          list.appendChild(li);
          return;
        }

        li.appendChild(el('span', { class: 'num', text: n === null ? '–' : C.formatNumber(n, s.format.numberStyle) }));
        li.appendChild(el('span', { class: 'name' }, [
          el('span', { class: 'label', text: f.label }),
          el('span', { class: 'meta', text: '  ' + count + '×' }),
          refreshed && !count ? el('span', { class: 'warnicon', text: ' ⚠', title: 'Uncited' }) : null,
          f.description ? el('span', { class: 'desc', text: f.description }) : null
        ]));
        var act = el('span', { class: 'actions' });
        var btn = function (text, title, fn, disabled) {
          var b = el('button', { class: 'icon', text: text, title: title, onclick: function (e) { e.stopPropagation(); fn(); } });
          if (disabled) b.disabled = true;
          act.appendChild(b);
        };
        if (manual) {
          btn('↑', 'Move up', function () { action(function () { return moveFigure(f.id, -1); }); }, idx === 0);
          btn('↓', 'Move down', function () { action(function () { return moveFigure(f.id, 1); }); }, idx === figs.length - 1);
        }
        btn('✎', 'Rename / edit / change series', function () { state.editingFigRow = f.id; renderFigures(); });
        btn('✕', 'Delete', function () { action(function () { return deleteFigure(f.id); }); });
        li.appendChild(act);
        list.appendChild(li);
      });
    });

    if (!state.addSeq || !C.seriesById(reg, state.addSeq)) state.addSeq = C.defaultSeries(reg).id;
    seriesOptions($('addSeries'), state.addSeq);
  }

  function renderInsert() {
    var reg = state.registry;
    // Figure dropdown, grouped by series
    var sel = $('insFigure');
    sel.textContent = '';
    var first = null;
    seriesList().forEach(function (s) {
      var figs = figuresOf(s);
      if (!figs.length) return;
      var group = el('optgroup', { label: s.name });
      figs.forEach(function (f) {
        if (!first) first = f.id;
        var n = lastNumber(f.id);
        group.appendChild(el('option', { value: f.id, text: (n === null ? '' : C.formatNumber(n, s.format.numberStyle) + ' · ') + f.label }));
      });
      sel.appendChild(group);
    });
    if (!first) sel.appendChild(el('option', { value: '', text: '(no figures)' }));
    if (state.selectedFig && C.figureById(reg, state.selectedFig)) sel.value = state.selectedFig;
    else { state.selectedFig = first; sel.value = first || ''; }

    var c = fmtFor(sel.value);

    // Panel toggles, in the selected figure's series style
    var box = $('panelToggles');
    box.textContent = '';
    var total = Math.min(C.maxPanels(c), 26);
    var shown = state.showAllToggles ? total : Math.min(total, VISIBLE_TOGGLES);
    var chosen = {};
    state.panels.forEach(function (p) { chosen[p] = true; });
    for (var i = 1; i <= shown; i++) {
      (function (idx) {
        box.appendChild(el('button', {
          text: C.panelLabel(idx, c), 'aria-pressed': chosen[idx] ? 'true' : 'false',
          onclick: function () {
            var s = state.panels.slice();
            var at = s.indexOf(idx);
            if (at >= 0) s.splice(at, 1); else s.push(idx);
            state.panels = C.normalizePanels(s);
            $('panelText').value = C.formatPanelText(state.panels, c);
            $('panelError').hidden = true;
            renderInsert();
          }
        }));
      })(i);
    }
    if (total > VISIBLE_TOGGLES) {
      box.appendChild(el('button', {
        text: state.showAllToggles ? 'less' : 'more…',
        onclick: function () { state.showAllToggles = !state.showAllToggles; renderInsert(); }
      }));
    }

    // Edit-mode labels
    var editing = state.editing;
    $('insertHeading').textContent = editing ? (editing.kind === 'l' ? 'Edit legend label' : 'Edit citation') : 'Insert';
    $('btnInsert').textContent = editing ? (editing.kind === 'l' ? 'Update legend label' : 'Update citation') : 'Insert citation';
    $('btnLegend').hidden = !!editing;
    $('btnCancelEdit').hidden = !editing;
    if (!state.busy) {
      $('btnInsert').disabled = !first;
      $('btnLegend').disabled = !first;
    }

    // Preview
    var fid = sel.value;
    var flags = $('numberOnly').checked ? 'n' : '';
    $('insPreview').style.fontWeight = c.bold ? '700' : '400';
    $('insPreview').textContent = !fid ? '' :
      (editing && editing.kind === 'l') ? C.renderLegend(lastNumber(fid), c) :
        C.renderCitation(lastNumber(fid), state.panels, flags, c);
  }

  function renderStatus() {
    var lr = state.registry.lastRefresh;
    $('lastRefresh').textContent = lr.at ? 'Last refresh: ' + new Date(lr.at).toLocaleString() : 'Not refreshed yet.';
    $('summary').textContent = state.summary;
    var w = state.warnings;
    $('warnBox').hidden = !w.length;
    $('warnCount').textContent = w.length;
    var list = $('warnList');
    list.textContent = '';
    w.forEach(function (warn) {
      list.appendChild(el('li', {}, [
        el('span', { text: '⚠ ' + warn.message }),
        warn.ref !== undefined && warn.ref !== null ? el('button', {
          class: 'icon', text: 'Go to',
          onclick: function () { action(function () { return goTo(warn.ref); }); }
        }) : null
      ]));
    });
    $('btnRegisterOrphans').hidden = !w.some(function (x) { return x.type === 'orphan'; });
  }

  // ---------------------------------------------------------------- events

  function bindEvents() {
    $('btnRefresh').addEventListener('click', function () { action(doRefresh); });
    $('btnSettings').addEventListener('click', function () {
      if (state.draft) closeSettings(); else openSettings();
    });
    $('version').textContent = 'v' + C.VERSION;
    $('cfgBackTop').addEventListener('click', closeSettings);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && state.draft && !state.busy) closeSettings();
    });
    $('btnImport').addEventListener('click', function () { action(doImport); });
    $('btnUnlink').addEventListener('click', function () { action(doUnlinkAll); });

    $('addSeries').addEventListener('change', function () { state.addSeq = this.value; });
    $('addForm').addEventListener('submit', function (e) {
      e.preventDefault();
      var label = $('addLabel').value.trim(), desc = $('addDesc').value.trim(), seq = $('addSeries').value;
      action(async function () {
        await addFigure(label, desc, seq);
        $('addLabel').value = '';
        $('addDesc').value = '';
      });
    });

    $('insFigure').addEventListener('change', function () {
      state.selectedFig = this.value;
      // Different series may use a different panel style.
      $('panelText').value = C.formatPanelText(state.panels, fmtFor(this.value));
      $('panelError').hidden = true;
      renderAll();
    });
    $('panelText').addEventListener('input', function () {
      var parsed = C.parsePanelText(this.value, fmtFor($('insFigure').value));
      var errBox = $('panelError');
      if (parsed.error) { errBox.textContent = parsed.error; errBox.hidden = false; return; }
      errBox.hidden = true;
      state.panels = parsed.panels;
      renderInsert();
    });
    $('numberOnly').addEventListener('change', renderInsert);

    $('btnInsert').addEventListener('click', function () {
      action(function () { return state.editing ? doUpdate() : doInsert('c'); });
    });
    $('btnLegend').addEventListener('click', function () { action(function () { return doInsert('l'); }); });
    $('btnLoadSel').addEventListener('click', function () { action(function () { return checkSelection(true); }); });
    $('btnCancelEdit').addEventListener('click', exitEdit);

    $('btnRegisterOrphans').addEventListener('click', function () { action(registerOrphans); });

    $('cfgSeries').addEventListener('change', function () { switchDraftSeries(this.value); });
    $('cfgAddSeries').addEventListener('click', addDraftSeries);
    $('cfgDelSeries').addEventListener('click', deleteDraftSeries);
    $('cfgForm').addEventListener('input', function (e) {
      if (e.target.id !== 'cfgSeries') updateSettingsPreview();
    });
    $('cfgForm').addEventListener('change', function (e) {
      if (e.target.id === 'cfgSeries') return;
      if (e.target.dataset.key === 'panelStyle') {
        var start = document.querySelector('#cfgForm [data-key="panelStart"]');
        start.value = C.convertPanelStart(start.value, e.target.value);
      }
      updateSettingsPreview();
    });
    $('cfgForm').addEventListener('submit', function (e) { e.preventDefault(); });
    $('cfgSave').addEventListener('click', function () { action(saveSettings); });
    $('cfgDefaults').addEventListener('click', function () {
      readSettingsForm();
      if (draftErrors().length) return;
      if (saveDefaults({ series: state.draft.series, config: state.draft.config })) showMessage('Saved as your defaults for new documents.');
      else showMessage('Could not store defaults on this computer.', 'error');
    });
    $('cfgBack').addEventListener('click', closeSettings);
  }
})();
