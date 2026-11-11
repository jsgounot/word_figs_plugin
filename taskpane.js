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
    maxUsedPanel: 0,
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
  function cfg() { return state.registry.config; }
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
    ccs.load('items/id,items/tag,items/text,items/title,items/appearance,items/cannotEdit,items/cannotDelete');
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

      var changes = [], textChanged = 0, citations = 0, maxPanel = 0;
      scan.items.forEach(function (it) {
        if (it.kind === 'c') {
          citations++;
          if (it.panels.length) maxPanel = Math.max(maxPanel, it.panels[it.panels.length - 1]);
        }
        var target = C.targetText(it, reg, res.numbers);
        var title = titleFor(it.kind, it.figId);
        var t = it.cc.text !== target;
        if (t) textChanged++;
        if (t || it.cc.title !== title || it.cc.appearance !== config.appearance || it.cc.cannotEdit !== true) {
          changes.push({ cc: it.cc, target: target, title: title, text: t });
        }
      });
      state.maxUsedPanel = maxPanel;

      var restoreMode = null;
      if (textChanged && config.suspendTrackChanges) {
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
          if (c.text) {
            c.cc.cannotEdit = false;
            c.cc.insertText(c.target, 'Replace');
          }
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

  async function maybeAutoRefresh() {
    if (cfg().autoRefreshOnInsert) await doRefresh();
  }

  // ---------------------------------------------------------------- insert

  function insertKindText(kind, figId, panels, flags) {
    var n = lastNumber(figId);
    return kind === 'l' ? C.renderLegend(n, cfg()) : C.renderCitation(n, panels, flags, cfg());
  }

  function readInsertForm() {
    var figId = $('insFigure').value;
    if (!figId || !C.figureById(state.registry, figId)) throw new Error('Choose a figure first (add one under Figures).');
    var parsed = C.parsePanelText($('panelText').value, cfg());
    if (parsed.error) throw new Error(parsed.error);
    var max = C.maxPanels(cfg());
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
      $('panelText').value = C.formatPanelText(p.panels, cfg());
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

  async function addFigure(label, description) {
    var err = C.validateLabel(label, figures()) || C.validateDescription(description);
    if (err) throw new Error(err);
    var f = {
      id: C.genId(figures().map(function (x) { return x.id; })),
      seq: 'fig',
      label: label,
      description: description || '',
      manualOrder: C.nextManualOrder(figures())
    };
    figures().push(f);
    state.selectedFig = f.id;
    await saveRegistry();
    return f;
  }

  async function renameFigure(id, label, description) {
    var err = C.validateLabel(label, figures(), id) || C.validateDescription(description);
    if (err) throw new Error(err);
    var f = C.figureById(state.registry, id);
    f.label = label;
    f.description = description || '';
    state.editingFigRow = null;
    await saveRegistry();
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

  async function moveFigure(id, delta) {
    var list = C.byManualOrder(figures());
    var i = list.findIndex(function (f) { return f.id === id; });
    var j = i + delta;
    if (j < 0 || j >= list.length) return;
    list.splice(j, 0, list.splice(i, 1)[0]);
    list.forEach(function (f, k) { f.manualOrder = k; });
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
      figures().push({ id: id, seq: 'fig', label: label, description: '', manualOrder: C.nextManualOrder(figures()) });
    });
    await saveRegistry();
    await doRefresh();
  }

  // ---------------------------------------------------------------- import

  var INSIDE = { Equal: 1, Inside: 1, InsideStart: 1, InsideEnd: 1, OverlapsBefore: 1, OverlapsAfter: 1 };

  /** §9.9 */
  async function doImport() {
    var converted = 0, created = 0, skipped = [];
    await Word.run(async function (ctx) {
      var body = ctx.document.body;
      var withPanel = body.search('Figure_[A-Za-z0-9]@_[A-Z]>', { matchWildcards: true });
      var noPanel = body.search('Figure_[A-Za-z0-9]@>', { matchWildcards: true });
      withPanel.load('items/text');
      noPanel.load('items/text');
      await ctx.sync();

      // Drop no-panel matches that sit inside a with-panel match ("Figure_droso" in "Figure_droso_A").
      var plain = noPanel.items;
      if (plain.length && withPanel.items.length) {
        var rel = plain.map(function (b) {
          return withPanel.items.map(function (a) { return b.compareLocationWith(a); });
        });
        await ctx.sync();
        plain = plain.filter(function (b, i) {
          return !rel[i].some(function (r) { return INSIDE[r.value]; });
        });
      }

      var matches = [];
      withPanel.items.concat(plain).forEach(function (range) {
        var p = C.parsePlaceholder(range.text);
        if (!p) return;
        if (C.validateLabel(p.label, []) !== null) { skipped.push(range.text); return; }
        matches.push({ range: range, p: p });
      });

      var byLabel = {};
      figures().forEach(function (f) { byLabel[f.label.toLowerCase()] = f; });
      matches.forEach(function (m) {
        var key = m.p.label.toLowerCase();
        if (!byLabel[key]) {
          var f = {
            id: C.genId(figures().map(function (x) { return x.id; })),
            seq: 'fig', label: m.p.label, description: '',
            manualOrder: C.nextManualOrder(figures())
          };
          figures().push(f);
          byLabel[key] = f;
          created++;
        }
        m.fig = byLabel[key];
      });

      // Each search result list is in document order and the two lists don't overlap,
      // so replacing back-to-front within each list keeps the remaining ranges valid.
      var withP = matches.filter(function (m) { return m.p.panel !== null; });
      var noP = matches.filter(function (m) { return m.p.panel === null; });
      var ordered = withP.reverse().concat(noP.reverse());
      for (var i = 0; i < ordered.length; i++) {
        var m = ordered[i];
        var panels = m.p.panel ? [m.p.panel] : [];
        var r = m.range.insertText(insertKindText('c', m.fig.id, panels, ''), 'Replace');
        var cc = r.insertContentControl();
        cc.tag = C.buildCitationTag(m.fig.id, panels, '');
        styleControl(cc, 'c', m.fig.id);
        converted++;
        if ((i + 1) % SYNC_BATCH === 0) await ctx.sync();
      }
      await ctx.sync();
    });
    await saveRegistry();
    await doRefresh();
    var msg = 'Imported ' + converted + ' placeholder' + (converted === 1 ? '' : 's') + ', created ' + created + ' figure' + (created === 1 ? '' : 's') + '.';
    if (skipped.length) msg += '\nSkipped (label too long): ' + skipped.join(', ');
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

  var CFG_INT = { numberStart: 1, panelRangeMinRun: 1 };

  function fillSettingsForm(c) {
    Array.prototype.forEach.call(document.querySelectorAll('#cfgForm [data-key]'), function (input) {
      var k = input.dataset.key;
      if (input.type === 'checkbox') input.checked = !!c[k];
      else input.value = c[k];
    });
  }

  function readSettingsForm() {
    var c = Object.assign({}, cfg());
    Array.prototype.forEach.call(document.querySelectorAll('#cfgForm [data-key]'), function (input) {
      var k = input.dataset.key;
      if (input.type === 'checkbox') c[k] = input.checked;
      else if (CFG_INT[k]) c[k] = /^\s*-?\d+\s*$/.test(input.value) ? parseInt(input.value, 10) : NaN;
      else c[k] = input.value;
    });
    return c;
  }

  function updateSettingsPreview() {
    var c = readSettingsForm();
    var errs = C.validateConfig(c, state.maxUsedPanel);
    var box = $('cfgErrors');
    box.hidden = !errs.length;
    box.textContent = errs.map(function (e) { return e.key + ': ' + e.message; }).join('\n');
    if (errs.length) {
      $('cfgPreview').textContent = '—';
    } else {
      $('cfgPreview').textContent = C.renderCitation(3, [1, 2, 3], '', c) + '   ·   ' + C.renderLegend(3, c);
    }
    var avail = $('panelsAvailable');
    if (C.isLetterStyle(c.panelStyle) && /^[A-Za-z]$/.test(String(c.panelStart))) {
      var m = C.maxPanels(c);
      avail.textContent = m + ' panels available (' + C.panelLabel(1, c) + '–' + C.panelLabel(m, c) + ')' +
        (state.maxUsedPanel ? '; the document uses up to ' + state.maxUsedPanel + '.' : '.');
    } else {
      avail.textContent = '';
    }
    $('cfgSave').disabled = !!errs.length;
    $('cfgDefaults').disabled = !!errs.length;
  }

  function openSettings() {
    clearMessage();
    fillSettingsForm(cfg());
    updateSettingsPreview();
    $('mainView').hidden = true;
    $('settingsView').hidden = false;
  }

  function closeSettings() {
    $('settingsView').hidden = true;
    $('mainView').hidden = false;
    $('panelText').value = C.formatPanelText(state.panels, cfg());
    $('panelError').hidden = true;
    renderAll();
  }

  async function saveSettings() {
    // Rescan so the panelStart check uses the document's real maximum panel.
    await Word.run(async function (ctx) {
      var scan = await scanControls(ctx);
      var max = 0;
      scan.items.forEach(function (it) { if (it.panels.length) max = Math.max(max, it.panels[it.panels.length - 1]); });
      state.maxUsedPanel = max;
    });
    var c = readSettingsForm();
    var errs = C.validateConfig(c, state.maxUsedPanel);
    if (errs.length) { updateSettingsPreview(); throw new Error('Fix the highlighted settings first.'); }
    state.registry.config = c;
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

  function sortedFigures() {
    if (cfg().orderMode === 'manual') return C.byManualOrder(figures());
    var nums = state.registry.lastRefresh.numbers;
    return figures().slice().sort(function (a, b) {
      var x = typeof nums[a.id] === 'number' ? nums[a.id] : Infinity;
      var y = typeof nums[b.id] === 'number' ? nums[b.id] : Infinity;
      return x - y;
    });
  }

  function renderFigures() {
    var list = $('figList');
    list.textContent = '';
    var figs = sortedFigures();
    $('figEmpty').hidden = figs.length > 0;
    var manual = cfg().orderMode === 'manual';
    var refreshed = !!state.registry.lastRefresh.at;
    var counts = state.registry.lastRefresh.counts || {};

    figs.forEach(function (f, idx) {
      var n = lastNumber(f.id);
      var count = counts[f.id] || 0;
      var li = el('li', { class: f.id === state.selectedFig ? 'selected' : '', title: f.description || f.label });
      li.addEventListener('click', function () { state.selectedFig = f.id; renderAll(); });

      if (state.editingFigRow === f.id) {
        var lab = el('input', { value: f.label, maxlength: '30' });
        var desc = el('input', { value: f.description, maxlength: '200', placeholder: 'Description' });
        var stop = function (e) { e.stopPropagation(); };
        lab.addEventListener('click', stop); desc.addEventListener('click', stop);
        li.appendChild(el('div', { class: 'editrow' }, [lab, desc, el('div', { class: 'row' }, [
          el('button', { class: 'primary', text: 'Save', onclick: function (e) {
            e.stopPropagation();
            action(function () { return renameFigure(f.id, lab.value.trim(), desc.value.trim()); });
          } }),
          el('button', { text: 'Cancel', onclick: function (e) { e.stopPropagation(); state.editingFigRow = null; renderFigures(); } })
        ])]));
        list.appendChild(li);
        return;
      }

      li.appendChild(el('span', { class: 'num', text: n === null ? '–' : C.formatNumber(n, cfg().numberStyle) }));
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
      btn('✎', 'Rename / edit', function () { state.editingFigRow = f.id; renderFigures(); });
      btn('✕', 'Delete', function () { action(function () { return deleteFigure(f.id); }); });
      li.appendChild(act);
      list.appendChild(li);
    });
  }

  function renderInsert() {
    var c = cfg();
    // Figure dropdown
    var sel = $('insFigure');
    sel.textContent = '';
    var figs = sortedFigures();
    if (!figs.length) sel.appendChild(el('option', { value: '', text: '(no figures)' }));
    figs.forEach(function (f) {
      var n = lastNumber(f.id);
      sel.appendChild(el('option', { value: f.id, text: (n === null ? '' : C.formatNumber(n, c.numberStyle) + ' · ') + f.label }));
    });
    if (state.selectedFig && C.figureById(state.registry, state.selectedFig)) sel.value = state.selectedFig;
    else if (figs.length) { state.selectedFig = figs[0].id; sel.value = state.selectedFig; }

    // Panel toggles
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
      $('btnInsert').disabled = !figs.length;
      $('btnLegend').disabled = !figs.length;
    }

    // Preview
    var fid = sel.value;
    var flags = $('numberOnly').checked ? 'n' : '';
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
    $('btnSettings').addEventListener('click', openSettings);
    $('btnImport').addEventListener('click', function () { action(doImport); });
    $('btnUnlink').addEventListener('click', function () { action(doUnlinkAll); });

    $('addForm').addEventListener('submit', function (e) {
      e.preventDefault();
      var label = $('addLabel').value.trim(), desc = $('addDesc').value.trim();
      action(async function () {
        await addFigure(label, desc);
        $('addLabel').value = '';
        $('addDesc').value = '';
      });
    });

    $('insFigure').addEventListener('change', function () { state.selectedFig = this.value; renderAll(); });
    $('panelText').addEventListener('input', function () {
      var parsed = C.parsePanelText(this.value, cfg());
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

    $('cfgForm').addEventListener('input', updateSettingsPreview);
    $('cfgForm').addEventListener('change', function (e) {
      if (e.target.dataset.key === 'panelStyle') {
        var start = document.querySelector('#cfgForm [data-key="panelStart"]');
        start.value = C.convertPanelStart(start.value, e.target.value);
      }
      updateSettingsPreview();
    });
    $('cfgSave').addEventListener('click', function () { action(saveSettings); });
    $('cfgDefaults').addEventListener('click', function () {
      var c = readSettingsForm();
      if (C.validateConfig(c).length) return;
      if (saveDefaults(c)) showMessage('Saved as your defaults for new documents.');
      else showMessage('Could not store defaults on this computer.', 'error');
    });
    $('cfgBack').addEventListener('click', closeSettings);
  }
})();
