/*
 * SuppRef core — pure logic, no Office.js dependency.
 * Rendering (§8), tag format (§5.3–5.5), numbering (§6), validation (§5.2, §7),
 * and series: independent counters (main / supplementary / tables / custom).
 * Loaded by taskpane.html via <script>; also loadable by the test runner (tests/run.sh).
 * ES2017 only: no optional chaining, no nullish coalescing, no object spread.
 *
 * Every function that takes a "config"/"fmt" argument expects a series FORMAT
 * (prefix, number and panel styles…), not the document-wide settings.
 */
(function (root) {
  'use strict';

  var VERSION = '1.1.0';  // add-in version shown in Settings; bump on each release
  var TAG_SCHEMA = 1;     // version inside content-control tags (unchanged by series)
  var SCHEMA = 2;         // version of the registry stored in document settings
  var SETTINGS_KEY = 'suppref';
  var TAG_PREFIX = 'suppref|';
  var MAX_LETTER_PANELS = 26;

  /** Per-series formatting and numbering. */
  var DEFAULT_FORMAT = Object.freeze({
    prefix: 'Supplementary Figure ',
    legendPrefix: 'Supplementary Figure ',
    legendSuffix: '.',
    numberStart: 1,
    numberStyle: 'arabic',
    numberPanelSeparator: '.',
    panelStyle: 'upper',
    panelStart: 'A',
    panelRangeSeparator: '–',
    panelListSeparator: ',',
    panelRangeMinRun: 3,
    orderMode: 'firstCitation'
  });

  /** Document-wide behaviour. */
  var DEFAULT_GLOBAL = Object.freeze({
    appearance: 'BoundingBox',
    autoRefreshOnInsert: true,
    suspendTrackChanges: true,
    warnUnresolved: true
  });

  /**
   * Series offered in new documents. "fig" is the supplementary series: it keeps the
   * id used by schema-1 documents, so their figures land there on migration.
   */
  var DEFAULT_SERIES = [
    { id: 'main', name: 'Main figures', placeholder: 'MainFig', format: { prefix: 'Figure ', legendPrefix: 'Figure ' } },
    { id: 'fig', name: 'Supplementary figures', placeholder: 'Figure', format: {} },
    { id: 'ext', name: 'Additional figures', placeholder: 'ExtFig',
      format: { prefix: 'Extended Data Figure ', legendPrefix: 'Extended Data Figure ' } },
    { id: 'table', name: 'Supplementary tables', placeholder: 'SuppTable',
      format: { prefix: 'Supplementary Table ', legendPrefix: 'Supplementary Table ' } }
  ];

  var ENUMS = {
    numberStyle: ['arabic', 'roman-upper', 'roman-lower'],
    panelStyle: ['upper', 'lower', 'arabic', 'roman-lower'],
    orderMode: ['firstCitation', 'manual'],
    appearance: ['BoundingBox', 'Hidden']
  };

  // ---------------------------------------------------------------- numbers

  function toRoman(n) {
    if (!(n > 0) || n !== Math.floor(n)) return String(n);
    var map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
      [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    var out = '';
    for (var i = 0; i < map.length; i++) {
      while (n >= map[i][0]) { out += map[i][1]; n -= map[i][0]; }
    }
    return out;
  }

  function fromRoman(s) {
    s = String(s).toUpperCase();
    if (!/^[MDCLXVI]+$/.test(s)) return null;
    var vals = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
    var total = 0;
    for (var i = 0; i < s.length; i++) {
      var v = vals[s[i]], next = vals[s[i + 1]] || 0;
      total += v < next ? -v : v;
    }
    // Reject non-canonical forms such as "IIII" or "VX".
    return toRoman(total) === s ? total : null;
  }

  function formatNumber(n, style) {
    if (n === null || n === undefined) return '??';
    if (style === 'roman-upper') return toRoman(n);
    if (style === 'roman-lower') return toRoman(n).toLowerCase();
    return String(n);
  }

  // ----------------------------------------------------------------- panels

  function isLetterStyle(style) { return style === 'upper' || style === 'lower'; }

  /** 0-based letter offset (A=0) for letter styles, or integer start for numeric styles. */
  function panelStartValue(config) {
    var s = config.panelStart;
    if (isLetterStyle(config.panelStyle)) {
      var c = String(s).toUpperCase();
      if (/^[A-Z]$/.test(c)) return c.charCodeAt(0) - 65;
      var n = parseInt(s, 10);
      return (n >= 1 && n <= 26) ? n - 1 : 0;
    }
    if (/^[A-Za-z]$/.test(String(s))) return String(s).toUpperCase().charCodeAt(0) - 64;
    var v = parseInt(s, 10);
    return v >= 1 ? v : 1;
  }

  /** Number of panel indices that can be rendered (letter styles stop at Z). */
  function maxPanels(config) {
    if (isLetterStyle(config.panelStyle)) return MAX_LETTER_PANELS - panelStartValue(config);
    return Infinity;
  }

  /** Label of a 1-based panel index, or null if it can't be rendered (beyond Z). */
  function panelLabel(index, config) {
    var start = panelStartValue(config);
    switch (config.panelStyle) {
      case 'lower':
      case 'upper': {
        var code = start + index - 1;
        if (code < 0 || code >= 26) return null;
        var ch = String.fromCharCode(65 + code);
        return config.panelStyle === 'lower' ? ch.toLowerCase() : ch;
      }
      case 'arabic': return String(start + index - 1);
      case 'roman-lower': return toRoman(start + index - 1).toLowerCase();
      default: return null;
    }
  }

  /** Inverse of panelLabel. Returns a 1-based index or null. */
  function panelIndexFromLabel(label, config) {
    var start = panelStartValue(config);
    var s = String(label).trim();
    var v;
    if (isLetterStyle(config.panelStyle)) {
      if (!/^[A-Za-z]$/.test(s)) return null;
      v = s.toUpperCase().charCodeAt(0) - 65 - start + 1;
    } else if (config.panelStyle === 'arabic') {
      if (!/^\d+$/.test(s)) return null;
      v = parseInt(s, 10) - start + 1;
    } else {
      var r = fromRoman(s);
      if (r === null) return null;
      v = r - start + 1;
    }
    return v >= 1 ? v : null;
  }

  function normalizePanels(panels) {
    var seen = {}, out = [];
    (panels || []).forEach(function (p) {
      p = Number(p);
      if (p >= 1 && p === Math.floor(p) && !seen[p]) { seen[p] = true; out.push(p); }
    });
    return out.sort(function (a, b) { return a - b; });
  }

  /** Split sorted indices into maximal runs of consecutive values. */
  function toRuns(sorted) {
    var runs = [];
    sorted.forEach(function (p) {
      var last = runs[runs.length - 1];
      if (last && p === last[last.length - 1] + 1) last.push(p);
      else runs.push([p]);
    });
    return runs;
  }

  /** Returns { text, overflow } where overflow lists indices that rendered as "?". */
  function renderPanels(panels, config) {
    var sorted = normalizePanels(panels);
    var overflow = [];
    var lab = function (i) {
      var l = panelLabel(i, config);
      if (l === null) { overflow.push(i); return '?'; }
      return l;
    };
    var minRun = Math.max(2, parseInt(config.panelRangeMinRun, 10) || 3);
    var pieces = [];
    toRuns(sorted).forEach(function (run) {
      if (run.length >= minRun) pieces.push(lab(run[0]) + config.panelRangeSeparator + lab(run[run.length - 1]));
      else run.forEach(function (i) { pieces.push(lab(i)); });
    });
    return { text: pieces.join(config.panelListSeparator), overflow: overflow };
  }

  /** §8. number may be null (never numbered) → "??". */
  function renderCitation(number, panels, flags, config) {
    var num = formatNumber(number, config.numberStyle);
    var prefix = (flags || '').indexOf('n') >= 0 ? '' : config.prefix;
    var p = normalizePanels(panels);
    if (!p.length) return prefix + num;
    return prefix + num + config.numberPanelSeparator + renderPanels(p, config).text;
  }

  function renderOrphan(flags, config) {
    return ((flags || '').indexOf('n') >= 0 ? '' : config.prefix) + '??';
  }

  function renderLegend(number, config) {
    return config.legendPrefix + formatNumber(number, config.numberStyle) + config.legendSuffix;
  }

  /**
   * Parse what the user typed in the panel text box ("A-C,E", "a, c", "i–iii and v")
   * into sorted 1-based indices. Returns { panels } or { error }.
   */
  function parsePanelText(text, config) {
    var s = String(text || '').trim();
    if (!s) return { panels: [] };
    var tokens = s.split(/\s*(?:,|;|&|\band\b|\s)\s*/i).filter(Boolean);
    var out = [];
    for (var i = 0; i < tokens.length; i++) {
      var t = tokens[i];
      var m = t.split(/[-–—]/);
      if (m.length === 1) {
        var idx = panelIndexFromLabel(m[0], config);
        if (idx === null) return { error: 'Unknown panel "' + t + '"' };
        out.push(idx);
      } else if (m.length === 2 && m[0] && m[1]) {
        var a = panelIndexFromLabel(m[0], config), b = panelIndexFromLabel(m[1], config);
        if (a === null || b === null) return { error: 'Unknown panel range "' + t + '"' };
        if (b < a) return { error: 'Reversed range "' + t + '"' };
        if (b - a > 200) return { error: 'Range too large "' + t + '"' };
        for (var k = a; k <= b; k++) out.push(k);
      } else {
        return { error: 'Cannot read "' + t + '"' };
      }
    }
    return { panels: normalizePanels(out) };
  }

  /** Format indices as the rendered-label text shown in the panel text box (always "A-C,E" style). */
  function formatPanelText(panels, config) {
    var cfg = Object.assign({}, config, { panelRangeSeparator: '-', panelListSeparator: ',', panelRangeMinRun: 3 });
    return renderPanels(panels, cfg).text;
  }

  // ------------------------------------------------------------------- tags

  /** Storage form of panel indices in tags: "", "2", "1-3", "1,3,4", "1-3,5". */
  function formatPanelsField(panels) {
    return toRuns(normalizePanels(panels)).map(function (run) {
      if (run.length >= 3) return run[0] + '-' + run[run.length - 1];
      return run.join(',');
    }).join(',');
  }

  function parsePanelsField(field) {
    if (field === '') return [];
    var out = [];
    var parts = field.split(',');
    for (var i = 0; i < parts.length; i++) {
      var m = /^(\d+)(?:-(\d+))?$/.exec(parts[i]);
      if (!m) return null;
      var a = parseInt(m[1], 10), b = m[2] ? parseInt(m[2], 10) : a;
      if (a < 1 || b < a || b - a > 1000) return null;
      for (var k = a; k <= b; k++) out.push(k);
    }
    return normalizePanels(out);
  }

  function buildCitationTag(figId, panels, flags) {
    return ['suppref', TAG_SCHEMA, 'c', figId, formatPanelsField(panels), flags || ''].join('|');
  }

  function buildLegendTag(figId) {
    return ['suppref', TAG_SCHEMA, 'l', figId].join('|');
  }

  var ID_RE = /^[0-9a-z]{1,32}$/;

  /**
   * §5.5. Returns null for foreign controls, { error } for malformed SuppRef tags,
   * or { kind: 'c'|'l', figId, panels, flags }.
   */
  function parseTag(tag) {
    if (typeof tag !== 'string' || tag.indexOf(TAG_PREFIX) !== 0) return null;
    var parts = tag.split('|');
    if (parts[1] !== String(TAG_SCHEMA)) return { error: 'Unsupported schema version "' + parts[1] + '"' };
    var kind = parts[2];
    if (kind === 'c') {
      if (parts.length !== 6) return { error: 'Malformed citation tag' };
      if (!ID_RE.test(parts[3])) return { error: 'Malformed figure id' };
      var panels = parsePanelsField(parts[4]);
      if (panels === null) return { error: 'Malformed panel list' };
      if (parts[5] !== '' && parts[5] !== 'n') return { error: 'Unknown flags "' + parts[5] + '"' };
      return { kind: 'c', figId: parts[3], panels: panels, flags: parts[5] };
    }
    if (kind === 'l') {
      if (parts.length !== 4) return { error: 'Malformed legend tag' };
      if (!ID_RE.test(parts[3])) return { error: 'Malformed figure id' };
      return { kind: 'l', figId: parts[3], panels: [], flags: '' };
    }
    return { error: 'Unknown control kind "' + kind + '"' };
  }


  // --------------------------------------------------------------- registry

  function genId(existingIds, rand) {
    rand = rand || Math.random;
    var taken = {};
    (existingIds || []).forEach(function (id) { taken[id] = true; });
    for (var tries = 0; tries < 1000; tries++) {
      var id = '';
      for (var i = 0; i < 6; i++) id += '0123456789abcdefghijklmnopqrstuvwxyz'[Math.floor(rand() * 36)];
      if (!taken[id]) return id;
    }
    throw new Error('Could not generate a unique id');
  }

  /** Figure labels are unique across all series. Returns an error string, or null if valid. */
  function validateLabel(label, figures, excludeId) {
    if (typeof label !== 'string' || !label.length) return 'Label is required';
    if (label.length > 30) return 'Label must be at most 30 characters';
    if (!/^[A-Za-z0-9_-]+$/.test(label)) return 'Label may only contain letters, digits, _ and -';
    var low = label.toLowerCase();
    for (var i = 0; i < figures.length; i++) {
      if (figures[i].id !== excludeId && figures[i].label.toLowerCase() === low) return 'Label "' + label + '" is already used';
    }
    return null;
  }

  function validateDescription(d) {
    if (d && String(d).length > 200) return 'Description must be at most 200 characters';
    return null;
  }

  function nextManualOrder(figures) {
    var max = -1;
    figures.forEach(function (f) { if (f.manualOrder > max) max = f.manualOrder; });
    return max + 1;
  }

  function figureById(registry, id) {
    for (var i = 0; i < registry.figures.length; i++) if (registry.figures[i].id === id) return registry.figures[i];
    return null;
  }

  function byManualOrder(figures) {
    return figures.slice().sort(function (a, b) { return a.manualOrder - b.manualOrder; });
  }

  // ----------------------------------------------------------------- series

  function seriesById(registry, id) {
    for (var i = 0; i < registry.series.length; i++) if (registry.series[i].id === id) return registry.series[i];
    return null;
  }

  /** Series used for orphans and as the fallback: supplementary ("fig") if present, else the first. */
  function defaultSeries(registry) {
    return seriesById(registry, 'fig') || registry.series[0];
  }

  /** Format of the series a figure belongs to (fallback series for unknown figures). */
  function formatOf(registry, figId) {
    var f = figureById(registry, figId);
    var s = f ? seriesById(registry, f.seq) : null;
    return (s || defaultSeries(registry)).format;
  }

  function makeSeries(id, name, placeholder, format) {
    return { id: id, name: name, placeholder: placeholder, format: mergeFormat(format) };
  }

  function defaultSeriesList() {
    return DEFAULT_SERIES.map(function (s) {
      return makeSeries(s.id, s.name, s.placeholder, Object.assign({}, DEFAULT_FORMAT, s.format));
    });
  }

  /** Name, and placeholder token for "Import placeholders". Returns [{ key, message }]. */
  function validateSeriesMeta(series, allSeries) {
    var errs = [];
    var name = typeof series.name === 'string' ? series.name.trim() : '';
    if (!name) errs.push({ key: 'name', message: 'Name is required' });
    else if (name.length > 40) errs.push({ key: 'name', message: 'Name must be at most 40 characters' });
    var ph = typeof series.placeholder === 'string' ? series.placeholder : '';
    if (ph && !/^[A-Za-z][A-Za-z0-9]{0,19}$/.test(ph)) {
      errs.push({ key: 'placeholder', message: 'Placeholder word: a letter then letters/digits, up to 20 (or empty)' });
    }
    (allSeries || []).forEach(function (o) {
      if (o.id === series.id) return;
      if (name && o.name.trim().toLowerCase() === name.toLowerCase()) errs.push({ key: 'name', message: 'Another series is called "' + name + '"' });
      if (ph && o.placeholder && o.placeholder.toLowerCase() === ph.toLowerCase()) {
        errs.push({ key: 'placeholder', message: 'Placeholder word "' + ph + '" is used by ' + o.name });
      }
    });
    return errs;
  }

  // ----------------------------------------------------------------- config

  function isInt(v) { return typeof v === 'number' && isFinite(v) && v === Math.floor(v); }

  /**
   * §7 validation of a series format. maxUsedPanel (optional) is the highest panel index
   * used by any citation of this series, so a panelStart that would push it past Z is rejected.
   * Returns [{ key, message }].
   */
  function validateFormat(fmt, maxUsedPanel) {
    var errs = [];
    var err = function (key, message) { errs.push({ key: key, message: message }); };
    ['prefix', 'legendPrefix', 'legendSuffix', 'numberPanelSeparator', 'panelRangeSeparator', 'panelListSeparator'].forEach(function (k) {
      if (typeof fmt[k] !== 'string') err(k, 'Must be text');
    });
    if (!isInt(fmt.numberStart) || fmt.numberStart < 0) err('numberStart', 'Must be a whole number ≥ 0');
    if (!isInt(fmt.panelRangeMinRun) || fmt.panelRangeMinRun < 2) err('panelRangeMinRun', 'Must be a whole number ≥ 2');
    ['numberStyle', 'panelStyle', 'orderMode'].forEach(function (k) {
      if (ENUMS[k].indexOf(fmt[k]) < 0) err(k, 'Invalid value');
    });
    if (ENUMS.panelStyle.indexOf(fmt.panelStyle) >= 0) {
      var ps = String(fmt.panelStart);
      if (isLetterStyle(fmt.panelStyle)) {
        if (!/^[A-Za-z]$/.test(ps)) err('panelStart', 'Must be a single letter A–Z');
        else if (maxUsedPanel && maxUsedPanel > maxPanels(fmt)) {
          err('panelStart', 'Starting at ' + ps.toUpperCase() + ' leaves ' + maxPanels(fmt) +
            ' panels, but the document uses up to ' + maxUsedPanel);
        }
      } else if (!/^\d+$/.test(ps) || parseInt(ps, 10) < 1) {
        err('panelStart', 'Must be a whole number ≥ 1');
      }
    }
    return errs;
  }

  function validateGlobal(cfg) {
    var errs = [];
    if (ENUMS.appearance.indexOf(cfg.appearance) < 0) errs.push({ key: 'appearance', message: 'Invalid value' });
    ['autoRefreshOnInsert', 'suspendTrackChanges', 'warnUnresolved'].forEach(function (k) {
      if (typeof cfg[k] !== 'boolean') errs.push({ key: k, message: 'Must be on or off' });
    });
    return errs;
  }

  /** Take known keys of the right type from src over defaults, then reset invalid ones. */
  function mergeKeys(defaults, src, validate) {
    var out = Object.assign({}, defaults);
    if (src && typeof src === 'object') {
      Object.keys(defaults).forEach(function (k) {
        if (src[k] !== undefined && typeof src[k] === typeof defaults[k]) out[k] = src[k];
      });
    }
    validate(out).forEach(function (e) { out[e.key] = defaults[e.key]; });
    return out;
  }

  function mergeFormat(fmt) {
    var out = mergeKeys(DEFAULT_FORMAT, fmt, validateFormat);
    // The default panelStart ('A') is only valid for letter styles.
    if (validateFormat(out).length) out.panelStart = convertPanelStart(out.panelStart, out.panelStyle);
    return out;
  }

  function mergeGlobal(cfg) { return mergeKeys(DEFAULT_GLOBAL, cfg, validateGlobal); }

  /** Convert panelStart when the panel style switches between letter and numeric styles. */
  function convertPanelStart(value, newStyle) {
    var s = String(value);
    if (isLetterStyle(newStyle)) {
      if (/^[A-Za-z]$/.test(s)) return s.toUpperCase();
      var n = parseInt(s, 10);
      return (n >= 1 && n <= 26) ? String.fromCharCode(64 + n) : 'A';
    }
    if (/^[A-Za-z]$/.test(s)) return String(s.toUpperCase().charCodeAt(0) - 64);
    return /^\d+$/.test(s) && parseInt(s, 10) >= 1 ? s : '1';
  }

  /** Normalize a stored series list; falls back to the defaults when empty or invalid. */
  function normalizeSeriesList(list) {
    var out = [], seen = {};
    (Array.isArray(list) ? list : []).forEach(function (s) {
      if (!s || typeof s.id !== 'string' || !ID_RE.test(s.id) || seen[s.id]) return;
      seen[s.id] = true;
      var name = typeof s.name === 'string' && s.name.trim() ? s.name.trim().slice(0, 40) : 'Series ' + s.id;
      var ph = typeof s.placeholder === 'string' && /^[A-Za-z][A-Za-z0-9]{0,19}$/.test(s.placeholder) ? s.placeholder : '';
      out.push(makeSeries(s.id, name, ph, s.format));
    });
    return out.length ? out : defaultSeriesList();
  }

  /**
   * "Save as my defaults" stores { series, config }. Defaults saved by v1 were a flat config;
   * those become the supplementary format plus global settings.
   */
  function seriesAndConfigFrom(src) {
    if (src && Array.isArray(src.series)) {
      return { series: normalizeSeriesList(src.series), config: mergeGlobal(src.config) };
    }
    var series = defaultSeriesList();
    if (src && typeof src === 'object') {
      var supp = series.filter(function (s) { return s.id === 'fig'; })[0];
      supp.format = mergeFormat(Object.assign({}, supp.format, src));
    }
    return { series: series, config: mergeGlobal(src) };
  }

  function newRegistry(defaults) {
    var sc = seriesAndConfigFrom(defaults);
    return {
      schema: SCHEMA,
      series: sc.series,
      figures: [],
      config: sc.config,
      lastRefresh: { at: null, numbers: {}, counts: {} }
    };
  }

  /**
   * Coerce whatever is stored in settings into a well-formed registry.
   * Schema 1 (no series): its config becomes the supplementary format, and its
   * figures (seq "fig") stay supplementary figures.
   */
  function normalizeRegistry(raw, defaults) {
    if (typeof raw === 'string') {
      try { raw = JSON.parse(raw); } catch (e) { raw = null; }
    }
    if (!raw || typeof raw !== 'object') return newRegistry(defaults);
    var reg = newRegistry(Array.isArray(raw.series) ? { series: raw.series, config: raw.config } : (raw.config || {}));
    var seqs = {};
    reg.series.forEach(function (s) { seqs[s.id] = true; });
    var fallback = defaultSeries(reg).id;
    var seenIds = {};
    (Array.isArray(raw.figures) ? raw.figures : []).forEach(function (f, i) {
      if (!f || typeof f.id !== 'string' || !ID_RE.test(f.id) || seenIds[f.id]) return;
      seenIds[f.id] = true;
      reg.figures.push({
        id: f.id,
        seq: seqs[f.seq] ? f.seq : fallback,
        label: typeof f.label === 'string' && f.label ? f.label : 'fig-' + f.id,
        description: typeof f.description === 'string' ? f.description : '',
        manualOrder: typeof f.manualOrder === 'number' ? f.manualOrder : i
      });
    });
    if (raw.lastRefresh && typeof raw.lastRefresh === 'object') {
      reg.lastRefresh = {
        at: raw.lastRefresh.at || null,
        numbers: raw.lastRefresh.numbers || {},
        counts: raw.lastRefresh.counts || {}
      };
    }
    return reg;
  }

  // -------------------------------------------------------------- numbering

  /**
   * §6, once per series. items: parsed controls in document order, each
   *   { kind: 'c'|'l', figId, panels, flags, ref }  (ref is opaque, e.g. the control id).
   * Returns { numbers: {figId: n}, counts: {figId: n}, warnings: [{ type, message, figId?, ref? }] }.
   */
  function computeNumbering(registry, items) {
    var known = {};
    registry.figures.forEach(function (f) { known[f.id] = f; });
    var numbers = {}, counts = {}, warnings = [];
    var firstCiteOrder = [], firstCiteRef = {};

    items.forEach(function (it) {
      if (it.kind !== 'c' || !known[it.figId]) return;
      counts[it.figId] = (counts[it.figId] || 0) + 1;
      if (!(it.figId in firstCiteRef)) { firstCiteRef[it.figId] = it.ref; firstCiteOrder.push(it.figId); }
    });

    registry.series.forEach(function (s) {
      var next = s.format.numberStart;
      var figs = registry.figures.filter(function (f) { return f.seq === s.id; });
      if (s.format.orderMode === 'manual') {
        byManualOrder(figs).forEach(function (f) { numbers[f.id] = next++; });
      } else {
        firstCiteOrder.forEach(function (id) { if (known[id].seq === s.id) numbers[id] = next++; });
        figs.forEach(function (f) { if (!(f.id in numbers)) numbers[f.id] = next++; });
      }
    });

    /** "Supplementary Figure 3" — how warnings name a figure. */
    var name = function (id) {
      var fmt = formatOf(registry, id);
      return (fmt.prefix + formatNumber(numbers[id], fmt.numberStyle)).trim();
    };
    var label = function (id) { return known[id] ? known[id].label : id; };
    var seqOf = function (id) { return known[id].seq; };

    if (registry.config.warnUnresolved) {
      registry.figures.forEach(function (f) {
        if (!counts[f.id]) warnings.push({ type: 'uncited', figId: f.id, message: 'Uncited: ' + f.label + ' (' + name(f.id) + ')' });
      });

      var orphanSeen = {};
      items.forEach(function (it) {
        if (known[it.figId] || orphanSeen[it.figId]) return;
        orphanSeen[it.figId] = true;
        warnings.push({ type: 'orphan', figId: it.figId, ref: it.ref,
          message: 'Orphan ' + (it.kind === 'l' ? 'legend' : 'citation') + ': figure id "' + it.figId + '" is not registered' });
      });

      // Manual mode: citation order that disagrees with the chosen order (per series).
      var maxSoFar = {};
      firstCiteOrder.forEach(function (id) {
        var s = seriesById(registry, seqOf(id));
        if (!s || s.format.orderMode !== 'manual') return;
        var prev = maxSoFar[s.id];
        if (prev && numbers[id] < numbers[prev]) {
          warnings.push({ type: 'order', figId: id, ref: firstCiteRef[id],
            message: name(prev) + ' is cited before ' + name(id) });
        }
        if (!prev || numbers[id] > numbers[prev]) maxSoFar[s.id] = id;
      });

      // Legend labels: ascending within each series; at most one per figure.
      var legendRefs = {}, prevLegend = {}, orderWarned = {};
      items.forEach(function (it) {
        if (it.kind !== 'l' || !known[it.figId]) return;
        (legendRefs[it.figId] = legendRefs[it.figId] || []).push(it.ref);
        var seq = seqOf(it.figId), prev = prevLegend[seq];
        if (prev && numbers[it.figId] < numbers[prev.figId] && !orderWarned[prev.figId + '>' + it.figId]) {
          orderWarned[prev.figId + '>' + it.figId] = true;
          warnings.push({ type: 'legendOrder', figId: prev.figId, ref: prev.ref,
            message: 'Legend for ' + name(prev.figId) + ' appears before ' + name(it.figId) });
        }
        if (!prev || numbers[it.figId] > numbers[prev.figId]) prevLegend[seq] = it;
      });
      Object.keys(legendRefs).forEach(function (id) {
        if (legendRefs[id].length > 1) {
          warnings.push({ type: 'duplicateLegend', figId: id, ref: legendRefs[id][1],
            message: 'Duplicate legend for ' + label(id) + ' (' + legendRefs[id].length + ' legends)' });
        }
      });
    }

    var overflowWarned = {};
    items.forEach(function (it) {
      if (it.kind !== 'c' || !known[it.figId] || overflowWarned[it.figId]) return;
      if (renderPanels(it.panels, formatOf(registry, it.figId)).overflow.length) {
        overflowWarned[it.figId] = true;
        warnings.push({ type: 'panelOverflow', figId: it.figId, ref: it.ref,
          message: 'Panel beyond Z in a citation of ' + label(it.figId) + ' (rendered as "?")' });
      }
    });

    return { numbers: numbers, counts: counts, warnings: warnings };
  }

  /** Target text for one parsed control given computed numbers. */
  function targetText(item, registry, numbers) {
    var known = !!figureById(registry, item.figId);
    var fmt = formatOf(registry, item.figId);
    if (item.kind === 'l') return renderLegend(known ? numbers[item.figId] : null, fmt);
    if (!known) return renderOrphan(item.flags, fmt);
    return renderCitation(numbers[item.figId], item.panels, item.flags, fmt);
  }

  function citationTitle(label) { return 'SuppRef: ' + label; }
  function legendTitle(label) { return 'SuppRef legend: ' + label; }

  // ------------------------------------------------------------ placeholders

  /**
   * §9.9, one pair of Word wildcard patterns per series with a placeholder word.
   * "<" anchors the word start so "Figure_x" isn't found inside "MainFigure_x".
   */
  function placeholderPatterns(seriesList) {
    var out = [];
    seriesList.forEach(function (s) {
      if (!s.placeholder) return;
      out.push({ seq: s.id, withPanel: true, pattern: '<' + s.placeholder + '_[A-Za-z0-9]@_[A-Z]>' });
      out.push({ seq: s.id, withPanel: false, pattern: '<' + s.placeholder + '_[A-Za-z0-9]@>' });
    });
    return out;
  }

  /**
   * "Figure_droso_A" → { seq: 'fig', label: 'droso', panel: 1 };
   * "MainFig_qpcr" → { seq: 'main', label: 'qpcr', panel: null }. Placeholder words are case-sensitive.
   */
  function parsePlaceholder(text, seriesList) {
    var m = /^([A-Za-z][A-Za-z0-9]*)_([A-Za-z0-9]+)(?:_([A-Z]))?$/.exec(String(text).trim());
    if (!m) return null;
    for (var i = 0; i < seriesList.length; i++) {
      if (seriesList[i].placeholder && seriesList[i].placeholder === m[1]) {
        return { seq: seriesList[i].id, label: m[2], panel: m[3] ? m[3].charCodeAt(0) - 64 : null };
      }
    }
    return null;
  }

  var api = {
    VERSION: VERSION,
    SCHEMA: SCHEMA,
    TAG_SCHEMA: TAG_SCHEMA,
    SETTINGS_KEY: SETTINGS_KEY,
    TAG_PREFIX: TAG_PREFIX,
    DEFAULT_FORMAT: DEFAULT_FORMAT,
    DEFAULT_GLOBAL: DEFAULT_GLOBAL,
    ENUMS: ENUMS,
    toRoman: toRoman,
    fromRoman: fromRoman,
    formatNumber: formatNumber,
    isLetterStyle: isLetterStyle,
    panelStartValue: panelStartValue,
    maxPanels: maxPanels,
    panelLabel: panelLabel,
    panelIndexFromLabel: panelIndexFromLabel,
    normalizePanels: normalizePanels,
    renderPanels: renderPanels,
    renderCitation: renderCitation,
    renderOrphan: renderOrphan,
    renderLegend: renderLegend,
    parsePanelText: parsePanelText,
    formatPanelText: formatPanelText,
    formatPanelsField: formatPanelsField,
    parsePanelsField: parsePanelsField,
    buildCitationTag: buildCitationTag,
    buildLegendTag: buildLegendTag,
    parseTag: parseTag,
    genId: genId,
    validateLabel: validateLabel,
    validateDescription: validateDescription,
    nextManualOrder: nextManualOrder,
    figureById: figureById,
    byManualOrder: byManualOrder,
    seriesById: seriesById,
    defaultSeries: defaultSeries,
    formatOf: formatOf,
    makeSeries: makeSeries,
    defaultSeriesList: defaultSeriesList,
    validateSeriesMeta: validateSeriesMeta,
    validateFormat: validateFormat,
    validateGlobal: validateGlobal,
    mergeFormat: mergeFormat,
    mergeGlobal: mergeGlobal,
    convertPanelStart: convertPanelStart,
    normalizeSeriesList: normalizeSeriesList,
    newRegistry: newRegistry,
    normalizeRegistry: normalizeRegistry,
    computeNumbering: computeNumbering,
    targetText: targetText,
    citationTitle: citationTitle,
    legendTitle: legendTitle,
    placeholderPatterns: placeholderPatterns,
    parsePlaceholder: parsePlaceholder
  };

  root.SuppRefCore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
