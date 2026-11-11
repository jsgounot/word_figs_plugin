/*
 * SuppRef core — pure logic, no Office.js dependency.
 * Rendering (§8), tag format (§5.3–5.5), numbering (§6), validation (§5.2, §7).
 * Loaded by taskpane.html via <script>; also loadable by the test runner (tests/run.sh).
 * ES2017 only: no optional chaining, no nullish coalescing, no object spread.
 */
(function (root) {
  'use strict';

  var SCHEMA = 1;
  var SETTINGS_KEY = 'suppref';
  var TAG_PREFIX = 'suppref|';
  var MAX_LETTER_PANELS = 26;

  var DEFAULT_CONFIG = Object.freeze({
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
    orderMode: 'firstCitation',
    appearance: 'BoundingBox',
    autoRefreshOnInsert: true,
    suspendTrackChanges: true,
    warnUnresolved: true
  });

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
    return ['suppref', SCHEMA, 'c', figId, formatPanelsField(panels), flags || ''].join('|');
  }

  function buildLegendTag(figId) {
    return ['suppref', SCHEMA, 'l', figId].join('|');
  }

  var ID_RE = /^[0-9a-z]{1,32}$/;

  /**
   * §5.5. Returns null for foreign controls, { error } for malformed SuppRef tags,
   * or { kind: 'c'|'l', figId, panels, flags }.
   */
  function parseTag(tag) {
    if (typeof tag !== 'string' || tag.indexOf(TAG_PREFIX) !== 0) return null;
    var parts = tag.split('|');
    if (parts[1] !== String(SCHEMA)) return { error: 'Unsupported schema version "' + parts[1] + '"' };
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

  /** Returns an error string, or null if valid. */
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

  function newRegistry(config) {
    return {
      schema: SCHEMA,
      figures: [],
      config: mergeConfig(config),
      lastRefresh: { at: null, numbers: {}, counts: {} }
    };
  }

  function mergeConfig(cfg) {
    var out = Object.assign({}, DEFAULT_CONFIG);
    if (cfg && typeof cfg === 'object') {
      Object.keys(DEFAULT_CONFIG).forEach(function (k) {
        if (cfg[k] !== undefined && typeof cfg[k] === typeof DEFAULT_CONFIG[k]) out[k] = cfg[k];
      });
    }
    if (validateConfig(out).length) {
      // Fall back key by key for anything invalid.
      validateConfig(out).forEach(function (e) { out[e.key] = DEFAULT_CONFIG[e.key]; });
    }
    // The default panelStart ('A') is only valid for letter styles.
    if (validateConfig(out).length) out.panelStart = convertPanelStart(out.panelStart, out.panelStyle);
    return out;
  }

  /** Coerce whatever is stored in settings into a well-formed registry. */
  function normalizeRegistry(raw, defaultConfig) {
    if (typeof raw === 'string') {
      try { raw = JSON.parse(raw); } catch (e) { raw = null; }
    }
    if (!raw || typeof raw !== 'object') return newRegistry(defaultConfig);
    var reg = newRegistry(raw.config || defaultConfig);
    var seenIds = {};
    (Array.isArray(raw.figures) ? raw.figures : []).forEach(function (f, i) {
      if (!f || typeof f.id !== 'string' || !ID_RE.test(f.id) || seenIds[f.id]) return;
      seenIds[f.id] = true;
      reg.figures.push({
        id: f.id,
        seq: f.seq || 'fig',
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

  // ----------------------------------------------------------------- config

  function isInt(v) { return typeof v === 'number' && isFinite(v) && v === Math.floor(v); }

  /**
   * §7 validation. maxUsedPanel (optional) is the highest panel index used by any
   * citation in the document, so a panelStart that would push it past Z is rejected.
   * Returns [{ key, message }].
   */
  function validateConfig(cfg, maxUsedPanel) {
    var errs = [];
    var err = function (key, message) { errs.push({ key: key, message: message }); };
    ['prefix', 'legendPrefix', 'legendSuffix', 'numberPanelSeparator', 'panelRangeSeparator', 'panelListSeparator'].forEach(function (k) {
      if (typeof cfg[k] !== 'string') err(k, 'Must be text');
    });
    if (!isInt(cfg.numberStart) || cfg.numberStart < 0) err('numberStart', 'Must be a whole number ≥ 0');
    if (!isInt(cfg.panelRangeMinRun) || cfg.panelRangeMinRun < 2) err('panelRangeMinRun', 'Must be a whole number ≥ 2');
    Object.keys(ENUMS).forEach(function (k) {
      if (ENUMS[k].indexOf(cfg[k]) < 0) err(k, 'Invalid value');
    });
    ['autoRefreshOnInsert', 'suspendTrackChanges', 'warnUnresolved'].forEach(function (k) {
      if (typeof cfg[k] !== 'boolean') err(k, 'Must be on or off');
    });
    if (ENUMS.panelStyle.indexOf(cfg.panelStyle) >= 0) {
      var ps = String(cfg.panelStart);
      if (isLetterStyle(cfg.panelStyle)) {
        if (!/^[A-Za-z]$/.test(ps)) err('panelStart', 'Must be a single letter A–Z');
        else if (maxUsedPanel && maxUsedPanel > maxPanels(cfg)) {
          err('panelStart', 'Starting at ' + ps.toUpperCase() + ' leaves ' + maxPanels(cfg) +
            ' panels, but the document uses up to ' + maxUsedPanel);
        }
      } else if (!/^\d+$/.test(ps) || parseInt(ps, 10) < 1) {
        err('panelStart', 'Must be a whole number ≥ 1');
      }
    }
    return errs;
  }

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

  // -------------------------------------------------------------- numbering

  /**
   * §6. items: parsed controls in document order, each
   *   { kind: 'c'|'l', figId, panels, flags, ref }  (ref is opaque, e.g. the control id).
   * Returns { numbers: {figId: n}, counts: {figId: n}, warnings: [{ type, message, ref? }] }.
   */
  function computeNumbering(registry, items) {
    var cfg = registry.config;
    var known = {};
    registry.figures.forEach(function (f) { known[f.id] = f; });
    var numbers = {}, counts = {}, warnings = [];
    var firstCiteOrder = [], firstCiteRef = {};

    items.forEach(function (it) {
      if (it.kind !== 'c' || !known[it.figId]) return;
      counts[it.figId] = (counts[it.figId] || 0) + 1;
      if (!(it.figId in firstCiteRef)) { firstCiteRef[it.figId] = it.ref; firstCiteOrder.push(it.figId); }
    });

    var next = cfg.numberStart;
    if (cfg.orderMode === 'manual') {
      byManualOrder(registry.figures).forEach(function (f) { numbers[f.id] = next++; });
    } else {
      firstCiteOrder.forEach(function (id) { numbers[id] = next++; });
      registry.figures.forEach(function (f) { if (!(f.id in numbers)) numbers[f.id] = next++; });
    }

    var fmt = function (id) { return formatNumber(numbers[id], cfg.numberStyle); };
    var label = function (id) { return known[id] ? known[id].label : id; };

    if (cfg.warnUnresolved) {
      var uncited = registry.figures.filter(function (f) { return !counts[f.id]; });
      uncited.forEach(function (f) {
        warnings.push({ type: 'uncited', figId: f.id, message: 'Uncited: ' + f.label + ' (Fig. ' + fmt(f.id) + ')' });
      });

      var orphanSeen = {};
      items.forEach(function (it) {
        if (known[it.figId] || orphanSeen[it.figId]) return;
        orphanSeen[it.figId] = true;
        warnings.push({ type: 'orphan', figId: it.figId, ref: it.ref,
          message: 'Orphan ' + (it.kind === 'l' ? 'legend' : 'citation') + ': figure id "' + it.figId + '" is not registered' });
      });

      if (cfg.orderMode === 'manual') {
        var maxSoFar = null;
        firstCiteOrder.forEach(function (id) {
          if (maxSoFar !== null && numbers[id] < numbers[maxSoFar]) {
            warnings.push({ type: 'order', figId: id, ref: firstCiteRef[id],
              message: 'Fig. ' + fmt(maxSoFar) + ' is cited before Fig. ' + fmt(id) });
          }
          if (maxSoFar === null || numbers[id] > numbers[maxSoFar]) maxSoFar = id;
        });
      }

      var legendRefs = {}, prevLegend = null, orderWarned = {};
      items.forEach(function (it) {
        if (it.kind !== 'l' || !known[it.figId]) return;
        (legendRefs[it.figId] = legendRefs[it.figId] || []).push(it.ref);
        if (prevLegend && numbers[it.figId] < numbers[prevLegend.figId] &&
            !orderWarned[prevLegend.figId + '>' + it.figId]) {
          orderWarned[prevLegend.figId + '>' + it.figId] = true;
          warnings.push({ type: 'legendOrder', figId: prevLegend.figId, ref: prevLegend.ref,
            message: 'Legend for Fig. ' + fmt(prevLegend.figId) + ' appears before Fig. ' + fmt(it.figId) });
        }
        if (!prevLegend || numbers[it.figId] > numbers[prevLegend.figId]) prevLegend = it;
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
      if (renderPanels(it.panels, cfg).overflow.length) {
        overflowWarned[it.figId] = true;
        warnings.push({ type: 'panelOverflow', figId: it.figId, ref: it.ref,
          message: 'Panel beyond Z in a citation of ' + label(it.figId) + ' (rendered as "?")' });
      }
    });

    return { numbers: numbers, counts: counts, warnings: warnings };
  }

  /** Target text for one parsed control given computed numbers. */
  function targetText(item, registry, numbers) {
    var cfg = registry.config;
    var known = !!figureById(registry, item.figId);
    if (item.kind === 'l') return renderLegend(known ? numbers[item.figId] : null, cfg);
    if (!known) return renderOrphan(item.flags, cfg);
    return renderCitation(numbers[item.figId], item.panels, item.flags, cfg);
  }

  function citationTitle(label) { return 'SuppRef: ' + label; }
  function legendTitle(label) { return 'SuppRef legend: ' + label; }

  // ------------------------------------------------------------ placeholders

  /** §9.9. "Figure_droso_A" → { label: 'droso', panel: 1 }; "Figure_qpcr" → { label: 'qpcr', panel: null }. */
  function parsePlaceholder(text) {
    var m = /^Figure_([A-Za-z0-9]+)(?:_([A-Z]))?$/.exec(String(text).trim());
    if (!m) return null;
    return { label: m[1], panel: m[2] ? m[2].charCodeAt(0) - 64 : null };
  }

  var api = {
    SCHEMA: SCHEMA,
    SETTINGS_KEY: SETTINGS_KEY,
    TAG_PREFIX: TAG_PREFIX,
    DEFAULT_CONFIG: DEFAULT_CONFIG,
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
    newRegistry: newRegistry,
    mergeConfig: mergeConfig,
    normalizeRegistry: normalizeRegistry,
    nextManualOrder: nextManualOrder,
    figureById: figureById,
    byManualOrder: byManualOrder,
    validateConfig: validateConfig,
    convertPanelStart: convertPanelStart,
    computeNumbering: computeNumbering,
    targetText: targetText,
    citationTitle: citationTitle,
    legendTitle: legendTitle,
    parsePlaceholder: parsePlaceholder
  };

  root.SuppRefCore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
