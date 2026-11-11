/* Unit tests for core.js. Run with tests/run.sh (JavaScriptCore or Node). */
(function () {
  'use strict';
  var C = globalThis.SuppRefCore;
  var log = typeof print === 'function' ? print : console.log;
  var pass = 0, fail = 0;

  function eq(actual, expected, name) {
    var a = JSON.stringify(actual), e = JSON.stringify(expected);
    if (a === e) { pass++; }
    else { fail++; log('FAIL ' + name + '\n  expected ' + e + '\n  actual   ' + a); }
  }
  function cfg(over) { return Object.assign({}, C.DEFAULT_CONFIG, over || {}); }

  // §8.1 rendering examples
  var d = cfg();
  eq(C.renderCitation(3, [1], '', d), 'Supplementary Figure 3.A', '8.1 single panel');
  eq(C.renderCitation(3, [1, 2, 3], '', d), 'Supplementary Figure 3.A–C', '8.1 range');
  eq(C.renderCitation(3, [1, 2], '', d), 'Supplementary Figure 3.A,B', '8.1 short run');
  eq(C.renderCitation(3, [1, 2, 3, 5], '', d), 'Supplementary Figure 3.A–C,E', '8.1 mixed');
  eq(C.renderCitation(3, [], '', d), 'Supplementary Figure 3', '8.1 no panels');
  eq(C.renderCitation(3, [1], '', cfg({ prefix: 'Fig. S', numberPanelSeparator: '' })), 'Fig. S3A', '8.1 Fig. S');
  eq(C.renderCitation(3, [1], '', cfg({ panelStyle: 'lower' })), 'Supplementary Figure 3.a', '8.1 lower');
  eq(C.renderCitation(3, [1], '', cfg({ panelStart: 'B' })), 'Supplementary Figure 3.B', '8.1 start B');
  eq(C.renderCitation(3, [2], 'n', d), '3.B', '8.1 number only');
  eq(C.renderLegend(3, d), 'Supplementary Figure 3.', '8.1 legend');
  eq(C.renderCitation(null, [1], '', d), 'Supplementary Figure ??.A', 'never numbered');
  eq(C.renderOrphan('', d), 'Supplementary Figure ??', 'orphan');
  eq(C.renderCitation(1, [1, 2], '', cfg({ panelRangeMinRun: 2 })), 'Supplementary Figure 1.A–B', 'minRun 2');
  eq(C.renderCitation(4, [], '', cfg({ numberStyle: 'roman-upper' })), 'Supplementary Figure IV', 'roman number');
  eq(C.renderCitation(1, [1, 3], '', cfg({ panelStyle: 'roman-lower', panelStart: '1', panelListSeparator: ' and ' })),
    'Supplementary Figure 1.i and iii', 'roman panels');
  eq(C.renderCitation(1, [26], '', cfg({ panelStart: 'B' })), 'Supplementary Figure 1.?', 'overflow beyond Z');

  // Panel text box round trip
  eq(C.parsePanelText('A-C,E', d), { panels: [1, 2, 3, 5] }, 'parse A-C,E');
  eq(C.parsePanelText('a, c', d), { panels: [1, 3] }, 'parse lowercase');
  eq(C.parsePanelText('B–D', cfg({ panelStart: 'B' })), { panels: [1, 2, 3] }, 'parse with start B');
  eq(C.parsePanelText('A', cfg({ panelStart: 'B' })).panels, undefined, 'A invalid with start B');
  eq(C.parsePanelText('C-A', d).error !== undefined, true, 'reversed range');
  eq(C.formatPanelText([1, 2, 3, 5], d), 'A-C,E', 'format panel text');

  // Tags
  eq(C.buildCitationTag('k3f9qa', [5, 1, 2, 3, 3], 'n'), 'suppref|1|c|k3f9qa|1-3,5|n', 'build tag');
  eq(C.parseTag('suppref|1|c|k3f9qa|1-3,5|n'), { kind: 'c', figId: 'k3f9qa', panels: [1, 2, 3, 5], flags: 'n' }, 'parse tag');
  eq(C.parseTag('suppref|1|c|k3f9qa||'), { kind: 'c', figId: 'k3f9qa', panels: [], flags: '' }, 'parse no panels');
  eq(C.parseTag('suppref|1|l|k3f9qa'), { kind: 'l', figId: 'k3f9qa', panels: [], flags: '' }, 'parse legend');
  eq(C.parseTag('other|x'), null, 'foreign tag');
  eq(!!C.parseTag('suppref|2|c|a||').error, true, 'schema 2');
  eq(!!C.parseTag('suppref|1|c|a|x|').error, true, 'bad panels');
  eq(C.formatPanelsField([1, 2]), '1,2', 'panels field short run');

  // Validation
  var figs = [{ id: 'aaaaaa', label: 'Droso' }];
  eq(C.validateLabel('droso', figs), 'Label "droso" is already used', 'dup label');
  eq(C.validateLabel('droso', figs, 'aaaaaa'), null, 'rename self');
  eq(!!C.validateLabel('a b', figs), true, 'bad chars');
  eq(C.validateConfig(cfg({ panelStart: 'B' }), 26).length, 1, 'panelStart pushes past Z');
  eq(C.validateConfig(cfg({ panelStart: 'B' }), 25).length, 0, 'panelStart ok');
  eq(C.validateConfig(cfg({ numberStart: -1 })).length, 1, 'negative start');
  eq(C.mergeConfig({ panelStyle: 'arabic' }).panelStart, '1', 'merge fixes panelStart for arabic');
  eq(C.convertPanelStart('C', 'arabic'), '3', 'convert letter->num');
  eq(C.convertPanelStart('3', 'upper'), 'C', 'convert num->letter');

  // §13 test 1: numbering by first citation
  var reg = C.newRegistry();
  reg.figures = [
    { id: 'f1', seq: 'fig', label: 'fig1', description: '', manualOrder: 0 },
    { id: 'f2', seq: 'fig', label: 'fig2', description: '', manualOrder: 1 },
    { id: 'f3', seq: 'fig', label: 'fig3', description: '', manualOrder: 2 },
    { id: 'f4', seq: 'fig', label: 'fig4', description: '', manualOrder: 3 }
  ];
  var items = [
    { kind: 'c', figId: 'f2', panels: [], flags: '', ref: 0 },
    { kind: 'c', figId: 'f1', panels: [], flags: '', ref: 1 },
    { kind: 'c', figId: 'f2', panels: [2], flags: '', ref: 2 },
    { kind: 'c', figId: 'f3', panels: [1, 2, 3], flags: '', ref: 3 },
    { kind: 'c', figId: 'zz', panels: [], flags: '', ref: 4 },
    { kind: 'l', figId: 'f3', panels: [], flags: '', ref: 5 },
    { kind: 'l', figId: 'f2', panels: [], flags: '', ref: 6 },
    { kind: 'l', figId: 'f2', panels: [], flags: '', ref: 7 }
  ];
  var r = C.computeNumbering(reg, items);
  eq(r.numbers, { f2: 1, f1: 2, f3: 3, f4: 4 }, 'first-citation numbers');
  eq(r.counts, { f2: 2, f1: 1, f3: 1 }, 'counts');
  eq(r.warnings.map(function (w) { return w.type; }).sort(),
    ['duplicateLegend', 'legendOrder', 'orphan', 'uncited'], 'warnings');
  eq(C.targetText(items[2], reg, r.numbers), 'Supplementary Figure 1.B', 'target fig2.B');
  eq(C.targetText(items[3], reg, r.numbers), 'Supplementary Figure 3.A–C', 'target fig3.A–C');
  eq(C.targetText(items[4], reg, r.numbers), 'Supplementary Figure ??', 'target orphan');

  // Manual mode
  reg.config = cfg({ orderMode: 'manual' });
  var rm = C.computeNumbering(reg, items.slice(0, 4));
  eq(rm.numbers, { f1: 1, f2: 2, f3: 3, f4: 4 }, 'manual numbers');
  eq(rm.warnings.filter(function (w) { return w.type === 'order'; }).length, 1, 'manual order warning');

  // numberStart
  reg.config = cfg({ numberStart: 4 });
  eq(C.computeNumbering(reg, items.slice(0, 1)).numbers.f2, 4, 'numberStart 4');

  // Placeholders
  eq(C.parsePlaceholder('Figure_droso_A'), { label: 'droso', panel: 1 }, 'placeholder with panel');
  eq(C.parsePlaceholder('Figure_qpcr'), { label: 'qpcr', panel: null }, 'placeholder no panel');
  eq(C.parsePlaceholder('Figure_'), null, 'placeholder empty');

  // Registry normalization
  var n = C.normalizeRegistry('{"figures":[{"id":"abc"},{"id":"abc"},{"id":"BAD!"}]}');
  eq(n.figures.length, 1, 'normalize dedups and drops bad ids');
  eq(C.normalizeRegistry('not json').figures, [], 'normalize bad json');

  log(pass + ' passed, ' + fail + ' failed');
  if (fail && typeof quit === 'function') quit(1);
  if (fail && typeof process !== 'undefined') process.exit(1);
})();
