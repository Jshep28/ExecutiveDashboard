/**
 * data-store.js  v4
 * - No hardcoded KPIs — starts empty, client adds their own
 * - isKey flag drives Overview page
 * - monthlyActuals: { "Jul":val, "Aug":val, ... } per KPI for FY monthly tracking
 * - fyStartMonth in settings drives which months are generated
 */

const DataStore = (() => {
  const STORAGE_KEY   = 'exec_kpis_v4';
  const SETTINGS_KEY  = 'exec_settings_v4';
  const THRESHOLD_KEY = 'exec_thresholds_v4';
  const FORMULA_KEY   = 'exec_formulas_v1';

  const DEFAULT_THRESHOLDS = [
    { id:'th_higher_better', name:'Higher is Better (90% / 70%)', description:'Green ≥90% of target, Amber ≥70%, Red below 70%', type:'relative',
      levels:[{rag:'green',op:'>=',value:0.90,label:'≥ 90% of target'},{rag:'amber',op:'>=',value:0.70,label:'≥ 70% of target'},{rag:'red',op:'<',value:0.70,label:'< 70% of target'}] },
    { id:'th_lower_better', name:'Lower is Better (at target / +20%)', description:'Green ≤ target, Amber ≤ target+20%, Red above', type:'relative_lower',
      levels:[{rag:'green',op:'<=',value:1.00,label:'≤ target'},{rag:'amber',op:'<=',value:1.20,label:'≤ target + 20%'},{rag:'red',op:'>',value:1.20,label:'> target + 20%'}] },
    { id:'th_percent_high', name:'Percentage — High Good (80% / 60%)', description:'Absolute % where higher is better', type:'absolute',
      levels:[{rag:'green',op:'>=',value:0.80,label:'≥ 80%'},{rag:'amber',op:'>=',value:0.60,label:'≥ 60%'},{rag:'red',op:'<',value:0.60,label:'< 60%'}] },
    { id:'th_percent_low', name:'Percentage — Low Good (10% / 20%)', description:'Absolute % where lower is better (churn, turnover)', type:'absolute',
      levels:[{rag:'green',op:'<=',value:0.10,label:'≤ 10%'},{rag:'amber',op:'<=',value:0.20,label:'≤ 20%'},{rag:'red',op:'>',value:0.20,label:'> 20%'}] },
    { id:'th_manual_only', name:'Manual RAG Only', description:'No auto-calculation — set RAG manually', type:'manual', levels:[] },
  ];

  // Month abbreviations in calendar order
  const ALL_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  // ── State ─────────────────────────────────────────────────────────────────
  let _kpis        = [];
  let _settings    = {};
  let _thresholds  = [];
  let _formulas    = [];   // formula KPI definitions
  let _subscribers = [];
  let _nextId      = 1;
  let _nextFmlId   = 1;

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    _settings   = _loadSettings();
    _thresholds = _loadThresholds();
    _kpis       = _loadKpis();
    _formulas   = _loadFormulas();
    _computeRag();
    _applyAllFormulas();
  }

  function _loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      const d = {
        companyName: 'My Company',
        fiscalYearLabel: 'FY26',
        reportingPeriod: 'monthly',
        fyStartMonth: 'Jul',      // fiscal year start month (default July for AU)
        decimals:       2,        // decimal places: 0 | 1 | 2 | 3
        largeNumFormat: 'auto',   // 'auto' | 'M' | 'K' | 'full'
        currencySymbol: '$',      // '$' | '€' | '£' | '¥' | custom
        pctStorage:     'decimal',// 'decimal' (0.85→85%) | 'direct' (85→85%)
        overviewKpiIds: null,     // null = show all key KPIs; array = filtered subset
      };
      return raw ? { ...d, ...JSON.parse(raw) } : d;
    } catch {
      return { companyName:'My Company', fiscalYearLabel:'FY26', reportingPeriod:'monthly', fyStartMonth:'Jul',
               decimals:2, largeNumFormat:'auto', currencySymbol:'$', pctStorage:'decimal', overviewKpiIds:null };
    }
  }

  function _loadThresholds() {
    try {
      const raw = localStorage.getItem(THRESHOLD_KEY);
      return raw ? JSON.parse(raw) : JSON.parse(JSON.stringify(DEFAULT_THRESHOLDS));
    } catch { return JSON.parse(JSON.stringify(DEFAULT_THRESHOLDS)); }
  }

  function _loadKpis() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        _nextId = parsed.reduce((m,k) => Math.max(m, parseInt(k.id?.replace('kpi_','')||0)), 0) + 1;
        // Migrate: ensure monthlyActuals exists on old records
        parsed.forEach(k => {
          if (!k.monthlyActuals) k.monthlyActuals = {};
          if (!k.ytdMethod) k.ytdMethod = 'sum';
        });
        return parsed;
      }
    } catch {}
    // Start empty — no hardcoded KPIs
    return [];
  }

  // ── Fiscal Year Month Helpers ─────────────────────────────────────────────
  /**
   * Returns ordered array of month abbreviations for the fiscal year,
   * starting from fyStartMonth. e.g. fyStartMonth='Jul' → ['Jul','Aug',...,'Jun']
   */
  function getFyMonths() {
    const start = ALL_MONTHS.indexOf(_settings.fyStartMonth || 'Jul');
    const idx   = start < 0 ? 6 : start;
    return [...ALL_MONTHS.slice(idx), ...ALL_MONTHS.slice(0, idx)];
  }

  function _loadFormulas() {
    try {
      const raw = localStorage.getItem(FORMULA_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        _nextFmlId = parsed.reduce((m,f) => Math.max(m, parseInt(f.id?.replace('fml_','')||0)), 0) + 1;
        return parsed;
      }
    } catch {}
    return [];
  }

  // ── Formula Engine ─────────────────────────────────────────────────────────
  /**
   * A formula KPI has a `formula` object:
   * {
   *   op:       'sum' | 'subtract' | 'multiply' | 'divide' | 'avg' | 'min' | 'max' | 'custom',
   *   operands: [ kpiId, kpiId, ... ],   // for standard ops
   *   expression: string,                // for 'custom' — uses kpi IDs as variables e.g. "kpi_1 + kpi_2 * 0.5"
   * }
   * The computed values replace ytd, actual, monthlyActuals on the formula KPI.
   */

  function _evalFormula(formula, valueMap) {
    const op       = formula.op;
    const operands = (formula.operands || []).map(id => valueMap[id] ?? null);
    const vals     = operands.filter(v => v !== null && !isNaN(v));
    if (op === 'sum')      return vals.length ? vals.reduce((a,b)=>a+b, 0) : null;
    if (op === 'subtract') {
      if (vals.length === 0) return null;
      return vals.slice(1).reduce((a,b)=>a-b, vals[0]);
    }
    if (op === 'multiply') return vals.length ? vals.reduce((a,b)=>a*b, 1) : null;
    if (op === 'divide')   {
      if (vals.length < 2 || vals[1] === 0) return null;
      return vals[0] / vals[1];
    }
    if (op === 'avg')      return vals.length ? vals.reduce((a,b)=>a+b, 0) / vals.length : null;
    if (op === 'min')      return vals.length ? Math.min(...vals) : null;
    if (op === 'max')      return vals.length ? Math.max(...vals) : null;
    if (op === 'custom' && formula.expression) {
      try {
        // Replace kpi IDs with their values in the expression
        let expr = formula.expression;
        Object.entries(valueMap).forEach(([id, val]) => {
          expr = expr.replace(new RegExp(id.replace('_','_'), 'g'), val !== null ? val : 0);
        });
        // Safe eval: only allow numbers and operators
        if (!/^[\d\s+\-*/().]+$/.test(expr)) return null;
        return Function('"use strict"; return (' + expr + ')')();
      } catch { return null; }
    }
    return null;
  }

  function _applyAllFormulas() {
    if (!_formulas.length) return;
    const fyMonths = getFyMonths();

    // Sort formulas so dependencies are resolved in order
    // (formula KPIs that reference other formula KPIs)
    const formulaKpiIds = new Set(_formulas.map(f => f.kpiId));

    _formulas.forEach(fml => {
      const kpi = _kpis.find(k => k.id === fml.kpiId);
      if (!kpi) return;

      // Compute per-month actuals
      const newMonthlyActuals = {};
      fyMonths.forEach(month => {
        const valueMap = {};
        (fml.operands || []).forEach(opId => {
          const opKpi = _kpis.find(k => k.id === opId);
          if (opKpi) valueMap[opId] = opKpi.monthlyActuals?.[month] ?? null;
        });
        const result = _evalFormula(fml, valueMap);
        if (result !== null) newMonthlyActuals[month] = result;
      });
      kpi.monthlyActuals = newMonthlyActuals;

      // Compute ytd from monthly actuals
      const allVals = fyMonths.map(m => newMonthlyActuals[m]).filter(v => v !== null && v !== undefined);
      const method = kpi.ytdMethod || 'sum';
      if (allVals.length === 0) {
        kpi.ytd = null; kpi.actual = null;
      } else {
        if      (method === 'sum')  kpi.ytd = allVals.reduce((a,b)=>a+b, 0);
        else if (method === 'avg')  kpi.ytd = allVals.reduce((a,b)=>a+b, 0) / allVals.length;
        else if (method === 'last') kpi.ytd = allVals[allVals.length - 1];
        else if (method === 'max')  kpi.ytd = Math.max(...allVals);
        else if (method === 'min')  kpi.ytd = Math.min(...allVals);
        else                        kpi.ytd = allVals.reduce((a,b)=>a+b, 0);
        kpi.actual = allVals[allVals.length - 1];
      }

      // Compute targets if operands have targets
      const targetMap = {};
      (fml.operands || []).forEach(opId => {
        const opKpi = _kpis.find(k => k.id === opId);
        if (opKpi) targetMap[opId] = opKpi.targetFY26 ?? null;
      });
      kpi.targetFY26 = _evalFormula(fml, targetMap);

      const tgtMonthMap = {};
      (fml.operands || []).forEach(opId => {
        const opKpi = _kpis.find(k => k.id === opId);
        if (opKpi) tgtMonthMap[opId] = opKpi.targetMonth ?? null;
      });
      kpi.targetMonth = _evalFormula(fml, tgtMonthMap);
    });

    _computeRag();
  }

  // ── Formula CRUD ──────────────────────────────────────────────────────────
  function getFormulas()       { return _formulas; }
  function getFormulaByKpiId(kpiId) { return _formulas.find(f => f.kpiId === kpiId); }

  function addFormulaKpi(kpiFields, formulaDef) {
    // 1. Create the underlying KPI (read-only data, managed by formula)
    const kpiId = addKpi({ ...kpiFields, isFormula: true });
    // 2. Store the formula definition
    const fml = {
      id:         'fml_' + (_nextFmlId++),
      kpiId,
      op:         formulaDef.op       || 'sum',
      operands:   formulaDef.operands || [],
      expression: formulaDef.expression || '',
    };
    _formulas.push(fml);
    _saveFormulas();
    _applyAllFormulas();
    _save(); _notify();
    return kpiId;
  }

  function updateFormulaKpi(kpiId, kpiFields, formulaDef) {
    updateKpi(kpiId, { ...kpiFields, isFormula: true });
    const fml = _formulas.find(f => f.kpiId === kpiId);
    if (fml) {
      Object.assign(fml, {
        op:         formulaDef.op       || fml.op,
        operands:   formulaDef.operands || fml.operands,
        expression: formulaDef.expression !== undefined ? formulaDef.expression : fml.expression,
      });
    }
    _saveFormulas();
    _applyAllFormulas();
    _save(); _notify();
  }

  function removeFormulaKpi(kpiId) {
    _formulas = _formulas.filter(f => f.kpiId !== kpiId);
    removeKpi(kpiId);
    _saveFormulas();
  }

  function recomputeFormulas() {
    _applyAllFormulas();
    _save(); _notify();
  }

  function _saveFormulas() {
    try { localStorage.setItem(FORMULA_KEY, JSON.stringify(_formulas)); } catch(e) {}
  }

  // ── Persistence ───────────────────────────────────────────────────────────
  function _save() {
    try {
      localStorage.setItem(STORAGE_KEY,   JSON.stringify(_kpis));
      localStorage.setItem(SETTINGS_KEY,  JSON.stringify(_settings));
      localStorage.setItem(THRESHOLD_KEY, JSON.stringify(_thresholds));
    } catch(e) { console.warn('Save failed:', e); }
  }

  // ── RAG Computation ───────────────────────────────────────────────────────
  function _computeRag() {
    _kpis.forEach(kpi => {
      if (kpi.ragOverride) { kpi.rag = kpi.ragOverride; return; }
      kpi.rag = _autoRag(kpi);
    });
  }

  function _autoRag(kpi) {
    const th = _thresholds.find(t => t.id === kpi.thresholdId);
    if (!th || th.type === 'manual' || !th.levels?.length) {
      const op     = kpi.targetFY26Op;
      const target = kpi.targetMonth !== null ? kpi.targetMonth : kpi.targetFY26;
      const actual = kpi.actual !== null ? kpi.actual : kpi.ytd;
      if (op && target !== null && actual !== null) {
        return _evalOp(actual, op, target) ? 'green' : 'red';
      }
      return 'neutral';
    }
    const actual = kpi.actual !== null ? kpi.actual : kpi.ytd;
    if (actual === null || actual === undefined) return 'neutral';
    const target = kpi.targetMonth !== null ? kpi.targetMonth : kpi.targetFY26;
    if (target === null || target === undefined) return 'neutral';
    for (const level of th.levels) {
      const cv = (th.type==='relative'||th.type==='relative_lower') ? (target!==0?actual/target:0) : actual;
      if (_evalOp(cv, level.op, level.value)) return level.rag;
    }
    return 'neutral';
  }

  /**
   * Compute RAG for a specific period's actual vs target.
   * Uses the same threshold logic as _autoRag but with caller-supplied actual/target.
   */
  function _ragForPeriod(kpi, actual, target) {
    if (kpi.ragOverride) return kpi.ragOverride;
    if (actual === null || actual === undefined) return 'neutral';
    if (target === null || target === undefined) return 'neutral';

    const th = _thresholds.find(t => t.id === kpi.thresholdId);

    // No threshold or manual only — fall back to operator-target check
    if (!th || th.type === 'manual' || !th.levels?.length) {
      const op = kpi.targetFY26Op;
      if (op) return _evalOp(actual, op, target) ? 'green' : 'red';
      return 'neutral';
    }

    for (const level of th.levels) {
      const cv = (th.type==='relative'||th.type==='relative_lower')
        ? (target !== 0 ? actual / target : 0)
        : actual;
      if (_evalOp(cv, level.op, level.value)) return level.rag;
    }
    return 'neutral';
  }

  function _evalOp(a, op, b) {
    switch(op) {
      case '>':return a>b; case '>=':return a>=b; case '<':return a<b;
      case '<=':return a<=b; case '=':case '==':return Math.abs(a-b)<0.0001;
      default:return false;
    }
  }

  // ── KPI CRUD ──────────────────────────────────────────────────────────────
  function getKpis()           { return _kpis; }
  function getKpiById(id)      { return _kpis.find(k=>k.id===id); }
  function getKpisBySection(s) { return _kpis.filter(k=>k.section===s); }
  function getKeyKpis()        { return _kpis.filter(k=>k.isKey); }
  function getSections()       { return [...new Set(_kpis.map(k=>k.section))]; }

  function addKpi(fields={}) {
    const id = 'kpi_'+(_nextId++);
    const kpi = {
      id,
      section:       fields.section       || 'General',
      metric:        fields.metric        || 'New KPI',
      who:           fields.who           || '',
      unit:          fields.unit          || '',          // '$' | '%' | ''
      targetFY26:    fields.targetFY26    !== undefined ? fields.targetFY26    : null,
      targetFY26Op:  fields.targetFY26Op  || null,        // '>' | '>=' | '<' | '<=' | null
      targetFY26Raw: fields.targetFY26Raw || '',          // legacy raw string (kept for compat)
      targetLabel:   fields.targetLabel   || '',          // human display label e.g. "> 80%"
      targetMonth:   fields.targetMonth   !== undefined ? fields.targetMonth   : null,
      ytd:           fields.ytd           !== undefined ? fields.ytd           : null,
      actual:        null,
      rag:           'neutral',
      ragOverride:   null,
      thresholdId:   fields.thresholdId   || 'th_manual_only',
      comment:       fields.comment       || '',
      isKey:         fields.isKey         || false,
      monthlyActuals: fields.monthlyActuals || {},
      ytdMethod:     fields.ytdMethod      || 'sum',  // 'sum' | 'avg' | 'last' | 'max' | 'min'
    };
    _kpis.push(kpi);
    _computeRag(); _save(); _notify();
    return id;
  }

  function updateKpi(id, fields) {
    const kpi = _kpis.find(k=>k.id===id);
    if (!kpi) return;
    Object.assign(kpi, fields);
    if (kpi.targetMonth!==null && fields.targetFY26===undefined && !kpi.targetFY26Op && kpi.targetMonth)
      kpi.targetFY26 = kpi.targetMonth * 12;
    _applyAllFormulas();
    _computeRag(); _save(); _notify();
  }

  /** Update a single month's actual value for a KPI */
  function updateMonthlyActual(id, month, value) {
    const kpi = _kpis.find(k=>k.id===id);
    if (!kpi) return;
    if (!kpi.monthlyActuals) kpi.monthlyActuals = {};
    if (value === null || value === '' || isNaN(parseFloat(value))) {
      delete kpi.monthlyActuals[month];
    } else {
      kpi.monthlyActuals[month] = parseFloat(value);
    }
    // Derive ytd from monthly actuals using the KPI's ytdMethod
    const fyMonths = getFyMonths();
    const enteredMonths = fyMonths.filter(m => kpi.monthlyActuals[m] !== undefined);
    const vals = enteredMonths.map(m => kpi.monthlyActuals[m]);
    if (vals.length === 0) {
      kpi.ytd = null;
    } else {
      const method = kpi.ytdMethod || 'sum';
      if      (method === 'sum')  kpi.ytd = vals.reduce((a,b)=>a+b, 0);
      else if (method === 'avg')  kpi.ytd = vals.reduce((a,b)=>a+b, 0) / vals.length;
      else if (method === 'last') kpi.ytd = vals[vals.length - 1];
      else if (method === 'max')  kpi.ytd = Math.max(...vals);
      else if (method === 'min')  kpi.ytd = Math.min(...vals);
      else                        kpi.ytd = vals.reduce((a,b)=>a+b, 0);
    }
    // Derive actual as most recent month entered (last in FY order)
    kpi.actual = enteredMonths.length > 0 ? kpi.monthlyActuals[enteredMonths[enteredMonths.length-1]] : null;
    _applyAllFormulas();
    _computeRag(); _save(); _notify();
  }

  function removeKpi(id) { _kpis=_kpis.filter(k=>k.id!==id); _save(); _notify(); }

  function reorderKpi(id, direction) {
    const idx=_kpis.findIndex(k=>k.id===id); if(idx<0)return;
    const newIdx=direction==='up'?idx-1:idx+1;
    if(newIdx<0||newIdx>=_kpis.length)return;
    [_kpis[idx],_kpis[newIdx]]=[_kpis[newIdx],_kpis[idx]];
    _save(); _notify();
  }

  /**
   * Re-assign all KPIs from index `fromIndex` onwards (within the same section)
   * to a new section name. Used for "insert section break" feature.
   */
  function splitSectionAt(kpiId, newSectionName) {
    if (!newSectionName) return;
    const idx = _kpis.findIndex(k=>k.id===kpiId);
    if (idx < 0) return;
    const oldSection = _kpis[idx].section;
    // All KPIs from this index onwards that belong to the same section get the new name
    for (let i = idx; i < _kpis.length; i++) {
      if (_kpis[i].section === oldSection) _kpis[i].section = newSectionName;
    }
    _save(); _notify();
  }

  // ── Section CRUD ──────────────────────────────────────────────────────────
  function addSection(name) { if(!name||getSections().includes(name))return; addKpi({section:name,metric:'New KPI'}); }
  function renameSection(oldName,newName) {
    if(!newName||oldName===newName)return;
    _kpis.forEach(k=>{if(k.section===oldName)k.section=newName;});
    _save(); _notify();
  }
  function removeSection(name) { _kpis=_kpis.filter(k=>k.section!==name); _save(); _notify(); }

  // ── Thresholds CRUD ───────────────────────────────────────────────────────
  function getThresholds()      { return _thresholds; }
  function getThresholdById(id) { return _thresholds.find(t=>t.id===id); }

  function addThreshold(fields={}) {
    const id='th_'+Date.now();
    const th={ id, name:fields.name||'New Threshold', description:fields.description||'',
      type:fields.type||'relative',
      levels:fields.levels||[{rag:'green',op:'>=',value:0.90,label:'≥ 90% of target'},{rag:'amber',op:'>=',value:0.70,label:'≥ 70% of target'},{rag:'red',op:'<',value:0.70,label:'< 70% of target'}] };
    _thresholds.push(th); _save(); _notify(); return id;
  }

  function updateThreshold(id, fields) {
    const th=_thresholds.find(t=>t.id===id); if(!th)return;
    Object.assign(th,fields); _computeRag(); _save(); _notify();
  }

  function removeThreshold(id) {
    if(id==='th_manual_only')return;
    _kpis.forEach(k=>{if(k.thresholdId===id)k.thresholdId='th_manual_only';});
    _thresholds=_thresholds.filter(t=>t.id!==id); _computeRag(); _save(); _notify();
  }

  // ── Settings ──────────────────────────────────────────────────────────────
  function getSettings()     { return _settings; }
  function updateSettings(f) { Object.assign(_settings,f); _save(); _notify(); }

  // ── Import ────────────────────────────────────────────────────────────────
  function importFromRows(rows) {
    let updated = 0, added = 0;
    rows.forEach(row => {
      if (!row.metric) return;
      const existing = _kpis.find(k => k.metric.toLowerCase().trim() === row.metric.toLowerCase().trim());
      if (existing) {
        const u = {};
        if (row.who        !== undefined)                    u.who          = row.who;
        if (row.unit)                                        u.unit         = row.unit;
        if (row.targetLabel)                                 u.targetLabel  = row.targetLabel;
        if (row.targetFY26Op)                                u.targetFY26Op = row.targetFY26Op;
        if (!isNaN(parseFloat(row.targetFY26)))              u.targetFY26   = parseFloat(row.targetFY26);
        if (!isNaN(parseFloat(row.ytd)))                     u.ytd          = parseFloat(row.ytd);
        if (!isNaN(parseFloat(row.targetMonth)))             u.targetMonth  = parseFloat(row.targetMonth);
        if (!isNaN(parseFloat(row.actual)))                  u.actual       = parseFloat(row.actual);
        if (row.rag)                                         u.ragOverride  = row.rag.toLowerCase().trim();
        if (row.comment)                                     u.comment      = row.comment;
        updateKpi(existing.id, u); updated++;
      } else {
        addKpi({
          section:      row.section      || 'Imported',
          metric:       row.metric,
          who:          row.who          || '',
          unit:         row.unit         || '',
          targetLabel:  row.targetLabel  || '',
          targetFY26Op: row.targetFY26Op || null,
          targetFY26:   isNaN(parseFloat(row.targetFY26))  ? null : parseFloat(row.targetFY26),
          targetMonth:  isNaN(parseFloat(row.targetMonth)) ? null : parseFloat(row.targetMonth),
          ytd:          isNaN(parseFloat(row.ytd))         ? null : parseFloat(row.ytd),
        });
        added++;
      }
    });
    return { updated, added };
  }

  // ── Formatting ────────────────────────────────────────────────────────────

  /**
   * Format a numeric value for display, using the KPI's unit field.
   * Respects settings: decimals, largeNumFormat, pctStorage.
   *
   * unit: '$' → currency formatting
   *       '%' → percentage
   *             pctStorage='decimal': value stored as 0.85 → displays as "85%"
   *             pctStorage='direct':  value stored as 85   → displays as "85%"
   *       ''  → plain number with locale separators
   *
   * largeNumFormat: 'auto' | 'M' | 'K' | 'full'
   *   auto: ≥1B→B, ≥1M→M, ≥1K→K
   *   M:    ≥1M→M only
   *   K:    ≥1K→K only
   *   full: always full number with commas
   */
  function formatValue(val, kpi) {
    if (val === null || val === undefined) return '—';
    const n = parseFloat(val);
    if (isNaN(n)) return String(val);

    const unit       = kpi?.unit || _inferUnit(kpi);
    const dec        = parseInt(_settings.decimals ?? 2);
    const largeFmt   = _settings.largeNumFormat || 'auto';
    const pctStorage = _settings.pctStorage || 'decimal'; // 'decimal' | 'direct'

    if (unit === '%') {
      const pctVal = pctStorage === 'decimal' ? n * 100 : n;
      return pctVal.toFixed(dec) + '%';
    }
    if (unit === '$') {
      const sym = _settings.currencySymbol || '$';
      return sym + _fmtNum(n, dec, largeFmt);
    }
    // Plain number
    return _fmtNum(n, dec, largeFmt);
  }

  /**
   * Core number formatter respecting large-number format and decimal setting.
   */
  function _fmtNum(n, dec, largeFmt) {
    const abs  = Math.abs(n);
    const sign = n < 0 ? '-' : '';
    if (largeFmt === 'auto') {
      if (abs >= 1e9) return sign + (abs / 1e9).toFixed(dec) + 'B';
      if (abs >= 1e6) return sign + (abs / 1e6).toFixed(dec) + 'M';
      if (abs >= 1e3) return sign + (abs / 1e3).toFixed(dec) + 'K';
    } else if (largeFmt === 'M' && abs >= 1e6) {
      return sign + (abs / 1e6).toFixed(dec) + 'M';
    } else if (largeFmt === 'K' && abs >= 1e3) {
      return sign + (abs / 1e3).toFixed(dec) + 'K';
    }
    // full or fallback
    return abs % 1 === 0 && dec === 0
      ? sign + abs.toLocaleString()
      : sign + parseFloat(abs.toFixed(dec)).toLocaleString(undefined, {
          minimumFractionDigits: dec,
          maximumFractionDigits: dec,
        });
  }

  /** Fallback: infer unit from legacy targetFY26Raw or value range */
  function _inferUnit(kpi) {
    const raw = kpi?.targetFY26Raw || kpi?.targetLabel || '';
    if (raw.includes('%')) return '%';
    if (raw.includes('$')) return '$';
    // Heuristic: if target is a decimal ≤ 1, likely a percentage stored as decimal
    if (kpi?.targetFY26 !== null && kpi?.targetFY26 > 0 && kpi?.targetFY26 <= 1) return '%';
    return '';
  }

  /**
   * Format the target for display on cards.
   * Priority: targetLabel (e.g. "> 80%") → formatted targetFY26 value
   */
  function formatTarget(kpi) {
    if (kpi.targetLabel) return kpi.targetLabel;
    if (kpi.targetFY26Raw) return kpi.targetFY26Raw;  // legacy compat
    const target = kpi.targetMonth !== null ? kpi.targetMonth : kpi.targetFY26;
    if (target === null) return '—';
    const base = formatValue(target, kpi);
    // Prepend operator if present (e.g. "> 80%")
    return kpi.targetFY26Op ? kpi.targetFY26Op + ' ' + base : base;
  }


  // ── Period Math ───────────────────────────────────────────────────────────

  /**
   * Returns the current calendar month abbreviation (e.g. "Mar").
   */
  function getCurrentMonth() {
    return ALL_MONTHS[new Date().getMonth()];
  }

  /**
   * Returns how many FY months have elapsed (including current month).
   * e.g. FY starts Jul, current month is Sep → 3 months elapsed.
   */
  function getFyMonthsElapsed() {
    const fyMonths = getFyMonths();
    const curMonth = getCurrentMonth();
    const idx = fyMonths.indexOf(curMonth);
    return idx < 0 ? fyMonths.length : idx + 1;
  }

  /**
   * Returns the FY months elapsed up to and including the current month.
   */
  function getFyMonthsToDate() {
    const fyMonths  = getFyMonths();
    const elapsed   = getFyMonthsElapsed();
    return fyMonths.slice(0, elapsed);
  }

  /**
   * Returns the current quarter number within the FY (1-indexed).
   * e.g. FY starts Jul: Jul-Sep=Q1, Oct-Dec=Q2, Jan-Mar=Q3, Apr-Jun=Q4
   */
  function getCurrentFyQuarter() {
    const fyMonths  = getFyMonths();
    const curMonth  = getCurrentMonth();
    const idx       = fyMonths.indexOf(curMonth);
    return idx < 0 ? 1 : Math.floor(idx / 3) + 1;
  }

  /**
   * Returns the three months for a given FY quarter number (1-4).
   */
  function getFyQuarterMonths(quarterNum) {
    const fyMonths = getFyMonths();
    const start    = (quarterNum - 1) * 3;
    return fyMonths.slice(start, start + 3);
  }

  /**
   * Compute display stats for a KPI based on the selected overview period mode.
   *
   * mode: 'monthly'   → current month actual vs targetMonth
   *       'quarterly' → sum of current quarter months vs targetMonth * 3
   *       'ytd'       → sum of FY months to date vs targetMonth * months_elapsed
   *       'yearly'    → targetFY26 and ytd (full year view)
   *       'last_fy'   → placeholder (no prior-year data yet)
   *
   * Returns: { actual, target, progressPct, label, periodLabel, isPlaceholder }
   *   actual      - numeric value for the period (or null)
   *   target      - numeric target for the period (or null)
   *   progressPct - 0-100 integer
   *   label       - short description e.g. "Q2 FY26"
   *   periodLabel - what the value represents e.g. "Monthly Actual"
   */
  function getPeriodStats(kpi, mode) {
    const ma          = kpi.monthlyActuals || {};
    const fyMonths    = getFyMonths();
    const curMonth    = getCurrentMonth();
    const tgtMonth    = kpi.targetMonth !== null && kpi.targetMonth !== undefined ? parseFloat(kpi.targetMonth) : null;
    const tgtFY       = kpi.targetFY26   !== null && kpi.targetFY26   !== undefined ? parseFloat(kpi.targetFY26)  : null;
    const fyLabel     = _settings.fiscalYearLabel || 'FY';

    // Helper: sum actuals for a set of months (skipping missing)
    const getVals = (months) =>
      months.map(m => ma[m]).filter(v => v !== null && v !== undefined && !isNaN(parseFloat(v))).map(parseFloat);

    const sumMonths = (months) => { const v = getVals(months); return v.length > 0 ? v.reduce((a,b)=>a+b,0) : null; };

    // Aggregate months according to the KPI's ytdMethod
    const method = kpi.ytdMethod || 'sum';
    const aggregateMonths = (months) => {
      const vals = getVals(months);
      if (vals.length === 0) return null;
      switch (method) {
        case 'avg':  return vals.reduce((a,b)=>a+b,0) / vals.length;
        case 'last': return vals[vals.length - 1];
        case 'max':  return Math.max(...vals);
        case 'min':  return Math.min(...vals);
        default:     return vals.reduce((a,b)=>a+b,0); // sum
      }
    };

    // Whether the method produces a representative single-value (not additive)
    const isPointInTime = method === 'avg' || method === 'last' || method === 'max' || method === 'min';

    let actual = null, target = null, label = '', periodLabel = '';

    if (mode === 'monthly') {
      // Monthly: always the current month's actual vs monthly target
      actual      = ma[curMonth] !== undefined ? parseFloat(ma[curMonth]) : (kpi.actual !== null ? kpi.actual : null);
      target      = tgtMonth;
      label       = curMonth + ' ' + fyLabel;
      periodLabel = 'Month Actual';

    } else if (mode === 'quarterly') {
      const q        = getCurrentFyQuarter();
      const qMonths  = getFyQuarterMonths(q);
      actual         = aggregateMonths(qMonths);
      // Point-in-time methods compare against monthly target; additive (sum) scales up
      target         = isPointInTime
        ? tgtMonth
        : (tgtMonth !== null ? tgtMonth * 3 : (tgtFY !== null ? tgtFY / 4 : null));
      label          = 'Q' + q + ' ' + fyLabel;
      periodLabel    = method === 'avg' ? 'Quarterly Avg' : method === 'last' ? 'Latest Month' : method === 'max' ? 'Quarter High' : method === 'min' ? 'Quarter Low' : 'Quarterly Total';

    } else if (mode === 'ytd') {
      const toDate   = getFyMonthsToDate();
      actual         = aggregateMonths(toDate);
      if (actual === null) actual = kpi.ytd !== null ? kpi.ytd : null;
      const elapsed  = toDate.length;
      // Point-in-time methods compare against monthly target; sum scales by elapsed months
      target         = isPointInTime
        ? tgtMonth
        : (tgtMonth !== null ? tgtMonth * elapsed : (tgtFY !== null ? (tgtFY / fyMonths.length) * elapsed : null));
      label          = 'YTD ' + fyLabel + ' (' + elapsed + ' mo)';
      periodLabel    = method === 'avg' ? 'YTD Average' : method === 'last' ? 'Latest Month' : method === 'max' ? 'YTD High' : method === 'min' ? 'YTD Low' : 'YTD Total';

    } else if (mode === 'yearly') {
      // Full FY: aggregate all FY months, or fall back to stored ytd
      actual         = aggregateMonths(fyMonths);
      if (actual === null) actual = kpi.ytd !== null ? kpi.ytd : null;
      target         = isPointInTime
        ? tgtMonth
        : (tgtFY !== null ? tgtFY : (tgtMonth !== null ? tgtMonth * fyMonths.length : null));
      label          = fyLabel + ' Full Year';
      periodLabel    = method === 'avg' ? 'FY Average' : method === 'last' ? 'Latest Month' : method === 'max' ? 'FY High' : method === 'min' ? 'FY Low' : 'Full Year Total';

    } else if (mode === 'last_fy') {
      return { actual: null, target: null, progressPct: 0, label: 'Last ' + fyLabel, periodLabel: 'Prior Year', isPlaceholder: true };
    }

    // Progress calculation
    let pct = 0;
    if (actual !== null && target !== null && target !== 0) {
      const th = _thresholds.find(t => t.id === kpi.thresholdId);
      if (th?.type === 'relative_lower') {
        pct = Math.min(100, Math.round((target / Math.max(actual, 0.0001)) * 100));
      } else {
        pct = Math.min(100, Math.round((actual / target) * 100));
      }
    }

    return { actual, target, progressPct: pct, rag: _ragForPeriod(kpi, actual, target), label, periodLabel, isPlaceholder: false };
  }

  function progressPct(kpi) {
    const actual = kpi.actual !== null ? kpi.actual : kpi.ytd;
    const target = kpi.targetMonth !== null ? kpi.targetMonth : kpi.targetFY26;
    if (actual === null || !target) return 0;
    // For operator targets (>80%), show actual vs target as-is
    if (kpi.targetFY26Op) return Math.min(100, Math.round((actual / target) * 100));
    const th = _thresholds.find(t => t.id === kpi.thresholdId);
    if (th?.type === 'relative_lower') return Math.min(100, Math.round((target / Math.max(actual, 0.0001)) * 100));
    return Math.min(100, Math.round((actual / target) * 100));
  }

  function subscribe(fn) { _subscribers.push(fn); }

  /** Get KPIs to show on the Overview page (filtered by overviewKpiIds if set). */
  function getOverviewKpis() {
    const ids = _settings.overviewKpiIds;
    const keyKpis = _kpis.filter(k => k.isKey);
    if (!ids) return keyKpis;
    return keyKpis.filter(k => ids.includes(k.id));
  }

  /** Persist the set of KPI ids selected for the overview panel. null = all key KPIs. */
  function setOverviewKpiIds(ids) {
    _settings.overviewKpiIds = ids;
    _save(); _notify();
  }
  function _notify() { _subscribers.forEach(fn=>fn()); }

  function resetToDefaults() {
    localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(SETTINGS_KEY); localStorage.removeItem(THRESHOLD_KEY);
    _thresholds=JSON.parse(JSON.stringify(DEFAULT_THRESHOLDS));
    _settings={ companyName:'My Company', fiscalYearLabel:'FY26', reportingPeriod:'monthly', fyStartMonth:'Jul', decimals:2, largeNumFormat:'auto', currencySymbol:'$', pctStorage:'decimal', overviewKpiIds:null };
    _kpis=[]; _nextId=1;
    _save(); _notify();
  }

  return {
    init, getKpis, getKpiById, getKpisBySection, getKeyKpis, getSections,
    addKpi, updateKpi, updateMonthlyActual, removeKpi, reorderKpi, splitSectionAt,
    addSection, renameSection, removeSection,
    getThresholds, getThresholdById, addThreshold, updateThreshold, removeThreshold,
    getFormulas, getFormulaByKpiId, addFormulaKpi, updateFormulaKpi, removeFormulaKpi, recomputeFormulas,
    getSettings, updateSettings, getFyMonths,
    getCurrentMonth, getFyMonthsElapsed, getFyMonthsToDate, getCurrentFyQuarter,
    getFyQuarterMonths, getPeriodStats,
    importFromRows, formatValue, formatTarget, progressPct, resetToDefaults, subscribe, getOverviewKpis, setOverviewKpiIds,
  };
})();
/**
 * components.js v3
 * Reusable UI components.
 * Sidebar now lists sections as individual nav items (full pages).
 */

const Components = (() => {

  const RAG_ICONS  = { green:'●', amber:'▲', red:'✕', neutral:'○' };
  const RAG_LABELS = { green:'On Track', amber:'At Risk', red:'Off Track', neutral:'No Data' };

  function ragBadge(rag) {
    const r = rag||'neutral';
    return `<span class="rag-badge ${r}">${RAG_ICONS[r]} ${RAG_LABELS[r]}</span>`;
  }

  function ragDot(rag) {
    const r = rag||'neutral';
    return `<span class="rag-badge ${r}" style="padding:2px 7px;font-size:10px">${RAG_ICONS[r]}</span>`;
  }

  // ── KPI Card ──────────────────────────────────────────────────────────────
  function kpiCard(kpi, periodStats, showKeyToggle) {
    let val, targetDisp, pct, periodLabel, cardLabel, rag;

    if (periodStats && !periodStats.isPlaceholder) {
      val         = DataStore.formatValue(periodStats.actual, kpi);
      targetDisp  = periodStats.target !== null ? DataStore.formatValue(periodStats.target, kpi) : '—';
      pct         = periodStats.progressPct;
      periodLabel = periodStats.periodLabel;
      cardLabel   = periodStats.label;
      rag         = periodStats.rag || kpi.rag || 'neutral';  // period-aware RAG
    } else if (periodStats && periodStats.isPlaceholder) {
      val         = '—';
      targetDisp  = '—';
      pct         = 0;
      periodLabel = periodStats.periodLabel;
      cardLabel   = periodStats.label;
      rag         = kpi.rag || 'neutral';
    } else {
      const actual = kpi.actual!==null?kpi.actual:kpi.ytd;
      val         = DataStore.formatValue(actual, kpi);
      targetDisp  = DataStore.formatTarget(kpi);
      pct         = DataStore.progressPct(kpi);
      periodLabel = 'Actual';
      cardLabel   = '';
      rag         = kpi.rag || 'neutral';
    }

    const barColor = {green:'var(--rag-green)',amber:'var(--rag-amber)',red:'var(--rag-red)',neutral:'var(--rag-neutral)'}[rag];

    const keyToggle = showKeyToggle ? `
      <button onclick="event.stopPropagation();App.toggleKpiKey('${kpi.id}')"
              title="${kpi.isKey?'Remove from Key Metrics':'Add to Key Metrics'}"
              style="background:none;border:none;cursor:pointer;font-size:15px;padding:2px 4px;line-height:1;
                     color:${kpi.isKey?'var(--brand-accent)':'var(--text-muted)'};transition:color 0.15s;flex-shrink:0">
        ${kpi.isKey?'<i class="fa-solid fa-star"></i>':'<i class="fa-regular fa-star"></i>'}
      </button>` : '';

    return `
      <div class="kpi-card rag-${rag}" data-kpi-id="${kpi.id}" onclick="App.openKpiDetail('${kpi.id}')">
        <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:6px">
          <div class="label-xs">${kpi.section}</div>
          <div style="display:flex;align-items:center;gap:4px">
            ${kpi.who?`<div class="label-xs" style="color:var(--brand-accent)">${kpi.who}</div>`:''}
            ${keyToggle}
          </div>
        </div>
        <div style="font-size:13px;font-weight:500;color:var(--text-primary);margin-bottom:10px;line-height:1.3">${kpi.metric}</div>
        <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:4px">
          <div class="value-xl">${val}</div>
          ${cardLabel?`<div class="label-xs" style="color:var(--text-muted)">${cardLabel}</div>`:''}
        </div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:6px;flex-wrap:wrap">
          <span class="label-xs">Target: ${targetDisp}</span>
          ${ragBadge(rag)}
        </div>
        <div class="progress-bar">
          <div class="progress-bar-fill" style="width:${pct}%;background:${barColor}"></div>
        </div>
        ${pct>0?`<div class="label-xs" style="margin-top:4px;text-align:right">${pct}% of target</div>`:''}
        ${kpi.comment?`<div style="margin-top:8px;font-size:11px;color:var(--rag-amber);font-style:italic">${kpi.comment}</div>`:''}
      </div>`;
  }

  // ── Sidebar ───────────────────────────────────────────────────────────────
  function sidebar(activePage) {
    const settings = DataStore.getSettings();
    const sections = DataStore.getSections();

    // Static nav items (top-level)
    const staticItems = [
      { id:'overview',    label:'Overview',   icon:'◈' },
      { id:'data-entry',  label:'Data Entry', icon:'✎' },
      { id:'formulas',    label:'Formulas',   icon:'∑' },
      { id:'settings',    label:'Settings',   icon:'⚙' },
    ];

    const navBtn = (id, label, icon, active) => `
      <button class="nav-item ${active?'active':''}"
              onclick="App.navigate('${id}')"
              style="width:100%;display:flex;align-items:center;gap:10px;padding:9px 12px;
                     border-radius:8px;font-size:13px;font-weight:500;text-align:left;
                     color:${active?'var(--brand-accent)':'var(--text-secondary)'};
                     background:${active?'rgba(0,194,168,0.10)':'transparent'};
                     border:none;cursor:pointer;transition:all 0.15s;margin-bottom:2px">
        <span style="font-size:15px;width:18px;text-align:center">${icon}</span>${label}
      </button>`;

    const sectionBtn = (section, active) => {
      const pageId = 'section:' + encodeURIComponent(section);
      const sKpis  = DataStore.getKpisBySection(section);
      const counts = {green:0,amber:0,red:0,neutral:0};
      sKpis.forEach(k=>{counts[k.rag||'neutral']++;});
      const redCount   = counts.red;
      const amberCount = counts.amber;
      const indicator  = redCount>0 ? `<span style="margin-left:auto;font-size:9px;background:var(--rag-red-bg);color:var(--rag-red);padding:1px 5px;border-radius:3px">${redCount}✕</span>`
                        : amberCount>0 ? `<span style="margin-left:auto;font-size:9px;background:var(--rag-amber-bg);color:var(--rag-amber);padding:1px 5px;border-radius:3px">${amberCount}▲</span>` : '';
      const short = section.length>24?section.slice(0,22)+'…':section;
      return `
        <button onclick="App.navigate('${pageId}')"
                style="width:100%;display:flex;align-items:center;padding:7px 12px;border-radius:6px;
                       font-size:12px;background:${active?'rgba(0,194,168,0.08)':'none'};
                       color:${active?'var(--brand-accent)':'var(--text-muted)'};border:none;cursor:pointer;
                       text-align:left;transition:all 0.15s;gap:6px"
                onmouseover="if(!${active})this.style.color='var(--text-primary)'"
                onmouseout="if(!${active})this.style.color='var(--text-muted)'">
          <span style="font-size:10px;width:10px;text-align:center;flex-shrink:0">▸</span>
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${short}</span>
          ${indicator}
        </button>`;
    };

    return `
      <div class="sidebar" id="sidebar">
        <div style="padding:20px 16px;border-bottom:1px solid var(--border-subtle)">
          <div style="font-family:var(--font-display);font-size:12px;font-weight:700;color:var(--brand-accent);letter-spacing:0.12em;text-transform:uppercase">Executive</div>
          <div style="font-family:var(--font-display);font-size:18px;font-weight:700;color:var(--text-primary);margin-top:2px;line-height:1.2">${settings.companyName}</div>
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">${settings.fiscalYearLabel} Dashboard</div>
        </div>

        <nav style="flex:1;padding:10px 8px;overflow-y:auto">
          ${staticItems.map(item => navBtn(item.id, item.label, item.icon, activePage===item.id)).join('')}

          <div class="label-xs" style="padding:12px 12px 4px;margin-top:4px">Sections</div>
          ${sections.map(s => sectionBtn(s, activePage==='section:'+encodeURIComponent(s))).join('')}

          <div style="margin-top:10px;padding:0 8px">
            <button onclick="App.openAddSectionModal()"
                    style="width:100%;display:flex;align-items:center;gap:8px;padding:7px 12px;border-radius:6px;
                           font-size:12px;color:var(--brand-accent);background:rgba(0,194,168,0.06);
                           border:1px dashed rgba(0,194,168,0.3);cursor:pointer;transition:all 0.15s">
              <span>+</span> Add Section
            </button>
          </div>
        </nav>

        <div style="padding:12px 16px;border-top:1px solid var(--border-subtle)">
          <div style="font-size:10px;color:var(--text-muted)">Executive Dashboard v3.0</div>
        </div>
      </div>`;
  }

  // ── Top Bar ───────────────────────────────────────────────────────────────
  function topBar(title, subtitle='') {
    const now = new Date().toLocaleDateString('en-AU',{weekday:'short',day:'numeric',month:'short',year:'numeric'});
    return `
      <div class="top-bar">
        <button class="hamburger" onclick="App.toggleSidebar()"><span></span><span></span><span></span></button>
        <div style="flex:1">
          <div class="page-title" style="font-size:16px">${title}</div>
          ${subtitle?`<div class="label-xs">${subtitle}</div>`:''}
        </div>
        <span class="label-xs" style="color:var(--text-muted)">${now}</span>
        <button class="btn btn-ghost" style="font-size:12px;padding:6px 12px" onclick="App.navigate('data-entry')">✎ Update Data</button>
      </div>`;
  }

  // ── RAG Summary Bar ───────────────────────────────────────────────────────
  function ragSummaryBar(kpis) {
    const counts = {green:0,amber:0,red:0,neutral:0};
    kpis.forEach(k=>{counts[k.rag||'neutral']++;});
    return `
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px">
        ${Object.entries(counts).map(([rag,cnt])=>cnt>0?`
          <div style="background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;gap:8px">
            <div class="rag-badge ${rag}">${rag.charAt(0).toUpperCase()+rag.slice(1)}</div>
            <span style="font-family:var(--font-display);font-size:20px;font-weight:700">${cnt}</span>
          </div>`:'').join('')}
        <div style="flex:1;background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;justify-content:flex-end;gap:6px">
          <span class="label-xs">Total KPIs:</span>
          <span style="font-size:14px;font-weight:700;color:var(--brand-accent)">${kpis.length}</span>
        </div>
      </div>`;
  }

  // ── KPI Detail Modal ──────────────────────────────────────────────────────
  function kpiDetailModal(kpi, periodStats) {
    let displayVal, displayTarget, displayPct, periodLabel, cardLabel, rag;
    if (periodStats && !periodStats.isPlaceholder) {
      displayVal    = DataStore.formatValue(periodStats.actual, kpi);
      displayTarget = periodStats.target !== null ? DataStore.formatValue(periodStats.target, kpi) : '—';
      displayPct    = periodStats.progressPct;
      periodLabel   = periodStats.periodLabel;
      cardLabel     = periodStats.label;
      rag           = periodStats.rag || kpi.rag || 'neutral';
    } else {
      const actual  = kpi.actual !== null ? kpi.actual : kpi.ytd;
      displayVal    = DataStore.formatValue(actual, kpi);
      displayTarget = DataStore.formatTarget(kpi);
      displayPct    = DataStore.progressPct(kpi);
      periodLabel   = 'Current / YTD';
      cardLabel     = '';
      rag           = kpi.rag || 'neutral';
    }
    const barColor = {green:'var(--rag-green)',amber:'var(--rag-amber)',red:'var(--rag-red)',neutral:'var(--rag-neutral)'}[rag];
    const th       = DataStore.getThresholdById(kpi.thresholdId);

    const storedMonthTarget = (kpi.targetMonth !== null && kpi.targetMonth !== undefined)
      ? DataStore.formatValue(kpi.targetMonth, kpi) : null;
    const storedFYTarget = (kpi.targetFY26 !== null && kpi.targetFY26 !== undefined)
      ? DataStore.formatValue(kpi.targetFY26, kpi) : null;

    const storedTargetRow = (storedMonthTarget || storedFYTarget || kpi.ytd !== null) ? `
      <div style="background:var(--bg-input);border-radius:8px;padding:10px 14px;margin-bottom:16px;
                  display:flex;gap:20px;flex-wrap:wrap;border:1px solid var(--border-subtle)">
        <div style="font-size:10px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;
                    color:var(--text-muted);width:100%;margin-bottom:4px">Stored Targets</div>
        ${storedMonthTarget ? `<div>
          <div style="font-size:10px;color:var(--text-muted);margin-bottom:2px;text-transform:uppercase;letter-spacing:0.06em">Target / Month</div>
          <div style="font-size:14px;font-weight:600;color:var(--text-secondary)">${storedMonthTarget}</div>
        </div>` : ''}
        ${storedFYTarget ? `<div>
          <div style="font-size:10px;color:var(--text-muted);margin-bottom:2px;text-transform:uppercase;letter-spacing:0.06em">Target / Year</div>
          <div style="font-size:14px;font-weight:600;color:var(--text-secondary)">${storedFYTarget}</div>
        </div>` : ''}
        ${kpi.ytd !== null && kpi.ytd !== undefined ? `<div>
          <div style="font-size:10px;color:var(--text-muted);margin-bottom:2px;text-transform:uppercase;letter-spacing:0.06em">YTD Actual</div>
          <div style="font-size:14px;font-weight:600;color:var(--text-secondary)">${DataStore.formatValue(kpi.ytd, kpi)}</div>
        </div>` : ''}
      </div>` : '';

    return `
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:520px">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px">
            <div>
              <div class="label-xs" style="margin-bottom:4px">${kpi.section}</div>
              <h3 style="font-family:var(--font-display);font-size:17px;font-weight:700;line-height:1.3">${kpi.metric}</h3>
              ${kpi.who?`<div style="font-size:12px;color:var(--brand-accent);margin-top:4px">Owner: ${kpi.who}</div>`:''}
            </div>
            <button onclick="App.closeModal()" style="color:var(--text-muted);font-size:20px;padding:4px;flex-shrink:0">✕</button>
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div class="card" style="padding:14px;text-align:center">
              <div class="label-xs" style="margin-bottom:4px">${periodLabel}</div>
              <div class="value-lg">${displayVal}</div>
              ${cardLabel?`<div style="font-size:10px;color:var(--text-muted);margin-top:3px">${cardLabel}</div>`:''}
            </div>
            <div class="card" style="padding:14px;text-align:center">
              <div class="label-xs" style="margin-bottom:4px">Period Target</div>
              <div class="value-lg">${displayTarget}</div>
              ${cardLabel?`<div style="font-size:10px;color:var(--text-muted);margin-top:3px">${cardLabel}</div>`:''}
            </div>
          </div>

          ${storedTargetRow}
          <div style="margin-bottom:16px">${ragBadge(rag)}</div>

          ${displayPct>0?`
            <div class="progress-bar" style="height:8px;margin-bottom:6px">
              <div class="progress-bar-fill" style="width:${displayPct}%;background:${barColor}"></div>
            </div>
            <div class="label-xs" style="margin-bottom:16px">${displayPct}% of period target achieved</div>`:
          '<div class="label-xs" style="margin-bottom:16px;color:var(--text-muted)">No progress data</div>'}

          ${th?`<div style="margin-bottom:14px">
            <div class="label-xs" style="margin-bottom:4px">RAG Rule</div>
            <div style="font-size:12px;color:var(--text-secondary)">${th.name}</div>
          </div>`:''}

          ${kpi.comment?`<div style="background:var(--bg-input);border-radius:8px;padding:12px;margin-bottom:16px">
            <div class="label-xs" style="margin-bottom:4px">Comment</div>
            <div style="font-size:13px;color:var(--rag-amber)">${kpi.comment}</div>
          </div>`:''}

          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" style="flex:1" onclick="App.openKpiInDataEntry('${kpi.id}');App.closeModal()">⊞ Monthly Data Entry</button>
            <button class="btn btn-ghost" onclick="App.openEditKpi('${kpi.id}');App.closeModal()">✎ Edit KPI</button>
            <button class="btn btn-ghost" onclick="App.closeModal()">Close</button>
          </div>
        </div>
      </div>`;
  }

  // ── Edit/Add KPI Modal ────────────────────────────────────────────────────
  function editKpiModal(kpi, isNew=false) {
    const sections   = DataStore.getSections();
    const thresholds = DataStore.getThresholds();

    return `
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:540px">
          <div style="display:flex;justify-content:space-between;margin-bottom:20px">
            <h3 style="font-family:var(--font-display);font-size:16px;font-weight:700">${isNew?'Add New KPI':'Edit KPI'}</h3>
            <button onclick="App.closeModal()" style="color:var(--text-muted);font-size:20px;padding:4px">✕</button>
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">KPI Name *</label>
              <input type="text" id="e-metric" class="input-field" value="${kpi.metric||''}" placeholder="e.g. ARR">
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Section</label>
              <input type="text" id="e-section" class="input-field" value="${kpi.section||''}" list="section-list-modal" placeholder="Section name">
              <datalist id="section-list-modal">${sections.map(s=>`<option value="${s}">`).join('')}</datalist>
            </div>
          </div>

          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Owner (Who)</label>
            <input type="text" id="e-who" class="input-field" value="${kpi.who||''}" placeholder="e.g. CPO, COO">
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Target / Year</label>
              <input type="number" id="e-tgt-fy" class="input-field" value="${kpi.targetFY26??''}" placeholder="Annual target" oninput="App.autoCalcAnnual()">
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Target / Month</label>
              <input type="number" id="e-tgt-month" class="input-field" value="${kpi.targetMonth??''}" placeholder="Monthly target" oninput="App.autoCalcAnnual()">
            </div>
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Operator</label>
              <select id="e-tgt-op" class="input-field">
                <option value="" ${!kpi.targetFY26Op?'selected':''}>None</option>
                <option value=">" ${kpi.targetFY26Op==='>'?'selected':''}>Greater than (&gt;)</option>
                <option value=">=" ${kpi.targetFY26Op==='>='?'selected':''}>At least (≥)</option>
                <option value="<" ${kpi.targetFY26Op==='<'?'selected':''}>Less than (&lt;)</option>
                <option value="<=" ${kpi.targetFY26Op==='<='?'selected':''}>At most (≤)</option>
              </select>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Unit</label>
              <select id="e-unit" class="input-field">
                <option value=""  ${!kpi.unit?'selected':''}>Number (#)</option>
                <option value="$" ${kpi.unit==='$'?'selected':''}>Dollar ($)</option>
                <option value="%" ${kpi.unit==='%'?'selected':''}>Percent (%)</option>
              </select>
            </div>
          </div>

          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">YTD Actual</label>
            <input type="number" id="e-ytd" class="input-field" value="${kpi.ytd??''}" placeholder="Year to date value">
            <div style="font-size:11px;color:var(--text-secondary);margin-top:4px">In monthly entry mode this is auto-calculated from monthly data. You can also set it manually here.</div>
          </div>

          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Monthly Aggregation Method</label>
            <select id="e-ytdmethod" class="input-field">
              <option value="sum"  ${(kpi.ytdMethod||'sum')==='sum' ?'selected':''}>∑ Sum — add all months (revenue, volume, count)</option>
              <option value="avg"  ${kpi.ytdMethod==='avg' ?'selected':''}>x̄ Average — mean of months (%, rates, scores)</option>
              <option value="last" ${kpi.ytdMethod==='last'?'selected':''}>→ Latest month — point-in-time value</option>
              <option value="max"  ${kpi.ytdMethod==='max' ?'selected':''}>↑ Highest month — peak performance</option>
              <option value="min"  ${kpi.ytdMethod==='min' ?'selected':''}>↓ Lowest month — best low-is-good metric</option>
            </select>
            <div style="font-size:11px;color:var(--text-secondary);margin-top:4px">Controls how monthly actuals combine for YTD, quarterly and full-year views. Also affects period targets and RAG.</div>
          </div>

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">RAG Rule</label>
              <select id="e-threshold" class="input-field">
                ${thresholds.map(th=>`<option value="${th.id}" ${kpi.thresholdId===th.id?'selected':''}>${th.name}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">RAG Override</label>
              <select id="e-rag-override" class="input-field">
                <option value="" ${!kpi.ragOverride?'selected':''}>Auto</option>
                <option value="green"  ${kpi.ragOverride==='green'?'selected':''}>● Green</option>
                <option value="amber"  ${kpi.ragOverride==='amber'?'selected':''}>▲ Amber</option>
                <option value="red"    ${kpi.ragOverride==='red'?'selected':''}>✕ Red</option>
                <option value="neutral"${kpi.ragOverride==='neutral'?'selected':''}>○ N/A</option>
              </select>
            </div>
          </div>

          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Comment / Note</label>
            <input type="text" id="e-comment" class="input-field" value="${kpi.comment||''}" placeholder="Optional note">
          </div>

          <div style="margin-bottom:16px;display:flex;align-items:center;gap:10px;padding:10px 12px;background:rgba(0,194,168,0.06);border-radius:8px;border:1px solid rgba(0,194,168,0.15)">
            <input type="checkbox" id="e-iskey" ${kpi.isKey?'checked':''} style="width:16px;height:16px;accent-color:var(--brand-accent)">
            <div>
              <label for="e-iskey" style="font-size:13px;font-weight:500;cursor:pointer">Key Metric</label>
              <div style="font-size:11px;color:var(--text-secondary)">Show this KPI on the Overview dashboard</div>
            </div>
          </div>

          <div class="divider"></div>
          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" style="flex:1" onclick="App.saveKpiEdit('${kpi.id||''}',${isNew})">
              ${isNew?'Add KPI':'Save Changes'}
            </button>
            ${!isNew?`<button class="btn btn-ghost" style="color:var(--rag-red)" onclick="App.confirmRemoveKpi('${kpi.id}')">Delete</button>`:''}
            <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
          </div>
        </div>
      </div>`;
  }

  // ── Threshold Edit Modal ──────────────────────────────────────────────────
  function thresholdEditModal(th, isNew=false) {
    const levelsHtml = (th?.levels||[]).map((lv,i)=>`
      <div style="display:grid;grid-template-columns:90px 60px 120px 1fr auto;gap:8px;align-items:center;margin-bottom:8px" id="level-row-${i}">
        <select class="input-field" style="padding:6px 8px;font-size:12px" id="lv-rag-${i}">
          ${['green','amber','red'].map(r=>`<option value="${r}" ${lv.rag===r?'selected':''}>${r}</option>`).join('')}
        </select>
        <select class="input-field" style="padding:6px 8px;font-size:12px" id="lv-op-${i}">
          ${['>','>=','<','<=','='].map(op=>`<option value="${op}" ${lv.op===op?'selected':''}>${op}</option>`).join('')}
        </select>
        <input type="number" class="input-field" style="padding:6px 8px;font-size:12px" id="lv-val-${i}" value="${lv.value}" step="0.01">
        <input type="text" class="input-field" style="padding:6px 8px;font-size:12px" id="lv-lbl-${i}" value="${lv.label||''}" placeholder="Label">
        <button onclick="document.getElementById('level-row-${i}').remove()" style="color:var(--rag-red);background:none;border:none;cursor:pointer;font-size:16px">✕</button>
      </div>`).join('');

    return `
      <div class="modal-overlay" id="th-edit-modal" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:560px">
          <div style="display:flex;justify-content:space-between;margin-bottom:12px">
            <h3 style="font-family:var(--font-display);font-size:16px;font-weight:700">${isNew?'New Colour Rule':'Edit Colour Rule'}</h3>
            <button onclick="App.closeModal()" style="color:var(--text-muted);font-size:20px;padding:4px">✕</button>
          </div>
          <div style="background:rgba(58,134,255,0.07);border:1px solid rgba(58,134,255,0.15);border-radius:8px;padding:10px 14px;margin-bottom:16px;font-size:12px;color:var(--text-secondary);line-height:1.6">
            This rule controls <strong style="color:var(--text-primary)">what colour (Green/Amber/Red) the KPI shows</strong>.
            It does <em>not</em> set the target — that's done per-KPI in the data table.
          </div>
          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Name *</label>
            <input type="text" id="th-name" class="input-field" value="${th?.name||''}" placeholder="e.g. Revenue — High Good">
          </div>
          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Description</label>
            <input type="text" id="th-desc" class="input-field" value="${th?.description||''}" placeholder="Short explanation">
          </div>
          <div style="margin-bottom:16px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Comparison Type</label>
            <select id="th-type" class="input-field">
              <option value="relative"       ${th?.type==='relative'?'selected':''}>Relative — % of target (higher = better)</option>
              <option value="relative_lower" ${th?.type==='relative_lower'?'selected':''}>Relative — % of target (lower = better)</option>
              <option value="absolute"       ${th?.type==='absolute'?'selected':''}>Absolute — compare actual value directly</option>
              <option value="manual"         ${th?.type==='manual'?'selected':''}>Manual only — no auto-calculation</option>
            </select>
          </div>
          <div style="margin-bottom:8px">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
              <label class="label-sm">Levels (evaluated top-down, first match wins)</label>
              <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px" onclick="App.addThresholdLevel()">+ Add Level</button>
            </div>
            <div style="display:grid;grid-template-columns:90px 60px 120px 1fr auto;gap:8px;margin-bottom:6px">
              <div class="label-xs">RAG</div><div class="label-xs">Op</div>
              <div class="label-xs">Value</div><div class="label-xs">Label</div><div></div>
            </div>
            <div id="th-levels-container">${levelsHtml}</div>
          </div>
          <div class="divider"></div>
          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" style="flex:1" onclick="App.saveThreshold('${th?.id||''}',${isNew})">
              ${isNew?'Create Threshold':'Save Changes'}
            </button>
            ${th&&!['th_manual_only','th_higher_better','th_lower_better','th_percent_high','th_percent_low'].includes(th.id)?
              `<button class="btn btn-ghost" style="color:var(--rag-red)" onclick="App.confirmRemoveThreshold('${th.id}')">Delete</button>`:''}
            <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
          </div>
        </div>
      </div>`;
  }

  // ── XLSX import panel ─────────────────────────────────────────────────────
  function xlsxImportPanel() {
    return `
      <div class="card" style="border:2px dashed var(--brand-accent);background:rgba(0,194,168,0.04);text-align:center;padding:28px">
        <div style="font-size:28px;margin-bottom:10px"><i class="fa-solid fa-chart-bar"></i></div>
        <div style="font-family:var(--font-display);font-size:15px;font-weight:600;margin-bottom:6px">Import XLSX / CSV</div>
        <div style="font-size:12px;color:var(--text-secondary);margin-bottom:14px;line-height:1.6">
          Upload a spreadsheet with columns:<br>
          <strong>A: Metric &nbsp;·&nbsp; B: Who &nbsp;·&nbsp; C: Target FY &nbsp;·&nbsp; D: YTD &nbsp;·&nbsp; E: Target Month</strong><br>
          Matching KPIs are updated. New ones are added.
        </div>
        <input type="file" id="xlsx-file-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="App.handleXlsxUpload(event)">
        <button class="btn btn-primary" onclick="document.getElementById('xlsx-file-input').click()">↑ Choose File</button>
        <div style="font-size:10px;color:var(--text-muted);margin-top:10px">Supports .xlsx, .xls, .csv</div>
      </div>`;
  }

  // ── Formula KPI Modal ─────────────────────────────────────────────────────
  function formulaKpiModal(kpi, fml, sourceKpis, isEdit) {
    const sections   = DataStore.getSections();
    const thresholds = DataStore.getThresholds();
    const opLabels   = [
      { v:'sum',      l:'Sum (+)',            desc:'Add all selected KPIs together' },
      { v:'subtract', l:'Subtract (−)',       desc:'First KPI minus all others' },
      { v:'multiply', l:'Multiply (×)',       desc:'Multiply all selected KPIs' },
      { v:'divide',   l:'Divide (÷)',         desc:'First KPI divided by the second' },
      { v:'avg',      l:'Average',            desc:'Mean of all selected KPIs' },
      { v:'min',      l:'Minimum',            desc:'Lowest value among selected KPIs' },
      { v:'max',      l:'Maximum',            desc:'Highest value among selected KPIs' },
      { v:'custom',   l:'Custom Expression',  desc:'Write your own formula using KPI IDs' },
    ];

    const curOp      = fml?.op || 'sum';
    const curIds     = fml?.operands || [];
    const curExpr    = fml?.expression || '';

    const sectionOpts = [...new Set([...sections, 'Formulas'])].map(s =>
      `<option value="${_esc(s)}" ${(kpi?.section||'Formulas')===s?'selected':''}>${_esc(s)}</option>`
    ).join('');

    const thresholdOpts = thresholds.map(t =>
      `<option value="${t.id}" ${(kpi?.thresholdId||'th_manual_only')===t.id?'selected':''}>${_esc(t.name)}</option>`
    ).join('');

    const opOptions = opLabels.map(o => `
      <div onclick="OverviewPanel._fmlSelectOp('${o.v}')"
           id="fml-op-btn-${o.v}"
           style="display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;cursor:pointer;
                  border:1px solid ${curOp===o.v?'var(--brand-accent)':'var(--border-card)'};
                  background:${curOp===o.v?'rgba(0,194,168,0.08)':'var(--bg-input)'};
                  transition:all 0.15s;margin-bottom:6px"
           onmouseover="this.style.borderColor='var(--brand-accent)'"
           onmouseout="if(document.getElementById('fml-op').value!=='${o.v}')this.style.borderColor='var(--border-card)'">
        <div style="flex:1">
          <div style="font-size:13px;font-weight:600;color:${curOp===o.v?'var(--brand-accent)':'var(--text-primary)'}">${o.l}</div>
          <div style="font-size:11px;color:var(--text-muted);margin-top:1px">${o.desc}</div>
        </div>
        <div id="fml-op-chk-${o.v}" style="width:16px;height:16px;border-radius:50%;border:2px solid ${curOp===o.v?'var(--brand-accent)':'var(--border-card)'};display:flex;align-items:center;justify-content:center;flex-shrink:0">
          ${curOp===o.v?'<div style="width:8px;height:8px;border-radius:50%;background:var(--brand-accent)"></div>':''}
        </div>
      </div>`).join('');

    const kpiCheckboxes = sourceKpis.length === 0
      ? '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No source KPIs available</div>'
      : sourceKpis.map(k => {
          const ps = DataStore.getPeriodStats(k, DataStore.getSettings().reportingPeriod || 'monthly');
          return `
            <label style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:6px;cursor:pointer;transition:background 0.1s;border-bottom:1px solid var(--border-subtle)"
                   onmouseover="this.style.background='var(--bg-card-hover)'" onmouseout="this.style.background='transparent'">
              <input type="checkbox" class="fml-operand-cb" value="${k.id}" ${curIds.includes(k.id)?'checked':''}
                     style="accent-color:var(--brand-accent);width:15px;height:15px;flex-shrink:0;cursor:pointer">
              <div style="flex:1;min-width:0">
                <div style="font-size:12px;font-weight:500;color:var(--text-primary);line-height:1.3">${_esc(k.metric)}</div>
                <div style="font-size:10px;color:var(--text-muted)">${_esc(k.section)}${k.unit?' · '+k.unit:''}</div>
              </div>
              <div style="font-size:12px;font-weight:600;color:var(--text-secondary);font-family:var(--font-mono);flex-shrink:0">
                ${ps.actual !== null ? DataStore.formatValue(ps.actual, k) : '—'}
              </div>
            </label>`;
        }).join('');

    // Build the KPI ID reference table for custom expressions
    const idRefTable = sourceKpis.map(k =>
      `<tr>
        <td style="padding:3px 8px;font-family:var(--font-mono);font-size:10px;color:var(--brand-accent);white-space:nowrap">${k.id}</td>
        <td style="padding:3px 8px;font-size:10px;color:var(--text-secondary)">${_esc(k.metric)}</td>
        <td style="padding:3px 8px;font-size:10px;color:var(--text-muted);font-family:var(--font-mono)">${DataStore.getPeriodStats(k, 'monthly').actual !== null ? DataStore.formatValue(DataStore.getPeriodStats(k,'monthly').actual, k) : '—'}</td>
       </tr>`
    ).join('');

    return `
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:640px;max-height:92vh;overflow-y:auto">

          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
            <div>
              <h3 style="font-family:var(--font-display);font-size:17px;font-weight:700;margin:0 0 3px">
                ${isEdit ? '✎ Edit Formula KPI' : '∑ New Formula KPI'}
              </h3>
              <div style="font-size:12px;color:var(--text-muted)">Computed automatically from source KPIs</div>
            </div>
            <button onclick="App.closeModal()" style="color:var(--text-muted);font-size:20px;background:none;border:none;cursor:pointer;padding:4px">✕</button>
          </div>

          <!-- Step 1: Name & meta -->
          <div style="background:var(--bg-input);border-radius:10px;padding:16px;margin-bottom:16px">
            <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted);margin-bottom:12px">① KPI Details</div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
              <div style="grid-column:1/-1">
                <label class="label-sm" style="display:block;margin-bottom:5px">KPI Name <span style="color:var(--rag-red)">*</span></label>
                <input type="text" id="fml-metric" class="input-field" placeholder="e.g. Total Expenses, Net Revenue"
                       value="${_esc(kpi?.metric||'')}">
              </div>
              <div>
                <label class="label-sm" style="display:block;margin-bottom:5px">Section</label>
                <select id="fml-section" class="input-field">
                  <option value="Formulas">Formulas</option>
                  ${sectionOpts}
                </select>
              </div>
              <div>
                <label class="label-sm" style="display:block;margin-bottom:5px">Unit</label>
                <select id="fml-unit" class="input-field">
                  <option value="" ${!kpi?.unit?'selected':''}>None</option>
                  <option value="$" ${kpi?.unit==='$'?'selected':''}>$ Currency</option>
                  <option value="%" ${kpi?.unit==='%'?'selected':''}>% Percentage</option>
                </select>
              </div>
              <div>
                <label class="label-sm" style="display:block;margin-bottom:5px">Threshold</label>
                <select id="fml-threshold" class="input-field">${thresholdOpts}</select>
              </div>
              <div style="display:flex;align-items:center;gap:8px;padding-top:18px">
                <input type="checkbox" id="fml-iskey" ${kpi?.isKey?'checked':''}
                       style="accent-color:var(--brand-accent);width:15px;height:15px;cursor:pointer">
                <label for="fml-iskey" class="label-sm" style="cursor:pointer"><i class="fa-solid fa-star"></i> Pin to Overview (Key KPI)</label>
              </div>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Comment / Note</label>
              <input type="text" id="fml-comment" class="input-field" placeholder="Optional note"
                     value="${_esc(kpi?.comment||'')}">
            </div>
          </div>

          <!-- Step 2: Operation -->
          <div style="background:var(--bg-input);border-radius:10px;padding:16px;margin-bottom:16px">
            <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted);margin-bottom:12px">② Formula Operation</div>
            <input type="hidden" id="fml-op" value="${curOp}">
            <div style="max-height:260px;overflow-y:auto;padding-right:4px">${opOptions}</div>
          </div>

          <!-- Step 3: Source KPIs (hidden when custom) -->
          <div id="fml-source-section" style="background:var(--bg-input);border-radius:10px;padding:16px;margin-bottom:16px;${curOp==='custom'?'display:none':''}">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
              <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted)">③ Source KPIs</div>
              <div style="display:flex;gap:6px">
                <button onclick="document.querySelectorAll('.fml-operand-cb').forEach(c=>c.checked=true)" class="btn btn-ghost" style="font-size:10px;padding:3px 8px">All</button>
                <button onclick="document.querySelectorAll('.fml-operand-cb').forEach(c=>c.checked=false)" class="btn btn-ghost" style="font-size:10px;padding:3px 8px">None</button>
              </div>
            </div>
            <div style="font-size:11px;color:var(--text-muted);margin-bottom:8px">Select the KPIs to include in this formula</div>
            <div style="max-height:220px;overflow-y:auto;border:1px solid var(--border-subtle);border-radius:8px;background:var(--bg-card)">
              ${kpiCheckboxes}
            </div>
          </div>

          <!-- Step 3b: Custom expression (shown when custom) -->
          <div id="fml-custom-section" style="background:var(--bg-input);border-radius:10px;padding:16px;margin-bottom:16px;${curOp!=='custom'?'display:none':''}">
            <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted);margin-bottom:10px">③ Custom Expression</div>
            <div style="font-size:12px;color:var(--text-secondary);margin-bottom:8px;line-height:1.6">
              Use KPI IDs as variables. Supports: <code style="background:var(--bg-card);padding:1px 5px;border-radius:3px;font-size:11px">+ - * / ( )</code> and numeric constants.
            </div>
            <input type="text" id="fml-expression" class="input-field" style="font-family:var(--font-mono);font-size:13px;margin-bottom:10px"
                   placeholder="e.g. kpi_1 + kpi_2 - kpi_3 * 0.1"
                   value="${_esc(curExpr)}">
            <div style="border:1px solid var(--border-subtle);border-radius:8px;overflow:hidden">
              <table style="width:100%;border-collapse:collapse">
                <thead>
                  <tr style="background:var(--bg-card)">
                    <th style="padding:5px 8px;font-size:10px;font-weight:600;color:var(--text-muted);text-align:left">KPI ID</th>
                    <th style="padding:5px 8px;font-size:10px;font-weight:600;color:var(--text-muted);text-align:left">Metric</th>
                    <th style="padding:5px 8px;font-size:10px;font-weight:600;color:var(--text-muted);text-align:left">Current</th>
                  </tr>
                </thead>
                <tbody>${idRefTable}</tbody>
              </table>
            </div>
          </div>

          <!-- Actions -->
          <div style="display:flex;gap:8px;justify-content:flex-end">
            <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
            <button class="btn btn-primary" style="padding:9px 22px"
                    onclick="App.saveFormulaKpi('${kpi?.id||''}', ${isEdit})">
              ${isEdit ? '✓ Save Changes' : '∑ Create Formula KPI'}
            </button>
          </div>
        </div>
      </div>`;
  }

  return {
    ragBadge, ragDot, kpiCard, sidebar, topBar,
    ragSummaryBar, kpiDetailModal, editKpiModal,
    thresholdEditModal, xlsxImportPanel, formulaKpiModal,
  };
})();
/**
 * overview-panel.js  v1
 * Replaces Pages.overview() with a rich, customisable dashboard.
 *
 * HERO WIDGET — large central card the client configures:
 *   • Chart type: odometer | bar | line | pie   (via ••• menu)
 *   • KPIs included: dropdown + search + checkbox picker
 *   • Position: centred (standalone) or embedded in the grid
 *
 * SURROUNDING CARDS — the normal key-metric KPI cards, arranged around the hero.
 *
 * STATE is kept in OverviewPanel (module-level, survives renders because
 * the render pipeline calls Pages.overview() fresh each time — we persist
 * state in a settings sub-key via DataStore so it survives page reloads).
 */

const OverviewPanel = (() => {

  // ─────────────────────────────────────────────────────────────────────────
  // State helpers — persisted inside DataStore.getSettings().overviewPanel
  // ─────────────────────────────────────────────────────────────────────────
  function _getState() {
    const s = DataStore.getSettings();
    return s.overviewPanel || {
      chartType:      'odometer',   // 'odometer' | 'bar' | 'line' | 'pie'
      heroKpiIds:     [],           // KPI ids shown inside the hero card
      heroPosition:   'center',     // 'center' (full width standalone) | 'grid' (sits in flow)
      kpiPickerOpen:  false,
      menuOpen:       false,
    };
  }

  function _setState(patch) {
    const s   = DataStore.getSettings();
    const cur = _getState();
    DataStore.updateSettings({ overviewPanel: { ...cur, ...patch } });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Public API called from inline handlers
  // ─────────────────────────────────────────────────────────────────────────
  function setChartType(type) {
    _setState({ chartType: type, menuOpen: false });
  }

  function toggleMenu() {
    const st = _getState();
    _setState({ menuOpen: !st.menuOpen });
  }

  function closeMenu() { _setState({ menuOpen: false }); }

  function toggleKpiPicker() {
    const st = _getState();
    _setState({ kpiPickerOpen: !st.kpiPickerOpen });
  }

  function toggleHeroKpi(id, checked) {
    const st  = _getState();
    let ids   = [...(st.heroKpiIds || [])];
    if (checked && !ids.includes(id)) ids.push(id);
    if (!checked) ids = ids.filter(x => x !== id);
    _setState({ heroKpiIds: ids });
  }

  function setHeroAllKpis(selected) {
    const allKpis = DataStore.getKpis();
    _setState({ heroKpiIds: selected ? allKpis.map(k => k.id) : [] });
  }

  function setHeroPosition(pos) {
    _setState({ heroPosition: pos });
  }

  function filterHeroKpiSearch(q) {
    const list = document.getElementById('hero-kpi-list');
    if (!list) return;
    const query = q.toLowerCase().trim();
    list.querySelectorAll('.hero-kpi-item').forEach(item => {
      const name = item.querySelector('.hki-name')?.textContent?.toLowerCase() || '';
      item.style.display = (!query || name.includes(query)) ? '' : 'none';
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Chart renderers — called post-render via OverviewPanel.mountCharts()
  // ─────────────────────────────────────────────────────────────────────────
  function mountCharts() {
    const st      = _getState();
    const heroKpiIds = st.heroKpiIds || [];
    const allKpis = DataStore.getKpis();
    const kpis    = heroKpiIds.length > 0
      ? allKpis.filter(k => heroKpiIds.includes(k.id))
      : DataStore.getKeyKpis();

    if (kpis.length === 0) return;

    const settings   = DataStore.getSettings();
    const mode       = settings.reportingPeriod || 'monthly';

    switch (st.chartType) {
      case 'odometer': _mountOdometers(kpis, mode, settings); break;
      case 'bar':      _mountBar(kpis, mode, settings);       break;
      case 'line':     _mountLine(kpis, mode, settings);      break;
      case 'pie':      _mountPie(kpis, mode, settings);       break;
    }
  }

  // ── CSS variable reader ──
  function _cssVar(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  }

  // ── RAG → colour map ──
  const RAG_COLOR = {
    green:   () => _cssVar('--rag-green')   || '#00C48C',
    amber:   () => _cssVar('--rag-amber')   || '#FF9F1C',
    red:     () => _cssVar('--rag-red')     || '#FF4444',
    neutral: () => _cssVar('--rag-neutral') || '#6B7A99',
  };

  // ─── ODOMETER / GAUGE ────────────────────────────────────────────────────
  function _mountOdometers(kpis, mode, settings) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;

    // Build gauges HTML using the Revmeter-style SVG approach
    const fmt = (v, kpi) => v === null || v === undefined ? '—' : DataStore.formatValue(v, kpi);

    const gaugeHtml = kpis.map((kpi, idx) => {
      const ps      = DataStore.getPeriodStats(kpi, mode);
      const actual  = ps.actual ?? 0;
      const target  = ps.target ?? (kpi.targetFY26 ?? 1);
      const rag     = ps.rag || kpi.rag || 'neutral';
      const color   = RAG_COLOR[rag]?.() || '#6B7A99';
      const safeT   = Math.max(1, target);
      const maxVal  = safeT / 0.75;
      const displayVal = Math.max(0, Math.min(actual, maxVal));
      const pct     = displayVal / maxVal;
      const circumf = 251.3;
      const offset  = circumf - (pct * circumf);
      const rotation= pct * 180;

      // Target marker angle (75% of arc = target position)
      const targetAngle = 0.75 * 180; // target sits at 75% of 180° sweep
      const tx = 100 + 80 * Math.cos((Math.PI - (targetAngle * Math.PI / 180)));
      const ty = 100 - 80 * Math.sin((targetAngle * Math.PI / 180));
      const tx2 = 100 + 70 * Math.cos((Math.PI - (targetAngle * Math.PI / 180)));
      const ty2 = 100 - 70 * Math.sin((targetAngle * Math.PI / 180));

      const isMain = idx === 0 && kpis.length <= 3;
      const size   = isMain ? 'odo-main' : 'odo-small';

      return `
        <div class="odo-block ${size}">
          <div class="odo-title">${_esc(kpi.metric)}</div>
          <svg viewBox="0 0 200 115" class="odo-svg">
            <!-- Track -->
            <path d="M 20 100 A 80 80 0 0 1 180 100"
                  fill="none" stroke="rgba(255,255,255,0.06)" stroke-width="14" stroke-linecap="round"/>
            <!-- Tick marks -->
            <path d="M 20 100 A 80 80 0 0 1 180 100"
                  fill="none" stroke="rgba(255,255,255,0.08)" stroke-width="18"
                  stroke-dasharray="2 14"/>
            <!-- Fill arc -->
            <path id="odo-fill-${kpi.id}" d="M 20 100 A 80 80 0 0 1 180 100"
                  fill="none" stroke="${color}" stroke-width="14" stroke-linecap="round"
                  stroke-dasharray="${circumf}" stroke-dashoffset="${circumf}"
                  style="transition:stroke-dashoffset 0.9s cubic-bezier(0.4,0,0.2,1),stroke 0.4s ease"/>
            <!-- Target marker line -->
            <line x1="${tx}" y1="${ty}" x2="${tx2}" y2="${ty2}"
                  stroke="rgba(255,255,255,0.4)" stroke-width="2.5"/>
            <!-- Needle -->
            <line id="odo-needle-${kpi.id}" x1="100" y1="100" x2="20" y2="100"
                  stroke="rgba(255,255,255,0.7)" stroke-width="2.5" stroke-linecap="round"
                  style="transform-origin:100px 100px;transition:transform 0.9s cubic-bezier(0.4,0,0.2,1)"/>
            <!-- Centre hub -->
            <circle cx="100" cy="100" r="7" fill="var(--bg-page)" stroke="rgba(255,255,255,0.2)" stroke-width="2"/>
          </svg>
          <div class="odo-value" style="color:${color}">${fmt(ps.actual, kpi)}</div>
          <div class="odo-meta">
            <span>Target: ${fmt(ps.target ?? kpi.targetFY26, kpi)}</span>
            <span class="odo-rag odo-rag-${rag}">${rag.charAt(0).toUpperCase()+rag.slice(1)}</span>
          </div>
          <div class="odo-progress-wrap">
            <div class="odo-progress-fill" style="width:${ps.progressPct||0}%;background:${color}"></div>
          </div>
        </div>`;
    }).join('');

    container.innerHTML = `<div class="odo-grid">${gaugeHtml}</div>`;

    // Animate after paint
    requestAnimationFrame(() => requestAnimationFrame(() => {
      kpis.forEach(kpi => {
        const ps       = DataStore.getPeriodStats(kpi, mode);
        const actual   = ps.actual ?? 0;
        const target   = ps.target ?? (kpi.targetFY26 ?? 1);
        const safeT    = Math.max(1, target);
        const maxVal   = safeT / 0.75;
        const displayV = Math.max(0, Math.min(actual, maxVal));
        const pct      = displayV / maxVal;
        const circumf  = 251.3;
        const offset   = circumf - pct * circumf;
        const rotation = pct * 180;

        const fillEl   = document.getElementById(`odo-fill-${kpi.id}`);
        const needleEl = document.getElementById(`odo-needle-${kpi.id}`);
        if (fillEl)   fillEl.style.strokeDashoffset = offset;
        if (needleEl) needleEl.style.transform = `rotate(${rotation}deg)`;
      });
    }));
  }

  // ─── BAR CHART ────────────────────────────────────────────────────────────
  function _mountBar(kpis, mode, settings) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;

    const W = container.clientWidth || 700;
    const H = 280;
    const pad = { top: 24, right: 20, bottom: 60, left: 52 };
    const chartW = W - pad.left - pad.right;
    const chartH = H - pad.top - pad.bottom;

    const data = kpis.map(kpi => {
      const ps = DataStore.getPeriodStats(kpi, mode);
      return {
        label:  kpi.metric.length > 14 ? kpi.metric.slice(0, 13) + '…' : kpi.metric,
        actual: ps.actual ?? 0,
        target: ps.target ?? kpi.targetFY26 ?? 0,
        rag:    ps.rag || kpi.rag || 'neutral',
        unit:   kpi.unit || '',
      };
    });

    const maxVal = Math.max(...data.flatMap(d => [d.actual, d.target]), 1);
    const barW   = Math.min(40, chartW / data.length / 2.5);
    const groupW = chartW / data.length;
    const yScale = v => chartH - (v / maxVal) * chartH;

    // Y axis ticks
    const ticks = 5;
    const tickLines = Array.from({ length: ticks + 1 }, (_, i) => {
      const v = (maxVal / ticks) * i;
      const y = pad.top + yScale(v);
      const label = _shortNum(v);
      return `
        <line x1="${pad.left}" y1="${y}" x2="${W - pad.right}" y2="${y}"
              stroke="rgba(255,255,255,0.05)" stroke-width="1"/>
        <text x="${pad.left - 6}" y="${y + 4}" fill="rgba(255,255,255,0.3)"
              font-size="10" text-anchor="end">${label}</text>`;
    }).join('');

    const bars = data.map((d, i) => {
      const cx     = pad.left + groupW * i + groupW / 2;
      const color  = RAG_COLOR[d.rag]?.() || '#6B7A99';
      const aH     = (d.actual / maxVal) * chartH;
      const tH     = 3; // target line height
      const ay     = pad.top + yScale(d.actual);
      const ty     = pad.top + yScale(d.target);
      const barX   = cx - barW / 2;

      return `
        <!-- Bar -->
        <rect class="bar-anim" x="${barX}" y="${ay}" width="${barW}" height="${aH}"
              rx="3" fill="${color}" opacity="0.85"
              style="transform-origin:${barX + barW/2}px ${pad.top + chartH}px">
          <title>${d.label}: ${d.actual}</title>
        </rect>
        <!-- Target line -->
        <rect x="${cx - barW * 0.7}" y="${ty - 1}" width="${barW * 1.4}" height="2.5"
              rx="1" fill="rgba(255,255,255,0.45)"/>
        <!-- Label -->
        <text x="${cx}" y="${pad.top + chartH + 18}" fill="rgba(255,255,255,0.55)"
              font-size="10" text-anchor="middle">${_esc(d.label)}</text>`;
    }).join('');

    const legend = `
      <div style="display:flex;gap:16px;justify-content:center;margin-top:12px;flex-wrap:wrap">
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:rgba(255,255,255,0.5)">
          <div style="width:24px;height:8px;border-radius:2px;background:rgba(255,255,255,0.6)"></div> Target
        </div>
        <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:rgba(255,255,255,0.5)">
          <div style="width:14px;height:14px;border-radius:2px;background:var(--rag-green)"></div> On Track
          <div style="width:14px;height:14px;border-radius:2px;background:var(--rag-amber)"></div> At Risk
          <div style="width:14px;height:14px;border-radius:2px;background:var(--rag-red)"></div> Off Track
        </div>
      </div>`;

    container.innerHTML = `
      <svg width="100%" height="${H}" viewBox="0 0 ${W} ${H}" style="overflow:visible">
        ${tickLines}
        ${bars}
      </svg>
      ${legend}`;

    // Animate bars in
    requestAnimationFrame(() => requestAnimationFrame(() => {
      container.querySelectorAll('.bar-anim').forEach((el, i) => {
        el.style.transition = `transform 0.7s cubic-bezier(0.4,0,0.2,1) ${i * 0.07}s, opacity 0.4s ease ${i * 0.07}s`;
        el.style.opacity = '0.85';
      });
    }));
  }

  // ─── LINE CHART ───────────────────────────────────────────────────────────
  function _mountLine(kpis, mode, settings) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;

    const fyMonths = DataStore.getFyMonths();
    const W = container.clientWidth || 700;
    const H = 280;
    const pad = { top: 24, right: 24, bottom: 50, left: 52 };
    const chartW = W - pad.left - pad.right;
    const chartH = H - pad.top - pad.bottom;

    // Collect all monthly actuals across the selected KPIs
    const allVals = kpis.flatMap(kpi =>
      fyMonths.map(m => kpi.monthlyActuals?.[m]).filter(v => v !== undefined && v !== null)
    ).map(Number);

    const maxVal = Math.max(...allVals, 1);
    const xStep  = chartW / Math.max(fyMonths.length - 1, 1);
    const yScale = v => chartH - (v / maxVal) * chartH;

    // Grid
    const gridLines = Array.from({ length: 5 }, (_, i) => {
      const v = (maxVal / 4) * i;
      const y = pad.top + yScale(v);
      return `
        <line x1="${pad.left}" y1="${y}" x2="${W - pad.right}" y2="${y}"
              stroke="rgba(255,255,255,0.05)" stroke-width="1"/>
        <text x="${pad.left - 6}" y="${y + 4}" fill="rgba(255,255,255,0.3)"
              font-size="10" text-anchor="end">${_shortNum(v)}</text>`;
    }).join('');

    // X labels
    const xLabels = fyMonths.map((m, i) => {
      const x = pad.left + i * xStep;
      const show = fyMonths.length <= 12 || i % 2 === 0;
      return show ? `<text x="${x}" y="${pad.top + chartH + 16}" fill="rgba(255,255,255,0.35)"
                           font-size="10" text-anchor="middle">${m}</text>` : '';
    }).join('');

    // Series lines
    const PALETTE = ['#00C2A8','#3A86FF','#FF9F1C','#FF4444','#A78BFA','#34D399'];
    const lines = kpis.map((kpi, ki) => {
      const color  = PALETTE[ki % PALETTE.length];
      const points = fyMonths
        .map((m, i) => {
          const v = kpi.monthlyActuals?.[m];
          return v !== undefined && v !== null ? { x: pad.left + i * xStep, y: pad.top + yScale(Number(v)) } : null;
        })
        .filter(Boolean);

      if (points.length < 2) return '';
      const d = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');

      // Area fill
      const areaD = `${d} L ${points[points.length-1].x} ${pad.top + chartH} L ${points[0].x} ${pad.top + chartH} Z`;

      return `
        <defs>
          <linearGradient id="lg-${ki}" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="${color}" stop-opacity="0.25"/>
            <stop offset="100%" stop-color="${color}" stop-opacity="0"/>
          </linearGradient>
        </defs>
        <path d="${areaD}" fill="url(#lg-${ki})"/>
        <path d="${d}" fill="none" stroke="${color}" stroke-width="2.5"
              stroke-linecap="round" stroke-linejoin="round" class="line-path"/>
        ${points.map(p => `<circle cx="${p.x}" cy="${p.y}" r="3" fill="${color}" opacity="0.9"/>`).join('')}`;
    }).join('');

    // Legend
    const legend = kpis.map((kpi, ki) => {
      const color = PALETTE[ki % PALETTE.length];
      return `<div style="display:flex;align-items:center;gap:6px;font-size:11px;color:rgba(255,255,255,0.6)">
        <div style="width:20px;height:3px;border-radius:2px;background:${color}"></div>${_esc(kpi.metric)}
      </div>`;
    }).join('');

    container.innerHTML = `
      <svg width="100%" height="${H}" viewBox="0 0 ${W} ${H}" style="overflow:visible">
        ${gridLines}${xLabels}${lines}
      </svg>
      <div style="display:flex;gap:14px;justify-content:center;margin-top:10px;flex-wrap:wrap">${legend}</div>`;
  }

  // ─── PIE / DONUT CHART ────────────────────────────────────────────────────
  function _mountPie(kpis, mode, settings) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;

    const mode2 = DataStore.getSettings().reportingPeriod || 'monthly';
    const data = kpis.map((kpi, i) => {
      const ps = DataStore.getPeriodStats(kpi, mode2);
      return {
        label:  kpi.metric,
        value:  Math.max(0, ps.actual ?? kpi.ytd ?? 0),
        rag:    ps.rag || kpi.rag || 'neutral',
      };
    }).filter(d => d.value > 0);

    if (data.length === 0) {
      container.innerHTML = `<div style="text-align:center;color:rgba(255,255,255,0.3);padding:40px">No data available for pie chart</div>`;
      return;
    }

    const PALETTE = ['#00C2A8','#3A86FF','#FF9F1C','#FF4444','#A78BFA','#34D399','#F472B6','#FB923C'];
    const total = data.reduce((s, d) => s + d.value, 0);
    const cx = 110, cy = 110, r = 80, inner = 48;

    let angle = -Math.PI / 2;
    const slices = data.map((d, i) => {
      const sweep = (d.value / total) * Math.PI * 2;
      const x1 = cx + r * Math.cos(angle);
      const y1 = cy + r * Math.sin(angle);
      angle += sweep;
      const x2 = cx + r * Math.cos(angle);
      const y2 = cy + r * Math.sin(angle);
      const ix1 = cx + inner * Math.cos(angle - sweep);
      const iy1 = cy + inner * Math.sin(angle - sweep);
      const ix2 = cx + inner * Math.cos(angle);
      const iy2 = cy + inner * Math.sin(angle);
      const large = sweep > Math.PI ? 1 : 0;
      const color = RAG_COLOR[d.rag]?.() || PALETTE[i % PALETTE.length];
      const midA = angle - sweep / 2;
      const lx = cx + (r + 14) * Math.cos(midA);
      const ly = cy + (r + 14) * Math.sin(midA);
      const pct = Math.round((d.value / total) * 100);

      return {
        path: `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}
               L ${ix2} ${iy2} A ${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1} Z`,
        color, pct, d, lx, ly,
      };
    });

    const paths = slices.map(s => `
      <path d="${s.path}" fill="${s.color}" opacity="0.88"
            class="pie-slice" style="transition:transform 0.2s ease;cursor:pointer;transform-origin:${cx}px ${cy}px"
            onmouseover="this.style.transform='scale(1.04)'" onmouseout="this.style.transform=''">
        <title>${_esc(s.d.label)}: ${s.pct}%</title>
      </path>`).join('');

    const legend = slices.map(s => `
      <div style="display:flex;align-items:center;gap:8px;font-size:12px;color:rgba(255,255,255,0.65)">
        <div style="width:12px;height:12px;border-radius:2px;flex-shrink:0;background:${s.color}"></div>
        <span>${_esc(s.d.label)}</span>
        <span style="margin-left:auto;font-weight:600;color:rgba(255,255,255,0.85)">${s.pct}%</span>
      </div>`).join('');

    container.innerHTML = `
      <div style="display:flex;align-items:center;gap:32px;flex-wrap:wrap;justify-content:center">
        <svg width="220" height="220" viewBox="0 0 220 220">
          ${paths}
          <text x="${cx}" y="${cy - 6}" text-anchor="middle" fill="rgba(255,255,255,0.4)"
                font-size="11" letter-spacing="1">TOTAL</text>
          <text x="${cx}" y="${cy + 14}" text-anchor="middle" fill="rgba(255,255,255,0.85)"
                font-size="16" font-weight="700">${data.length}</text>
          <text x="${cx}" y="${cy + 28}" text-anchor="middle" fill="rgba(255,255,255,0.35)"
                font-size="10">KPIs</text>
        </svg>
        <div style="display:flex;flex-direction:column;gap:10px;min-width:160px;max-width:240px">
          ${legend}
        </div>
      </div>`;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Main HTML renderer — called by Pages.overview()
  // ─────────────────────────────────────────────────────────────────────────
  function render() {
    const settings     = DataStore.getSettings();
    const overviewKpis = DataStore.getOverviewKpis();
    const allKpis      = DataStore.getKpis();
    const sections     = DataStore.getSections();
    const st           = _getState();
    const mode         = settings.reportingPeriod || 'monthly';

    const ragCounts = { green: 0, amber: 0, red: 0, neutral: 0 };
    allKpis.forEach(k => { ragCounts[k.rag || 'neutral']++; });

    // Hero KPIs — what's selected for the big card
    const heroKpiIds  = st.heroKpiIds || [];
    const heroKpis    = heroKpiIds.length > 0
      ? allKpis.filter(k => heroKpiIds.includes(k.id))
      : DataStore.getKeyKpis();

    // Surrounding cards — key KPIs NOT in the hero (or all if hero is grid-embedded)
    const surroundKpis = overviewKpis;

    // ── Header ──────────────────────────────────────────────────────────────
    const header = _renderHeader(settings, overviewKpis, allKpis);

    // ── RAG pill strip ───────────────────────────────────────────────────────
    const ragStrip = _renderRagStrip(ragCounts, allKpis.length);

    // ── Hero card ────────────────────────────────────────────────────────────
    const heroCard = _renderHeroCard(st, heroKpis, allKpis, settings);

    // ── Surrounding KPI cards ────────────────────────────────────────────────
    const kpiGrid = surroundKpis.length > 0
      ? `<div style="margin-bottom:12px">
           <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px">
             <h3 style="font-family:var(--font-display);font-size:15px;font-weight:600;margin:0;color:var(--text-primary)">Key Metrics</h3>
             <span class="label-xs" style="color:var(--text-muted)">${surroundKpis.length} pinned · ${mode} view</span>
           </div>
           <div class="grid-auto">
             ${surroundKpis.map(kpi => Components.kpiCard(kpi, DataStore.getPeriodStats(kpi, mode), true)).join('')}
           </div>
         </div>`
      : allKpis.length > 0 ? '' : `
         <div class="card" style="text-align:center;padding:40px;margin-bottom:24px">
           <div style="font-size:32px;margin-bottom:12px">◈</div>
           <div style="font-size:16px;font-weight:600;margin-bottom:6px">No Key Metrics yet</div>
           <div style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">
             In Data Entry, tick the <strong>Ovw</strong> checkbox next to any KPI to pin it here.
           </div>
           <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
         </div>`;

    // ── Section grid ─────────────────────────────────────────────────────────
    const sectionGrid = sections.length > 0 ? `
      <div style="margin-top:8px">
        <h3 style="font-family:var(--font-display);font-size:15px;font-weight:600;margin-bottom:12px;color:var(--text-primary)">All Sections</h3>
        <div class="grid-3">
          ${sections.map(section => {
            const skpis  = DataStore.getKpisBySection(section);
            const counts = { green:0, amber:0, red:0, neutral:0 };
            skpis.forEach(k => { counts[k.rag || 'neutral']++; });
            const pageId = 'section:' + encodeURIComponent(section);
            return `
              <div class="card" style="cursor:pointer;transition:all 0.15s" onclick="App.navigate('${pageId}')"
                   onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform=''">
                <div style="font-family:var(--font-display);font-size:14px;font-weight:600;margin-bottom:8px;line-height:1.3">${_esc(section)}</div>
                <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px">
                  ${Object.entries(counts).filter(([,v])=>v>0).map(([r,c])=>`<span class="rag-badge ${r}" style="font-size:10px;padding:2px 7px">${c} ${r}</span>`).join('')}
                </div>
                <div class="label-xs">${skpis.length} KPIs &nbsp;→</div>
              </div>`;
          }).join('')}
        </div>
      </div>` : '';

    return `
      <div style="margin-bottom:24px">
        ${header}
        ${ragStrip}
        ${heroCard}
        ${kpiGrid}
        ${sectionGrid}
      </div>

      <style>
        /* ── Odometer / Gauge ── */
        .odo-grid { display:flex;flex-wrap:wrap;gap:24px;justify-content:center;align-items:flex-end;padding:8px 0 }
        .odo-block { display:flex;flex-direction:column;align-items:center }
        .odo-main  { min-width:220px;max-width:280px;flex:0 0 260px }
        .odo-small { min-width:160px;max-width:200px;flex:0 0 180px }
        .odo-svg   { width:100%;display:block;overflow:visible }
        .odo-title { font-size:11px;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-secondary);margin-bottom:6px;text-align:center }
        .odo-value { font-family:var(--font-display);font-size:1.6rem;font-weight:700;margin-top:-4px;letter-spacing:-0.5px }
        .odo-meta  { display:flex;align-items:center;gap:10px;margin-top:4px;font-size:11px;color:var(--text-muted) }
        .odo-rag   { font-size:10px;font-weight:700;padding:2px 7px;border-radius:4px;text-transform:uppercase;letter-spacing:0.06em }
        .odo-rag-green  { color:var(--rag-green);background:var(--rag-green-bg) }
        .odo-rag-amber  { color:var(--rag-amber);background:var(--rag-amber-bg) }
        .odo-rag-red    { color:var(--rag-red);background:var(--rag-red-bg) }
        .odo-rag-neutral{ color:var(--rag-neutral);background:var(--rag-neutral-bg) }
        .odo-progress-wrap { width:80%;height:3px;background:rgba(255,255,255,0.07);border-radius:2px;margin-top:8px;overflow:hidden }
        .odo-progress-fill { height:100%;border-radius:2px;transition:width 0.9s ease }

        /* ── Hero card menu ── */
        .hero-menu-btn { opacity:0.5;transition:opacity 0.15s }
        .hero-menu-btn:hover { opacity:1 }

        /* ── Chart type tabs ── */
        .ct-tab { padding:5px 12px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;transition:all 0.15s;border:1px solid transparent }
        .ct-tab.active { background:rgba(0,194,168,0.15);border-color:rgba(0,194,168,0.35);color:var(--brand-accent) }
        .ct-tab:not(.active) { color:var(--text-muted) }
        .ct-tab:not(.active):hover { color:var(--text-primary);background:rgba(255,255,255,0.04) }

        /* ── KPI picker ── */
        .hero-kpi-item { display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--border-subtle);transition:background 0.1s;cursor:default }
        .hero-kpi-item:hover { background:var(--bg-card-hover) }
        .hki-name { font-size:12px;font-weight:500;color:var(--text-primary);line-height:1.35;word-break:break-word }
        .hki-section { font-size:10px;color:var(--text-muted);margin-top:2px }
      </style>`;
  }

  // ─── Header row ───────────────────────────────────────────────────────────
  function _renderHeader(settings, overviewKpis, allKpis) {
    const overviewIds = settings.overviewKpiIds;
    const manageList = allKpis.length === 0 ? '' : allKpis.map(k => {
      const isVisible = k.isKey && (!overviewIds || overviewIds.includes(k.id));
      return `<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;
                          border-bottom:1px solid var(--border-subtle);
                          transition:background 0.1s;cursor:default"
                   onmouseover="this.style.background='var(--bg-card-hover)'" onmouseout="this.style.background=''">
        <div style="flex:1;min-width:0">
          <div style="font-size:12px;font-weight:500;color:var(--text-primary);line-height:1.35;word-break:break-word">${_esc(k.metric)}</div>
          <div style="font-size:10px;color:var(--text-muted);margin-top:2px">${_esc(k.section)}</div>
        </div>
        <input type="checkbox" ${isVisible?'checked':''}
               style="accent-color:var(--brand-accent);width:15px;height:15px;flex-shrink:0;cursor:pointer"
               onchange="App.addKpiToOverview('${k.id}',this.checked)">
      </div>`;
    }).join('');

    return `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
        <h2 class="section-title" style="margin:0">Overview Dashboard</h2>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          ${allKpis.length > 0 ? `
          <div style="position:relative" id="manage-kpi-wrap">
            <button onclick="App.toggleManageKpiDropdown()"
                    style="display:flex;align-items:center;gap:6px;padding:7px 14px;border-radius:8px;
                           font-size:12px;font-weight:500;border:1px solid var(--border-card);
                           background:var(--bg-card);color:var(--text-secondary);cursor:pointer;transition:all 0.15s"
                    onmouseover="this.style.color='var(--text-primary)'" onmouseout="this.style.color='var(--text-secondary)'">
              ⊞ Manage KPIs
              <span style="background:rgba(0,194,168,0.15);color:var(--brand-accent);font-size:10px;
                           padding:1px 6px;border-radius:10px;font-weight:700">${overviewKpis.length}/${allKpis.length}</span>
            </button>
            <div id="manage-kpi-dropdown" style="display:none;position:absolute;right:0;top:calc(100% + 6px);
                 width:340px;background:var(--bg-modal);border:1px solid var(--border-card);
                 border-radius:10px;box-shadow:var(--shadow-modal);z-index:150;overflow:hidden">
              <div style="padding:10px 12px;border-bottom:1px solid var(--border-subtle);
                          display:flex;align-items:center;justify-content:space-between">
                <div style="font-size:12px;font-weight:600;color:var(--text-primary)">Manage Overview KPIs</div>
                <button onclick="App.toggleManageKpiDropdown()"
                        style="color:var(--text-muted);font-size:16px;background:none;border:none;cursor:pointer;padding:0 2px">✕</button>
              </div>
              <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
                <input type="text" id="manage-kpi-search" class="input-field"
                       style="padding:6px 10px;font-size:12px" placeholder="Search KPIs…"
                       oninput="App.filterManageKpiList(this.value)">
              </div>
              <div id="manage-kpi-list" style="max-height:260px;overflow-y:auto">
                ${manageList || '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs yet</div>'}
              </div>
              <div style="padding:8px 12px;border-top:1px solid var(--border-subtle);display:flex;gap:8px">
                <button onclick="App.setAllOverviewKpis(true)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">Select All</button>
                <button onclick="App.setAllOverviewKpis(false)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">Clear All</button>
              </div>
            </div>
          </div>` : ''}
          <select class="input-field" style="width:auto;padding:6px 28px 6px 10px;font-size:12px"
                  onchange="DataStore.updateSettings({reportingPeriod:this.value})">
            <option value="monthly"   ${settings.reportingPeriod==='monthly'?'selected':''}>Monthly</option>
            <option value="quarterly" ${settings.reportingPeriod==='quarterly'?'selected':''}>Quarterly</option>
            <option value="ytd"       ${settings.reportingPeriod==='ytd'?'selected':''}>YTD</option>
            <option value="yearly"    ${settings.reportingPeriod==='yearly'?'selected':''}>Full Year</option>
            <option value="last_fy"   ${settings.reportingPeriod==='last_fy'?'selected':''}>Last FY</option>
          </select>
        </div>
      </div>`;
  }

  // ─── RAG Strip ────────────────────────────────────────────────────────────
  function _renderRagStrip(ragCounts, total) {
    return `
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px">
        ${Object.entries(ragCounts).map(([rag,cnt])=>cnt>0?`
          <div style="background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;gap:8px">
            <div class="rag-badge ${rag}">${rag.charAt(0).toUpperCase()+rag.slice(1)}</div>
            <span style="font-family:var(--font-display);font-size:20px;font-weight:700">${cnt}</span>
          </div>`:'').join('')}
        <div style="flex:1;background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;justify-content:flex-end;gap:6px">
          <span class="label-xs">Total KPIs:</span>
          <span style="font-size:14px;font-weight:700;color:var(--brand-accent)">${total}</span>
        </div>
      </div>`;
  }

  // ─── Hero Card ────────────────────────────────────────────────────────────
  function _renderHeroCard(st, heroKpis, allKpis, settings) {
    const chartType    = st.chartType || 'odometer';
    const heroPosition = st.heroPosition || 'center';
    const menuOpen     = st.menuOpen || false;
    const pickerOpen   = st.kpiPickerOpen || false;
    const heroKpiIds   = st.heroKpiIds || [];

    const chartLabels = { odometer:'Odometer', bar:'Bar Chart', line:'Line Chart', pie:'Pie Chart' };
    const chartIcons  = { odometer:'◎', bar:'▊', line:'∿', pie:'◕' };

    // ── Chart type menu ──
    const menuDropdown = menuOpen ? `
      <div style="position:absolute;right:0;top:calc(100% + 6px);
                  width:200px;background:var(--bg-modal);border:1px solid var(--border-card);
                  border-radius:10px;box-shadow:var(--shadow-modal);z-index:160;overflow:hidden">
        <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
          <div style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:var(--text-muted)">Chart Type</div>
        </div>
        ${['odometer','bar','line','pie'].map(type => `
          <div onclick="OverviewPanel.setChartType('${type}')"
               style="display:flex;align-items:center;gap:10px;padding:9px 14px;cursor:pointer;
                      background:${type===chartType?'rgba(0,194,168,0.08)':'transparent'};
                      color:${type===chartType?'var(--brand-accent)':'var(--text-secondary)'};transition:background 0.1s"
               onmouseover="this.style.background='var(--bg-card-hover)'" onmouseout="this.style.background='${type===chartType?'rgba(0,194,168,0.08)':'transparent'}'">
            <span style="font-size:16px;width:18px;text-align:center">${chartIcons[type]}</span>
            <span style="font-size:13px;font-weight:500">${chartLabels[type]}</span>
            ${type===chartType?'<span style="margin-left:auto;font-size:11px;color:var(--brand-accent)">✓</span>':''}
          </div>`).join('')}
        <div style="padding:8px 12px;border-top:1px solid var(--border-subtle)">
          <div style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:var(--text-muted);margin-bottom:6px">Position</div>
          ${['center','grid'].map(pos => `
            <div onclick="OverviewPanel.setHeroPosition('${pos}')"
                 style="display:flex;align-items:center;gap:8px;padding:6px 2px;cursor:pointer;
                        color:${pos===heroPosition?'var(--brand-accent)':'var(--text-secondary)'};font-size:12px">
              <span style="width:14px;height:14px;border-radius:50%;border:2px solid currentColor;display:inline-flex;align-items:center;justify-content:center">
                ${pos===heroPosition?'<span style="width:7px;height:7px;border-radius:50%;background:currentColor"></span>':''}
              </span>
              ${pos==='center'?'Full-width (standalone)':'In KPI grid flow'}
            </div>`).join('')}
        </div>
      </div>` : '';

    // ── KPI picker ──
    const pickerDropdown = pickerOpen ? `
      <div style="position:absolute;left:0;top:calc(100% + 6px);
                  width:320px;background:var(--bg-modal);border:1px solid var(--border-card);
                  border-radius:10px;box-shadow:var(--shadow-modal);z-index:160;overflow:hidden">
        <div style="padding:10px 12px;border-bottom:1px solid var(--border-subtle);
                    display:flex;align-items:center;justify-content:space-between">
          <div style="font-size:12px;font-weight:600;color:var(--text-primary)">Select KPIs for Hero Card</div>
          <button onclick="OverviewPanel.toggleKpiPicker()"
                  style="color:var(--text-muted);font-size:16px;background:none;border:none;cursor:pointer">✕</button>
        </div>
        <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
          <input type="text" class="input-field" style="padding:6px 10px;font-size:12px"
                 placeholder="Search KPIs…" oninput="OverviewPanel.filterHeroKpiSearch(this.value)">
        </div>
        <div id="hero-kpi-list" style="max-height:240px;overflow-y:auto">
          ${allKpis.length === 0
            ? '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs yet</div>'
            : allKpis.map(k => {
                const isSelected = heroKpiIds.includes(k.id);
                return `<div class="hero-kpi-item">
                  <div style="flex:1;min-width:0">
                    <div class="hki-name">${_esc(k.metric)}</div>
                    <div class="hki-section">${_esc(k.section)}</div>
                  </div>
                  <input type="checkbox" ${isSelected?'checked':''}
                         style="accent-color:var(--brand-accent);width:15px;height:15px;flex-shrink:0;cursor:pointer"
                         onchange="OverviewPanel.toggleHeroKpi('${k.id}',this.checked)">
                </div>`;
              }).join('')}
        </div>
        <div style="padding:8px 12px;border-top:1px solid var(--border-subtle);display:flex;gap:8px">
          <button onclick="OverviewPanel.setHeroAllKpis(true)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">All</button>
          <button onclick="OverviewPanel.setHeroAllKpis(false)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">None</button>
          <button onclick="OverviewPanel.toggleKpiPicker()" class="btn btn-primary" style="flex:1;font-size:11px;padding:5px">Done</button>
        </div>
      </div>` : '';

    const heroCount = heroKpiIds.length > 0 ? heroKpiIds.length : (DataStore.getKeyKpis().length || allKpis.length);
    const isEmpty   = allKpis.length === 0;

    const cardContent = isEmpty
      ? `<div style="text-align:center;padding:48px 24px;color:var(--text-muted)">
           <div style="font-size:36px;margin-bottom:12px">◈</div>
           <div style="font-size:15px;font-weight:600;margin-bottom:6px;color:var(--text-secondary)">No KPIs configured</div>
           <div style="font-size:13px;margin-bottom:16px">Add KPIs in Data Entry to visualise them here.</div>
           <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
         </div>`
      : `<div id="hero-chart-area" style="min-height:200px;padding:8px 0">
           <div style="text-align:center;padding:32px;color:var(--text-muted);font-size:12px">Loading chart…</div>
         </div>`;

    return `
      <div style="margin-bottom:24px">
        <!-- Hero wrapper — position drives layout -->
        <div style="position:relative;background:var(--bg-card);border:1px solid var(--border-card);
                    border-radius:16px;padding:20px 24px;box-shadow:var(--shadow-card);
                    ${heroPosition==='center' ? 'width:100%;' : ''}">

          <!-- Top bar of hero card -->
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;gap:12px;flex-wrap:wrap">

            <!-- Left: KPI picker trigger -->
            <div style="position:relative" id="hero-picker-wrap">
              <button onclick="OverviewPanel.toggleKpiPicker()"
                      style="display:flex;align-items:center;gap:7px;padding:6px 13px;border-radius:8px;
                             font-size:12px;font-weight:600;border:1px solid var(--border-card);
                             background:var(--bg-input);color:var(--text-secondary);cursor:pointer;transition:all 0.15s"
                      onmouseover="this.style.borderColor='var(--brand-accent)';this.style.color='var(--text-primary)'"
                      onmouseout="this.style.borderColor='var(--border-card)';this.style.color='var(--text-secondary)'">
                <span style="font-size:15px">◉</span>
                ${heroCount} KPI${heroCount !== 1 ? 's' : ''} selected
                <span style="font-size:9px;opacity:0.5">▾</span>
              </button>
              ${pickerDropdown}
            </div>

            <!-- Centre: chart type tabs -->
            <div style="display:flex;gap:4px;background:var(--bg-input);border-radius:8px;padding:3px;border:1px solid var(--border-subtle)">
              ${['odometer','bar','line','pie'].map(type => `
                <button class="ct-tab ${type===chartType?'active':''}"
                        onclick="OverviewPanel.setChartType('${type}')">
                  ${chartIcons[type]} ${chartLabels[type]}
                </button>`).join('')}
            </div>

            <!-- Right: ••• menu -->
            <div style="position:relative" id="hero-menu-wrap">
              <button class="hero-menu-btn" onclick="OverviewPanel.toggleMenu()"
                      style="background:var(--bg-input);border:1px solid var(--border-card);
                             border-radius:7px;padding:5px 10px;cursor:pointer;
                             color:var(--text-secondary);font-size:16px;letter-spacing:2px">
                •••
              </button>
              ${menuDropdown}
            </div>
          </div>

          <!-- Decorative accent bar -->
          <div style="height:2px;background:linear-gradient(90deg,var(--brand-accent),var(--brand-accent-2),transparent);
                      border-radius:1px;margin-bottom:20px;opacity:0.6"></div>

          <!-- Chart area -->
          ${cardContent}
        </div>
      </div>`;
  }

  // ─── Utilities ────────────────────────────────────────────────────────────
  function _esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  function _shortNum(n) {
    if (Math.abs(n) >= 1e9) return (n/1e9).toFixed(1)+'B';
    if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(1)+'M';
    if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(0)+'K';
    return Math.round(n).toString();
  }

  // ── Formula modal helper (called from formulaKpiModal inline handlers) ────
  function _fmlSelectOp(op) {
    const hidden = document.getElementById('fml-op');
    if (!hidden) return;
    const prev = hidden.value;
    hidden.value = op;

    // Update button styles
    ['sum','subtract','multiply','divide','avg','min','max','custom'].forEach(o => {
      const btn = document.getElementById('fml-op-btn-'+o);
      const chk = document.getElementById('fml-op-chk-'+o);
      if (!btn) return;
      const active = o === op;
      btn.style.borderColor  = active ? 'var(--brand-accent)' : 'var(--border-card)';
      btn.style.background   = active ? 'rgba(0,194,168,0.08)' : 'var(--bg-input)';
      btn.querySelector('div:first-child > div:first-child').style.color = active ? 'var(--brand-accent)' : 'var(--text-primary)';
      if (chk) chk.innerHTML = active ? '<div style="width:8px;height:8px;border-radius:50%;background:var(--brand-accent)"></div>' : '';
      if (chk) chk.style.borderColor = active ? 'var(--brand-accent)' : 'var(--border-card)';
    });

    // Show/hide source vs custom sections
    const srcSec = document.getElementById('fml-source-section');
    const cstSec = document.getElementById('fml-custom-section');
    if (srcSec) srcSec.style.display = op === 'custom' ? 'none' : '';
    if (cstSec) cstSec.style.display = op === 'custom' ? '' : 'none';
  }

  // ─────────────────────────────────────────────────────────────────────────
  return {
    render,
    mountCharts,
    setChartType,
    toggleMenu,
    closeMenu,
    toggleKpiPicker,
    toggleHeroKpi,
    setHeroAllKpis,
    setHeroPosition,
    filterHeroKpiSearch,
    _fmlSelectOp,
  };
})();
/**
 * pages.js v5
 * Rebuilt Data Entry:
 *  - Simple mode: compact table with inline operator + unit selectors on target cells
 *  - Advanced mode: per-KPI accordion cards — metadata collapses, months get full width
 *  - Operator choices: none, >, >=, <, <=, =, <>
 *  - Unit choices: none, $, %
 */

const Pages = (() => {

  const OPS   = ['', '>', '>=', '<', '<=', '=', '<>'];

  /**
   * Format a raw value as a readable hint for data entry inputs.
   * Unit-aware: % values show as "26%", $ values show with commas, plain numbers get commas.
   * Sized for readability inside inputs.
   */
  function _fmtRaw(val, unit) {
    if (val === null || val === undefined || val === '') return '';
    const n = parseFloat(val);
    if (isNaN(n)) return '';
    if (unit === '%') {
      // Stored as decimal (0.26) → show 26%
      const s = DataStore.getSettings();
      const pctVal = (s.pctStorage || 'decimal') === 'decimal' ? n * 100 : n;
      return pctVal % 1 === 0 ? pctVal.toFixed(0) + '%' : pctVal.toFixed(1) + '%';
    }
    if (unit === '$') {
      const s   = DataStore.getSettings();
      const sym = s.currencySymbol || '$';
      const abs = Math.abs(n);
      if (abs >= 1e9) return sym + (n / 1e9).toFixed(2) + 'B';
      if (abs >= 1e6) return sym + (n / 1e6).toFixed(2) + 'M';
      return sym + n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }
    // Plain number — just commas
    return n % 1 === 0
      ? n.toLocaleString()
      : n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }
  const UNITS = ['', '$', '%'];

  // ── OVERVIEW — delegates to OverviewPanel module ──────────────────────────
  function overview() {
    // OverviewPanel.render() returns the full HTML.
    // After the DOM is painted we mount the live chart into #hero-chart-area.
    const html = OverviewPanel.render();
    // Schedule chart mount after the current render cycle writes to DOM
    setTimeout(() => OverviewPanel.mountCharts(), 0);
    return html;
  }

  // ── SECTION PAGE ──────────────────────────────────────────────────────────
  function sectionPage(sectionName) {
    const kpis = DataStore.getKpisBySection(sectionName);
    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:10px">
          <div>
            <div class="label-xs" style="margin-bottom:4px;cursor:pointer;color:var(--brand-accent)" onclick="App.navigate('overview')">← Overview</div>
            <h2 class="section-title" style="margin:0">${sectionName}</h2>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn btn-ghost" style="font-size:12px" onclick="App.promptRenameSection('${encodeURIComponent(sectionName)}')">✎ Rename</button>
            <button class="btn btn-ghost" style="font-size:12px;color:var(--rag-red)" onclick="App.confirmRemoveSection('${encodeURIComponent(sectionName)}')"><i class="fa-solid fa-trash-can"></i> Remove</button>
          </div>
        </div>
        ${Components.ragSummaryBar(kpis)}
        ${kpis.length>0?`
          <div class="grid-auto">${kpis.map(kpi=>Components.kpiCard(kpi, DataStore.getPeriodStats(kpi, DataStore.getSettings().reportingPeriod||'monthly'), false)).join('')}</div>`:`
          <div class="card" style="text-align:center;padding:40px">
            <div style="font-size:28px;margin-bottom:10px"><i class="fa-solid fa-clipboard-list"></i></div>
            <div style="font-size:15px;font-weight:600;margin-bottom:6px">No KPIs in this section</div>
            <div style="font-size:13px;color:var(--text-secondary);margin-bottom:14px">Add KPIs from Data Entry and assign them to this section.</div>
            <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
          </div>`}
      </div>`;
  }

  // ── DATA ENTRY ────────────────────────────────────────────────────────────
  function dataEntry(activeTab, advancedMode, expandedKpiId) {
    activeTab    = activeTab    || 'kpis';
    advancedMode = !!advancedMode;

    const tabBtn = (id, label, active) =>
      `<button onclick="App.setDataEntryTab('${id}')"
               style="padding:8px 18px;font-size:13px;font-weight:500;border:none;cursor:pointer;border-radius:8px;
                      background:${active?'var(--brand-accent)':'transparent'};
                      color:${active?'var(--brand-primary)':'var(--text-secondary)'};transition:all 0.15s">${label}</button>`;

    const content = activeTab==='kpis'
      ? _renderKpiTab(advancedMode, expandedKpiId)
      : _renderThresholdsTab();

    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
          <h2 class="section-title" style="margin:0">Data Entry</h2>
          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" onclick="App.openAddKpi()">+ New KPI</button>
            <button class="btn btn-ghost" onclick="App.openAddSectionModal()">+ New Section</button>
          </div>
        </div>

        <div style="display:flex;gap:4px;background:var(--bg-card);border:1px solid var(--border-card);border-radius:10px;padding:4px;width:fit-content;margin-bottom:24px">
          ${tabBtn('kpis','KPI Management',activeTab==='kpis')}
          ${tabBtn('thresholds','Colour Rules (RAG)',activeTab==='thresholds')}
        </div>

        ${content}
      </div>`;
  }

  // ── KPI TAB ───────────────────────────────────────────────────────────────
  function _renderKpiTab(advancedMode, expandedKpiId) {
    const allKpis    = DataStore.getKpis();
    const sections   = DataStore.getSections();
    const thresholds = DataStore.getThresholds();
    const fyMonths   = DataStore.getFyMonths();

    if (allKpis.length === 0) {
      return `
        <div class="card" style="text-align:center;padding:56px 40px;border:2px dashed var(--border-card)">
          <div style="font-size:40px;margin-bottom:14px"><i class="fa-solid fa-clipboard-list"></i></div>
          <div style="font-family:var(--font-display);font-size:18px;font-weight:700;margin-bottom:8px">No KPIs yet</div>
          <div style="font-size:13px;color:var(--text-secondary);max-width:360px;margin:0 auto 20px">
            Start by creating a section, then add your KPIs. You can also bulk-import from XLSX or CSV.
          </div>
          <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
            <button class="btn btn-primary" onclick="App.openAddKpi()">+ Add First KPI</button>
            <button class="btn btn-ghost" onclick="App.openAddSectionModal()">+ New Section</button>
          </div>
          <div style="margin-top:28px">${Components.xlsxImportPanel()}</div>
        </div>`;
    }

    return `
      <!-- Bulk Import -->
      <div style="margin-bottom:20px">
        <details style="background:var(--bg-card);border:1px solid var(--border-card);border-radius:10px;overflow:hidden">
          <summary style="padding:12px 16px;cursor:pointer;font-size:13px;font-weight:500;color:var(--text-secondary);list-style:none;display:flex;align-items:center;gap:8px">
            <i class="fa-solid fa-chart-bar"></i> Bulk Import via XLSX / CSV
            <span style="margin-left:auto;font-size:11px;color:var(--text-muted)">click to expand ▾</span>
          </summary>
          <div style="padding:16px;border-top:1px solid var(--border-subtle)">${Components.xlsxImportPanel()}</div>
        </details>
      </div>

      <!-- Mode toggle + search -->
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
        <div style="display:flex;align-items:center;gap:10px;flex:1;min-width:0">
          ${advancedMode ? `
            <div style="position:relative;flex:1;max-width:280px">
              <input type="text" id="kpi-search-adv" class="input-field"
                     style="padding:7px 12px 7px 32px;font-size:12px"
                     placeholder="Search KPIs…"
                     value="${App._kpiSearchQuery||''}"
                     oninput="App.setKpiSearch(this.value)">
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none"><i class="fa-solid fa-magnifying-glass"></i></span>
            </div>` : `
            <div style="position:relative;flex:1;max-width:280px">
              <input type="text" id="kpi-search-simple" class="input-field"
                     style="padding:7px 12px 7px 32px;font-size:12px"
                     placeholder="Search KPIs…"
                     value="${App._simpleSearchQuery||''}"
                     oninput="App.setSimpleKpiSearch(this.value)">
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none"><i class="fa-solid fa-magnifying-glass"></i></span>
            </div>`}
        </div>
        <button onclick="App.toggleAdvancedEntry()"
                style="display:flex;align-items:center;gap:7px;padding:7px 16px;border-radius:8px;font-size:12px;font-weight:600;
                       border:1px solid ${advancedMode?'var(--brand-accent)':'var(--border-card)'};
                       background:${advancedMode?'rgba(0,194,168,0.12)':'transparent'};
                       color:${advancedMode?'var(--brand-accent)':'var(--text-secondary)'};cursor:pointer;white-space:nowrap">
          ${advancedMode ? '✕ Exit Monthly Entry' : '⊞ Monthly Data Entry'}
        </button>
      </div>

      ${advancedMode ? `
        <!-- Sticky legend -->
        <div style="position:sticky;top:var(--header-height);z-index:80;
                    background:rgba(9,25,41,0.95);border-bottom:2px solid var(--brand-accent);
                    margin-left:-24px;margin-right:-24px;padding:9px 24px;
                    margin-bottom:18px;display:flex;align-items:center;gap:14px;
                    backdrop-filter:blur(10px)">
          <span style="font-size:13px;font-weight:700;color:var(--brand-accent)">⊞ Monthly Entry</span>
          <span style="font-size:11px;color:var(--text-secondary)">Click a KPI header to collapse its fields · Teal inputs = data entered · YTD auto-sums monthly actuals</span>
          <button onclick="App.toggleAdvancedEntry()"
                  style="margin-left:auto;padding:5px 14px;border-radius:7px;font-size:11px;font-weight:700;
                         border:1px solid var(--brand-accent);background:transparent;
                         color:var(--brand-accent);cursor:pointer">✕ Exit</button>
        </div>` : ''}

      <!-- Sections -->
      <div id="kpi-entry-tables">
        ${sections.map(section => {
          const sKpis = allKpis.filter(k=>k.section===section);
          return advancedMode
            ? _buildAdvancedSection(section, sKpis, thresholds, fyMonths)
            : _buildSimpleSection(section, sKpis, thresholds, expandedKpiId, fyMonths);
        }).join('')}
      </div>

      <datalist id="who-list-de">
        ${[...new Set(allKpis.map(k=>k.who).filter(Boolean))].map(w=>`<option value="${w}">`).join('')}
      </datalist>

      <style>
        /* ── Spacer insert bar ── */
        tr.insert-spacer td { padding:0;height:3px;background:transparent;position:relative;cursor:pointer;transition:all 0.12s; }
        tr.insert-spacer:hover td { height:20px;background:rgba(0,194,168,0.06);border-top:2px solid var(--brand-accent);border-bottom:2px solid var(--brand-accent); }
        tr.insert-spacer td .spacer-label { display:none;position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);font-size:10px;font-weight:700;color:var(--brand-accent);white-space:nowrap;pointer-events:none; }
        tr.insert-spacer:hover td .spacer-label { display:block; }

        /* ── Compound target cell: [op][unit][value] ── */
        .tc { display:flex;align-items:stretch;border:1px solid var(--border-card);border-radius:7px;overflow:hidden;background:var(--bg-input); }
        .tc select { border:none;background:var(--bg-card);color:var(--brand-accent);font-family:var(--font-mono);font-size:11px;font-weight:700;padding:0 3px 0 5px;cursor:pointer;outline:none;flex-shrink:0;border-right:1px solid var(--border-subtle); }
        .tc select.unit-sel { border-right:none;border-left:1px solid var(--border-subtle);color:var(--text-secondary); }
        .tc input[type=number] { border:none;background:transparent;color:var(--text-primary);font-size:11px;padding:5px 6px;flex:1;min-width:48px;outline:none;width:100%; }
        .tc input[type=number]::-webkit-inner-spin-button { opacity:0.3; }

        /* ── Advanced mode: KPI accordion card ── */
        .adv-card { background:var(--bg-card);border:1px solid var(--border-card);border-radius:10px;margin-bottom:8px;overflow:hidden; }
        .adv-hdr  { display:flex;align-items:center;gap:10px;padding:11px 14px;cursor:pointer;user-select:none; }
        .adv-hdr:hover { background:rgba(255,255,255,0.02); }
        .adv-meta { display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px;padding:12px 14px;border-top:1px solid var(--border-subtle); }
        .adv-meta.hidden { display:none; }
        .adv-fld  { display:flex;flex-direction:column;gap:4px; }
        .adv-fld label { font-size:10px;font-weight:700;color:var(--text-muted);letter-spacing:0.07em;text-transform:uppercase; }
        .adv-months-wrap { padding:12px 14px 14px;border-top:1px solid var(--border-subtle);background:rgba(0,194,168,0.025); }
        .month-grid { display:grid;grid-template-columns:repeat(6,1fr);gap:6px;margin-top:8px; }
        @media(max-width:860px){ .month-grid{grid-template-columns:repeat(4,1fr);} }
        @media(max-width:560px){ .month-grid{grid-template-columns:repeat(3,1fr);} }
        .mc { display:flex;flex-direction:column;gap:3px; }
        .mc label { font-size:10px;font-weight:700;color:var(--text-muted);letter-spacing:0.05em;text-align:center; }
        .mc input { padding:6px 4px;font-size:12px;text-align:center;border-radius:6px;border:1px solid var(--border-card);background:var(--bg-input);color:var(--text-secondary);width:100%;outline:none;transition:border-color 0.15s,background 0.15s; }
        .mc input:focus { border-color:var(--brand-accent);background:rgba(0,194,168,0.04);color:var(--text-primary); }
        .mc input.filled { border-color:rgba(0,194,168,0.45);background:rgba(0,194,168,0.06);color:var(--text-primary); }
      </style>
      <script>
        // Re-apply simple search filter after re-render
        (function(){
          var q = (App._simpleSearchQuery||'').toLowerCase().trim();
          if (!q) return;
          setTimeout(function(){ App.setSimpleKpiSearch(App._simpleSearchQuery||''); }, 0);
        })();
      </script>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SIMPLE MODE — compact table with compound target cells
  // ══════════════════════════════════════════════════════════════════════════
  function _buildSimpleSection(section, sKpis, thresholds, expandedKpiId, fyMonths) {
    const enc = encodeURIComponent(section);
    return `
      <div style="margin-bottom:24px">
        ${_sectionHeader(section, enc)}
        <div style="overflow-x:auto;border:1px solid var(--border-card);border-radius:10px;background:var(--bg-card)">
          <table style="width:100%;border-collapse:collapse;font-size:12px;min-width:740px" data-section-table>
            <thead>
              <tr style="background:var(--bg-input);color:var(--text-muted)">
                <th style="padding:8px 6px;text-align:center;width:40px" title="Show on Overview page">Ovw</th>
                <th style="padding:8px;text-align:left">KPI Name</th>
                <th style="padding:8px;text-align:left;width:78px">Who</th>
                <th style="padding:8px;text-align:left;width:182px">
                  Target FY
                  <span style="display:block;font-size:9px;font-weight:400;color:var(--text-muted);margin-top:1px">op · unit · value</span>
                </th>
                <th style="padding:8px;text-align:left;width:170px">Target / Month</th>
                <th style="padding:8px;text-align:left;width:120px">YTD Actual</th>
                <th style="padding:8px;text-align:left;width:138px">Colour Rule</th>
                <th style="padding:8px;text-align:center;width:46px">RAG</th>
                <th style="width:28px"></th>
              </tr>
            </thead>
            <tbody>
              ${sKpis.map((kpi,i)=>_simpleRow(kpi,i,thresholds,expandedKpiId,fyMonths)).join('')}
              ${sKpis.length===0?`<tr><td colspan="9" style="padding:20px;text-align:center;color:var(--text-muted);font-style:italic">No KPIs yet — add one above.</td></tr>`:''}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  function _simpleRow(kpi, i, thresholds, expandedKpiId, fyMonths) {
    const rag = kpi.rag||'neutral';
    const isExpanded = expandedKpiId === kpi.id;
    const ma = kpi.monthlyActuals || {};
    const spacer = `
      <tr class="insert-spacer" onclick="App.promptInsertSection('${kpi.id}')">
        <td colspan="9"><span class="spacer-label">＋ Insert section break above "${_esc(kpi.metric)}"</span></td>
      </tr>`;
    const row = `
      <tr data-kpi-id="${kpi.id}" style="border-top:1px solid var(--border-subtle);${i%2?'background:rgba(255,255,255,0.012)':''}${isExpanded?';background:rgba(0,194,168,0.04)!important;outline:1px solid rgba(0,194,168,0.25);outline-offset:-1px':''}">
        <td style="padding:6px;text-align:center">
          <input type="checkbox" ${kpi.isKey?'checked':''} title="Show on Overview page"
                 style="width:14px;height:14px;accent-color:var(--brand-accent);cursor:pointer"
                 onchange="App.addKpiToOverview('${kpi.id}',this.checked)">
        </td>
        <td style="padding:6px 8px">
          <div style="display:flex;align-items:center;gap:6px">
            <button onclick="App.toggleKpiMonthly('${kpi.id}')"
                    title="${isExpanded?'Collapse monthly entry':'Expand monthly entry for this KPI'}"
                    style="background:${isExpanded?'rgba(0,194,168,0.18)':'rgba(255,255,255,0.04)'};
                           border:1px solid ${isExpanded?'var(--brand-accent)':'var(--border-card)'};
                           color:${isExpanded?'var(--brand-accent)':'var(--text-muted)'};
                           border-radius:4px;cursor:pointer;font-size:11px;padding:2px 6px;
                           flex-shrink:0;transition:all 0.15s;font-family:var(--font-mono)"
                    >⊞</button>
            <div>
              <div style="font-weight:500;color:var(--text-primary)">${_esc(kpi.metric)}</div>
              ${kpi.comment?`<div style="font-size:10px;color:var(--rag-amber);margin-top:1px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${_esc(kpi.comment)}</div>`:''}
            </div>
          </div>
        </td>
        <td style="padding:6px 8px">
          <input type="text" class="input-field" style="padding:4px 5px;font-size:11px;width:68px"
                 value="${_esc(kpi.who||'')}" list="who-list-de"
                 onchange="App.quickUpdate('${kpi.id}','who',this.value)">
        </td>
        <td style="padding:6px 8px">${_tc(kpi,'targetFY26')}</td>
        <td style="padding:6px 8px">${_tc(kpi,'targetMonth')}</td>
        <td style="padding:6px 8px">
          <div style="position:relative">
            <input type="number" class="input-field" style="padding:4px 5px 12px;font-size:11px;width:106px"
                   value="${kpi.ytd??''}" placeholder="YTD"
                   title="Formatted: ${DataStore.formatValue(kpi.ytd,kpi)}"
                   onchange="App.quickUpdate('${kpi.id}','ytd',this.value)">
            ${kpi.ytd!==null&&kpi.ytd!==undefined?`<span style="position:absolute;bottom:2px;right:6px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.ytd, kpi.unit)}</span>`:''}
          </div>
        </td>
        <td style="padding:6px 8px">
          <select class="input-field" style="padding:4px 5px;font-size:11px;width:130px"
                  title="Controls RAG colour — separate from target"
                  onchange="App.quickUpdate('${kpi.id}','thresholdId',this.value)">
            ${thresholds.map(th=>`<option value="${th.id}" ${kpi.thresholdId===th.id?'selected':''}>${th.name.length>22?th.name.slice(0,20)+'…':th.name}</option>`).join('')}
          </select>
        </td>
        <td style="padding:6px;text-align:center">${Components.ragDot(rag)}</td>
        <td style="padding:4px 2px;text-align:center;white-space:nowrap">
          <button onclick="App.openEditKpi('${kpi.id}')" title="Full edit"
                  style="color:var(--text-muted);background:none;border:none;cursor:pointer;font-size:13px;padding:2px">✎</button>
          <button onclick="App.confirmRemoveKpi('${kpi.id}')" title="Remove"
                  style="color:var(--rag-red);background:none;border:none;cursor:pointer;font-size:13px;padding:2px">✕</button>
        </td>
      </tr>`;

    // Inline full expand row (metadata + month grid)
    const filled   = fyMonths ? fyMonths.filter(m=>ma[m]!==undefined&&ma[m]!==null).length : 0;
    const op       = kpi.targetFY26Op||'';
    const unit     = kpi.unit||'';
    const opOpts   = OPS.map(o=>`<option value="${_esc(o)}" ${op===o?'selected':''}>${o||'—'}</option>`).join('');
    const unitOpts = UNITS.map(u=>`<option value="${u}" ${unit===u?'selected':''}>${u||'#'}</option>`).join('');
    const monthRow = isExpanded && fyMonths ? `
      <tr style="border-top:1px solid rgba(0,194,168,0.2)">
        <td colspan="9" style="padding:0;background:rgba(0,194,168,0.03)">
          <div style="padding:14px 16px 16px;display:flex;flex-direction:column;gap:14px">

            <!-- Header bar -->
            <div style="display:flex;align-items:center;justify-content:space-between">
              <span style="font-size:11px;font-weight:700;color:var(--brand-accent);letter-spacing:0.07em;text-transform:uppercase">
                ⊞ ${_esc(kpi.metric)}
              </span>
              <button onclick="App.toggleKpiMonthly('${kpi.id}')"
                      style="padding:3px 10px;border-radius:5px;font-size:11px;font-weight:600;
                             border:1px solid var(--brand-accent);background:transparent;
                             color:var(--brand-accent);cursor:pointer">✕ Close</button>
            </div>

            <!-- Metadata grid -->
            <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px">
              <div class="adv-fld">
                <label>KPI Name</label>
                <input type="text" class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                       value="${_esc(kpi.metric)}"
                       onchange="App.quickUpdate('${kpi.id}','metric',this.value)">
              </div>
              <div class="adv-fld">
                <label>Who / Owner</label>
                <input type="text" class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                       value="${_esc(kpi.who||'')}" list="who-list-de"
                       onchange="App.quickUpdate('${kpi.id}','who',this.value)">
              </div>
              <div class="adv-fld">
                <label>Target FY <span style="font-weight:400;color:var(--text-muted);text-transform:none;letter-spacing:0">op · unit · value</span></label>
                <div class="tc">
                  <select title="Operator" onchange="App.quickUpdate('${kpi.id}','targetFY26Op',this.value||null)">${opOpts}</select>
                  <select class="unit-sel" title="Unit" onchange="App.quickUpdate('${kpi.id}','unit',this.value)">${unitOpts}</select>
                  <div style="position:relative;flex:1">
                    <input type="number" value="${kpi.targetFY26??''}" placeholder="Annual"
                           style="padding:5px 6px ${kpi.targetFY26!==null&&kpi.targetFY26!==undefined?'16px':'5px'} 6px;width:100%"
                           onchange="App.quickUpdate('${kpi.id}','targetFY26',this.value)">
                    ${kpi.targetFY26!==null&&kpi.targetFY26!==undefined?`<span style="position:absolute;bottom:2px;right:5px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.targetFY26,kpi.unit)}</span>`:''}
                  </div>
                </div>
              </div>
              <div class="adv-fld">
                <label>Target / Month</label>
                <div class="tc">
                  <select title="Operator" onchange="App.quickUpdate('${kpi.id}','targetFY26Op',this.value||null)">${opOpts}</select>
                  <select class="unit-sel" title="Unit" onchange="App.quickUpdate('${kpi.id}','unit',this.value)">${unitOpts}</select>
                  <div style="position:relative;flex:1">
                    <input type="number" value="${kpi.targetMonth??''}" placeholder="Monthly"
                           style="padding:5px 6px ${kpi.targetMonth!==null&&kpi.targetMonth!==undefined?'16px':'5px'} 6px;width:100%"
                           onchange="App.quickUpdate('${kpi.id}','targetMonth',this.value)">
                    ${kpi.targetMonth!==null&&kpi.targetMonth!==undefined?`<span style="position:absolute;bottom:2px;right:5px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.targetMonth,kpi.unit)}</span>`:''}
                  </div>
                </div>
              </div>
              <div class="adv-fld">
                <label>YTD Method <span style="font-weight:400;color:var(--text-muted);text-transform:none;letter-spacing:0">— how months combine</span></label>
                <select class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                        onchange="App.quickUpdate('${kpi.id}','ytdMethod',this.value)">
                  <option value="sum"  ${(kpi.ytdMethod||'sum')==='sum' ?'selected':''}>∑ Sum (revenue, volume)</option>
                  <option value="avg"  ${kpi.ytdMethod==='avg' ?'selected':''}>x̄ Average (%, rates, scores)</option>
                  <option value="last" ${kpi.ytdMethod==='last'?'selected':''}>→ Latest month (point-in-time)</option>
                  <option value="max"  ${kpi.ytdMethod==='max' ?'selected':''}>↑ Highest month (peak)</option>
                  <option value="min"  ${kpi.ytdMethod==='min' ?'selected':''}>↓ Lowest month (best low)</option>
                </select>
              </div>
              <div class="adv-fld">
                <label>YTD Actual <span style="font-weight:400;color:var(--text-muted);text-transform:none;letter-spacing:0">(${{sum:'auto-sums',avg:'auto-averages',last:'latest month',max:'highest month',min:'lowest month'}[kpi.ytdMethod||'sum']})</span></label>
                <div style="position:relative">
                  <input type="number" class="input-field" style="padding:5px 7px 13px;font-size:12px;width:100%"
                         value="${kpi.ytd??''}" placeholder="or override manually"
                         onchange="App.quickUpdate('${kpi.id}','ytd',this.value)">
                  ${kpi.ytd!==null&&kpi.ytd!==undefined?`<span style="position:absolute;bottom:3px;right:8px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.ytd,kpi.unit)}</span>`:''}
                </div>
              </div>
              <div class="adv-fld">
                <label>Colour Rule</label>
                <select class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                        onchange="App.quickUpdate('${kpi.id}','thresholdId',this.value)">
                  ${thresholds.map(th=>`<option value="${th.id}" ${kpi.thresholdId===th.id?'selected':''}>${th.name}</option>`).join('')}
                </select>
              </div>
              <div class="adv-fld">
                <label>RAG Override</label>
                <select class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                        onchange="App.quickUpdate('${kpi.id}','ragOverride',this.value||null)">
                  <option value="" ${!kpi.ragOverride?'selected':''}>Auto</option>
                  <option value="green"   ${kpi.ragOverride==='green'  ?'selected':''}>● Green</option>
                  <option value="amber"   ${kpi.ragOverride==='amber'  ?'selected':''}>▲ Amber</option>
                  <option value="red"     ${kpi.ragOverride==='red'    ?'selected':''}>✕ Red</option>
                  <option value="neutral" ${kpi.ragOverride==='neutral'?'selected':''}>○ N/A</option>
                </select>
              </div>
              <div class="adv-fld" style="grid-column:1/-1">
                <label>Comment / Note</label>
                <input type="text" class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                       value="${_esc(kpi.comment||'')}" placeholder="Optional note…"
                       onchange="App.quickUpdate('${kpi.id}','comment',this.value)">
              </div>
              <div style="grid-column:1/-1;display:flex;gap:8px;padding-top:2px">
                <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px" onclick="App.openEditKpi('${kpi.id}')">✎ Full Edit</button>
                <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px;color:var(--rag-red)" onclick="App.confirmRemoveKpi('${kpi.id}')">✕ Remove</button>
              </div>
            </div>

            <!-- Month grid -->
            <div>
              <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
                <span style="font-size:10px;font-weight:700;color:var(--text-muted);letter-spacing:0.07em;text-transform:uppercase">
                  Monthly Actuals
                  <span style="font-weight:400;color:var(--brand-accent);margin-left:6px;text-transform:none;letter-spacing:0">
                    · ${{sum:'summing',avg:'averaging',last:'latest month',max:'peak',min:'lowest'}[kpi.ytdMethod||'sum']}
                  </span>
                </span>
                ${filled>0
                  ? `<span style="font-size:11px;color:var(--brand-accent)">${filled}/${fyMonths.length} entered · YTD: <strong>${DataStore.formatValue(kpi.ytd,kpi)}</strong></span>`
                  : `<span style="font-size:11px;color:var(--text-muted)">No months entered yet</span>`}
              </div>
              <div class="month-grid">
                ${fyMonths.map(m=>{
                  const hasVal = ma[m]!==undefined && ma[m]!==null;
                  return `
                    <div class="mc">
                      <label>${m}</label>
                      <div style="position:relative">
                        <input type="number" class="${hasVal?'filled':''}"
                               value="${hasVal?ma[m]:''}" placeholder="—"
                               style="padding:6px 4px ${hasVal?'16px':'6px'};width:100%"
                               onfocus="this.select()"
                               onchange="App.updateMonthlyActual('${kpi.id}','${m}',this.value)"
                               title="${m} actual for ${_esc(kpi.metric)}">
                        ${hasVal?`<span style="position:absolute;bottom:2px;left:50%;transform:translateX(-50%);font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none;white-space:nowrap">${_fmtRaw(ma[m],kpi.unit)}</span>`:''}
                      </div>
                    </div>`;
                }).join('')}
              </div>
            </div>

          </div>
        </td>
      </tr>` : '';

    return spacer + row + monthRow;
  }

  // Compound target cell [op ▾][unit ▾][ number ]
  // Both FY and Monthly share the same op+unit (stored on kpi.targetFY26Op / kpi.unit)
  function _tc(kpi, field) {
    const val  = kpi[field];
    const op   = kpi.targetFY26Op || '';
    const unit = kpi.unit || '';
    const opOpts   = OPS.map(o=>`<option value="${_esc(o)}" ${op===o?'selected':''}>${o||'—'}</option>`).join('');
    const unitOpts = UNITS.map(u=>`<option value="${u}" ${unit===u?'selected':''}>${u||'#'}</option>`).join('');
    const hint     = (val !== null && val !== undefined && val !== '') ? `<span style="font-size:10px;color:var(--brand-accent);position:absolute;bottom:2px;right:5px;pointer-events:none;letter-spacing:0;font-weight:500">${_fmtRaw(val, unit)}</span>` : '';
    return `
      <div class="tc" style="position:relative">
        <select title="Comparison operator" onchange="App.quickUpdate('${kpi.id}','targetFY26Op',this.value||null)">${opOpts}</select>
        <select class="unit-sel" title="Unit: # = number, $ = dollars, % = percent" onchange="App.quickUpdate('${kpi.id}','unit',this.value)">${unitOpts}</select>
        <div style="position:relative;flex:1;display:flex;flex-direction:column;justify-content:center">
          <input type="number" value="${val??''}" placeholder="—" style="padding:5px 6px ${val?'14px':'5px'} 6px"
                 onchange="App.quickUpdate('${kpi.id}','${field}',this.value)">
          ${hint}
        </div>
      </div>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  ADVANCED MODE — per-KPI accordion cards with month grid
  // ══════════════════════════════════════════════════════════════════════════
  function _buildAdvancedSection(section, sKpis, thresholds, fyMonths) {
    const enc = encodeURIComponent(section);
    // Filter by search query if one is active
    const q = (App._kpiSearchQuery||'').toLowerCase().trim();
    const filtered = q ? sKpis.filter(k => k.metric.toLowerCase().includes(q) || (k.who||'').toLowerCase().includes(q)) : sKpis;
    if (filtered.length === 0 && q) return ''; // hide section entirely if no match
    return `
      <div style="margin-bottom:24px">
        ${_sectionHeader(section, enc)}
        ${filtered.length===0
          ? `<div style="color:var(--text-muted);font-size:13px;font-style:italic;padding:10px 4px">No KPIs in this section yet.</div>`
          : filtered.map(kpi=>_advCard(kpi,thresholds,fyMonths)).join('')}
      </div>`;
  }

  function _advCard(kpi, thresholds, fyMonths) {
    const rag    = kpi.rag||'neutral';
    const ma     = kpi.monthlyActuals||{};
    const filled = fyMonths.filter(m=>ma[m]!==undefined&&ma[m]!==null).length;
    const metaId = 'meta-'+kpi.id;
    const arrId  = 'arr-'+kpi.id;

    const ragCol = {green:'var(--rag-green)',amber:'var(--rag-amber)',red:'var(--rag-red)',neutral:'var(--rag-neutral)'}[rag];
    const ragIco = {green:'●',amber:'▲',red:'✕',neutral:'○'}[rag];

    const op   = kpi.targetFY26Op||'';
    const unit = kpi.unit||'';
    const opOpts   = OPS.map(o=>`<option value="${_esc(o)}" ${op===o?'selected':''}>${o||'—'}</option>`).join('');
    const unitOpts = UNITS.map(u=>`<option value="${u}" ${unit===u?'selected':''}>${u||'#'}</option>`).join('');

    return `
      <div class="adv-card" style="border-left:3px solid ${ragCol}">

        <!-- Header (always visible) — click to toggle meta -->
        <div class="adv-hdr" onclick="(function(){
          var m=document.getElementById('${metaId}');
          var a=document.getElementById('${arrId}');
          var open=!m.classList.contains('hidden');
          m.classList.toggle('hidden',open);
          a.textContent=open?'▸':'▾';
        })()">
          <span style="color:${ragCol};font-size:15px;flex-shrink:0">${ragIco}</span>
          <div style="flex:1;min-width:0">
            <div style="font-weight:600;font-size:13px;color:var(--text-primary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${_esc(kpi.metric)}</div>
            <div style="font-size:10px;color:var(--text-muted);margin-top:1px">${_esc(kpi.section)}${kpi.who?' · '+_esc(kpi.who):''}</div>
          </div>
          <!-- Target pill -->
          <code style="font-size:12px;font-weight:700;color:var(--brand-accent);background:rgba(0,194,168,0.09);
                       padding:3px 9px;border-radius:5px;white-space:nowrap;flex-shrink:0">
            ${_esc(DataStore.formatTarget(kpi)||'no target')}
          </code>
          <!-- YTD -->
          <div style="font-size:11px;color:var(--text-secondary);white-space:nowrap;flex-shrink:0">
            YTD <strong style="color:var(--text-primary)">${DataStore.formatValue(kpi.ytd,kpi)||'—'}</strong>
          </div>
          <!-- Months filled -->
          <div style="font-size:10px;white-space:nowrap;flex-shrink:0;
                      color:${filled>0?'var(--brand-accent)':'var(--text-muted)'}">
            ${filled}/${fyMonths.length} months
          </div>
          <!-- Overview checkbox -->
          <input type="checkbox" ${kpi.isKey?'checked':''} title="Show on Overview page"
                 style="width:14px;height:14px;accent-color:var(--brand-accent);cursor:pointer;flex-shrink:0"
                 onclick="event.stopPropagation()"
                 onchange="App.addKpiToOverview('${kpi.id}',this.checked)">
          <!-- Expand arrow -->
          <span id="${arrId}" style="color:var(--text-muted);font-size:11px;flex-shrink:0;width:12px;text-align:center">▾</span>
        </div>

        <!-- Collapsible metadata grid -->
        <div class="adv-meta" id="${metaId}">
          <div class="adv-fld">
            <label>KPI Name</label>
            <input type="text" class="input-field" style="padding:5px 7px;font-size:12px"
                   value="${_esc(kpi.metric)}"
                   onchange="App.quickUpdate('${kpi.id}','metric',this.value)">
          </div>
          <div class="adv-fld">
            <label>Who / Owner</label>
            <input type="text" class="input-field" style="padding:5px 7px;font-size:12px"
                   value="${_esc(kpi.who||'')}" list="who-list-de"
                   onchange="App.quickUpdate('${kpi.id}','who',this.value)">
          </div>
          <div class="adv-fld">
            <label>Target FY &nbsp;<span style="color:var(--text-muted);font-weight:400;text-transform:none;letter-spacing:0">op · unit · value</span></label>
            <div class="tc">
              <select title="Operator" onchange="App.quickUpdate('${kpi.id}','targetFY26Op',this.value||null)">${opOpts}</select>
              <select class="unit-sel" title="Unit" onchange="App.quickUpdate('${kpi.id}','unit',this.value)">${unitOpts}</select>
              <div style="position:relative;flex:1;display:flex;flex-direction:column;justify-content:center">
                <input type="number" value="${kpi.targetFY26??''}" placeholder="Annual"
                       style="padding:5px 6px ${kpi.targetFY26!==null&&kpi.targetFY26!==undefined?'16px':'5px'} 6px"
                       onchange="App.quickUpdate('${kpi.id}','targetFY26',this.value)">
                ${kpi.targetFY26!==null&&kpi.targetFY26!==undefined?`<span style="position:absolute;bottom:2px;right:5px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.targetFY26, kpi.unit)}</span>`:''}
              </div>
            </div>
          </div>
          <div class="adv-fld">
            <label>Target / Month</label>
            <div class="tc">
              <select title="Operator" onchange="App.quickUpdate('${kpi.id}','targetFY26Op',this.value||null)">${opOpts}</select>
              <select class="unit-sel" title="Unit" onchange="App.quickUpdate('${kpi.id}','unit',this.value)">${unitOpts}</select>
              <div style="position:relative;flex:1;display:flex;flex-direction:column;justify-content:center">
                <input type="number" value="${kpi.targetMonth??''}" placeholder="Monthly"
                       style="padding:5px 6px ${kpi.targetMonth!==null&&kpi.targetMonth!==undefined?'16px':'5px'} 6px"
                       onchange="App.quickUpdate('${kpi.id}','targetMonth',this.value)">
                ${kpi.targetMonth!==null&&kpi.targetMonth!==undefined?`<span style="position:absolute;bottom:2px;right:5px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.targetMonth, kpi.unit)}</span>`:''}
              </div>
            </div>
          </div>
          <div class="adv-fld">
            <label>YTD Method
              <span style="font-weight:400;color:var(--text-muted);text-transform:none;letter-spacing:0">
                — how months combine
              </span>
            </label>
            <select class="input-field" style="padding:5px 7px;font-size:12px"
                    onchange="App.quickUpdate('${kpi.id}','ytdMethod',this.value)">
              <option value="sum"  ${(kpi.ytdMethod||'sum')==='sum' ?'selected':''}>∑ Sum (revenue, volume)</option>
              <option value="avg"  ${kpi.ytdMethod==='avg' ?'selected':''}>x̄ Average (%, rates, scores)</option>
              <option value="last" ${kpi.ytdMethod==='last'?'selected':''}>→ Latest month (point-in-time)</option>
              <option value="max"  ${kpi.ytdMethod==='max' ?'selected':''}>↑ Highest month (peak)</option>
              <option value="min"  ${kpi.ytdMethod==='min' ?'selected':''}>↓ Lowest month (best low)</option>
            </select>
          </div>
          <div class="adv-fld">
            <label>YTD Actual <span style="font-weight:400;color:var(--text-muted)">(${{sum:'auto-sums',avg:'auto-averages',last:'latest month',max:'highest month',min:'lowest month'}[kpi.ytdMethod||'sum']})</span></label>
            <div style="position:relative">
              <input type="number" class="input-field" style="padding:5px 7px 13px;font-size:12px;width:100%"
                     value="${kpi.ytd??''}" placeholder="or override manually"
                     onchange="App.quickUpdate('${kpi.id}','ytd',this.value)">
              ${kpi.ytd!==null&&kpi.ytd!==undefined?`<span style="position:absolute;bottom:3px;right:8px;font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none">${_fmtRaw(kpi.ytd, kpi.unit)}</span>`:''}
            </div>
          </div>
          <div class="adv-fld">
            <label>Colour Rule</label>
            <select class="input-field" style="padding:5px 7px;font-size:12px"
                    onchange="App.quickUpdate('${kpi.id}','thresholdId',this.value)">
              ${thresholds.map(th=>`<option value="${th.id}" ${kpi.thresholdId===th.id?'selected':''}>${th.name}</option>`).join('')}
            </select>
          </div>
          <div class="adv-fld">
            <label>RAG Override</label>
            <select class="input-field" style="padding:5px 7px;font-size:12px"
                    onchange="App.quickUpdate('${kpi.id}','ragOverride',this.value||null)">
              <option value="" ${!kpi.ragOverride?'selected':''}>Auto</option>
              <option value="green"  ${kpi.ragOverride==='green'?'selected':''}>● Green</option>
              <option value="amber"  ${kpi.ragOverride==='amber'?'selected':''}>▲ Amber</option>
              <option value="red"    ${kpi.ragOverride==='red'?'selected':''}>✕ Red</option>
              <option value="neutral"${kpi.ragOverride==='neutral'?'selected':''}>○ N/A</option>
            </select>
          </div>
          <div class="adv-fld" style="grid-column:1/-1">
            <label>Comment / Note</label>
            <input type="text" class="input-field" style="padding:5px 7px;font-size:12px;width:100%"
                   value="${_esc(kpi.comment||'')}" placeholder="Optional note…"
                   onchange="App.quickUpdate('${kpi.id}','comment',this.value)">
          </div>
          <div style="grid-column:1/-1;display:flex;gap:8px;padding-top:2px">
            <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px" onclick="App.openEditKpi('${kpi.id}')">✎ Full Edit</button>
            <button class="btn btn-ghost" style="font-size:11px;padding:4px 10px;color:var(--rag-red)" onclick="App.confirmRemoveKpi('${kpi.id}')">✕ Remove</button>
          </div>
        </div>

        <!-- Month grid (always visible in advanced mode) -->
        <div class="adv-months-wrap">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
            <span style="font-size:10px;font-weight:700;color:var(--text-muted);letter-spacing:0.07em;text-transform:uppercase">
              Monthly Actuals
              <span style="font-weight:400;color:var(--brand-accent);margin-left:6px;text-transform:none;letter-spacing:0">
                · ${{sum:'summing',avg:'averaging',last:'latest month',max:'peak',min:'lowest'}[kpi.ytdMethod||'sum']}
              </span>
            </span>
            ${filled>0?`<span style="font-size:11px;color:var(--brand-accent)">
              ${filled}/${fyMonths.length} entered · YTD: <strong>${DataStore.formatValue(kpi.ytd,kpi)}</strong>
            </span>`:'<span style="font-size:11px;color:var(--text-muted)">No months entered yet</span>'}
          </div>
          <div class="month-grid">
            ${fyMonths.map(m=>{
              const hasVal = ma[m]!==undefined && ma[m]!==null;
              return `
                <div class="mc">
                  <label>${m}</label>
                  <div style="position:relative">
                    <input type="number" class="${hasVal?'filled':''}"
                           value="${hasVal?ma[m]:''}" placeholder="—"
                           style="padding:6px 4px ${hasVal?'16px':'6px'};width:100%"
                           onfocus="this.select()"
                           onchange="App.updateMonthlyActual('${kpi.id}','${m}',this.value)"
                           title="${m} actual for ${_esc(kpi.metric)}">
                    ${hasVal?`<span style="position:absolute;bottom:2px;left:50%;transform:translateX(-50%);font-size:10px;color:var(--brand-accent);font-weight:500;pointer-events:none;white-space:nowrap">${_fmtRaw(ma[m], kpi.unit)}</span>`:''}
                  </div>
                </div>`;
            }).join('')}
          </div>
        </div>

      </div>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  COLOUR RULES TAB
  // ══════════════════════════════════════════════════════════════════════════
  function _renderThresholdsTab() {
    const ths  = DataStore.getThresholds();
    const kpis = DataStore.getKpis();
    const BUILT_IN = ['th_manual_only','th_higher_better','th_lower_better','th_percent_high','th_percent_low'];

    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:10px">
          <div>
            <h3 style="font-size:15px;font-weight:600;margin-bottom:4px">Colour Rules (RAG)</h3>
            <p style="font-size:13px;color:var(--text-secondary);max-width:540px;line-height:1.6">
              These rules control <strong>what colour a KPI shows</strong> — Green, Amber, or Red — based on actual vs target.
              Separate from the target itself, which is set per-KPI with the op &amp; unit selectors.
            </p>
          </div>
          <button class="btn btn-primary" onclick="App.openAddThreshold()">+ New Rule</button>
        </div>

        <div style="background:rgba(58,134,255,0.07);border:1px solid rgba(58,134,255,0.18);border-radius:10px;
                    padding:12px 16px;margin-bottom:20px;font-size:12px;color:var(--text-secondary);line-height:1.7">
          <i class="fa-solid fa-lightbulb"></i> <strong style="color:var(--text-primary)">Target vs Colour Rule:</strong>
          The target (e.g. <code style="font-family:var(--font-mono);color:var(--brand-accent)">&gt; 80%</code>) defines success.
          The colour rule defines <em>how close you need to be</em> to turn Green or Amber.
          A KPI with an operator target and no rule auto-colours: Green if met, Red if not.
        </div>

        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px;margin-bottom:28px">
          ${ths.map(th=>{
            const usageCount = kpis.filter(k=>k.thresholdId===th.id).length;
            const isBuiltIn  = BUILT_IN.includes(th.id);
            return `
              <div class="card" style="position:relative">
                ${isBuiltIn?`<div class="label-xs" style="position:absolute;top:14px;right:14px;color:var(--text-muted)">Built-in</div>`:''}
                <div style="font-family:var(--font-display);font-size:14px;font-weight:600;margin-bottom:4px;padding-right:56px">${th.name}</div>
                <div style="font-size:12px;color:var(--text-secondary);margin-bottom:10px;line-height:1.5">${th.description||''}</div>
                <div class="label-xs" style="margin-bottom:8px">
                  Type: <span style="color:var(--brand-accent)">${{relative:'Relative (higher better)',relative_lower:'Relative (lower better)',absolute:'Absolute value',manual:'Manual only'}[th.type]||th.type}</span>
                </div>
                ${th.levels?.length?`
                  <div style="margin-bottom:10px">
                    ${th.levels.map(lv=>`
                      <div style="display:flex;align-items:center;gap:8px;padding:4px 8px;border-radius:6px;margin-bottom:3px;background:var(--bg-input)">
                        <span class="rag-badge ${lv.rag}" style="font-size:10px;padding:1px 6px;min-width:48px;text-align:center">${lv.rag}</span>
                        <span style="font-family:var(--font-mono);font-size:12px">${lv.op} ${lv.value}</span>
                        <span style="font-size:11px;color:var(--text-secondary)">${lv.label||''}</span>
                      </div>`).join('')}
                  </div>`:`<div style="font-size:12px;color:var(--text-muted);margin-bottom:10px;font-style:italic">Manual colouring only</div>`}
                <div style="display:flex;align-items:center;justify-content:space-between">
                  <span class="label-xs">${usageCount} KPI${usageCount!==1?'s':''} using this</span>
                  <div style="display:flex;gap:6px">
                    <button class="btn btn-ghost" style="font-size:11px;padding:3px 9px" onclick="App.openEditThreshold('${th.id}')">✎ Edit</button>
                    ${!isBuiltIn?`<button class="btn btn-ghost" style="font-size:11px;padding:3px 9px;color:var(--rag-red)" onclick="App.confirmRemoveThreshold('${th.id}')">Delete</button>`:''}
                  </div>
                </div>
              </div>`;
          }).join('')}
        </div>

        ${kpis.length>0?`
          <h3 style="font-size:14px;font-weight:600;margin-bottom:4px">Quick Assign Colour Rules</h3>
          <p style="font-size:12px;color:var(--text-secondary);margin-bottom:10px">Override forces a specific colour regardless of actuals.</p>
          <div style="overflow-x:auto;border:1px solid var(--border-card);border-radius:10px;background:var(--bg-card)">
            <table style="width:100%;border-collapse:collapse;font-size:12px">
              <thead>
                <tr style="background:var(--bg-input);color:var(--text-muted)">
                  <th style="padding:8px;text-align:left">KPI</th>
                  <th style="padding:8px;text-align:left;width:120px">Section</th>
                  <th style="padding:8px;text-align:left;width:110px">Target</th>
                  <th style="padding:8px;text-align:left;width:200px">Colour Rule</th>
                  <th style="padding:8px;text-align:left;width:88px">Override</th>
                  <th style="padding:8px;text-align:center;width:54px">RAG</th>
                </tr>
              </thead>
              <tbody>
                ${kpis.map((kpi,i)=>`
                  <tr style="border-top:1px solid var(--border-subtle);${i%2?'background:rgba(255,255,255,0.01)':''}">
                    <td style="padding:7px 8px;font-weight:500">${_esc(kpi.metric)}</td>
                    <td style="padding:7px 8px;color:var(--text-secondary);font-size:11px">${_esc(kpi.section)}</td>
                    <td style="padding:7px 8px;font-family:var(--font-mono);font-size:11px;color:var(--brand-accent)">${DataStore.formatTarget(kpi)||'—'}</td>
                    <td style="padding:7px 8px">
                      <select class="input-field" style="padding:4px 5px;font-size:11px;width:100%"
                              onchange="App.quickUpdate('${kpi.id}','thresholdId',this.value)">
                        ${ths.map(th=>`<option value="${th.id}" ${kpi.thresholdId===th.id?'selected':''}>${th.name}</option>`).join('')}
                      </select>
                    </td>
                    <td style="padding:7px 8px">
                      <select class="input-field" style="padding:4px 5px;font-size:11px;width:80px"
                              onchange="App.quickUpdate('${kpi.id}','ragOverride',this.value||null)">
                        <option value="" ${!kpi.ragOverride?'selected':''}>Auto</option>
                        <option value="green"  ${kpi.ragOverride==='green'?'selected':''}>● Green</option>
                        <option value="amber"  ${kpi.ragOverride==='amber'?'selected':''}>▲ Amber</option>
                        <option value="red"    ${kpi.ragOverride==='red'?'selected':''}>✕ Red</option>
                        <option value="neutral"${kpi.ragOverride==='neutral'?'selected':''}>○ N/A</option>
                      </select>
                    </td>
                    <td style="padding:7px 8px;text-align:center">${Components.ragDot(kpi.rag)}</td>
                  </tr>`).join('')}
              </tbody>
            </table>
          </div>`:``}
      </div>`;
  }

  // ── SETTINGS ──────────────────────────────────────────────────────────────
  function settings() {
    const s = DataStore.getSettings();
    const ALL_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `
      <div style="max-width:560px">
        <h2 class="section-title" style="margin-bottom:4px">Settings</h2>
        <p style="color:var(--text-secondary);font-size:13px;margin-bottom:24px">Configure your dashboard appearance and fiscal year.</p>

        <div class="card" style="margin-bottom:16px">
          <h3 style="font-size:14px;font-weight:600;margin-bottom:16px">Number Formatting</h3>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Currency Symbol</label>
              <select class="input-field" id="s-currency">
                ${['$','€','£','¥','A$','NZ$'].map(sym=>`<option value="${sym}" ${s.currencySymbol===sym?'selected':''}>${sym}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Decimal Places</label>
              <select class="input-field" id="s-decimals">
                ${[0,1,2,3].map(d=>`<option value="${d}" ${(s.decimals??2)===d?'selected':''}>${d} decimal${d!==1?'s':''}</option>`).join('')}
              </select>
            </div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Large Numbers</label>
              <select class="input-field" id="s-largenum">
                <option value="auto"  ${s.largeNumFormat==='auto'?'selected':''}>Auto (K / M / B)</option>
                <option value="M"     ${s.largeNumFormat==='M'?'selected':''}>Millions (M) only</option>
                <option value="K"     ${s.largeNumFormat==='K'?'selected':''}>Thousands (K) only</option>
                <option value="full"  ${s.largeNumFormat==='full'?'selected':''}>Full number</option>
              </select>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">% Storage Format</label>
              <select class="input-field" id="s-pctstorage">
                <option value="decimal" ${(s.pctStorage||'decimal')==='decimal'?'selected':''}>Decimal (0.40 → 40%)</option>
                <option value="direct"  ${s.pctStorage==='direct'?'selected':''}>Direct (40 → 40%)</option>
              </select>
            </div>
          </div>
          <div style="background:var(--bg-input);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text-secondary);margin-bottom:14px">
            Preview: <strong style="color:var(--brand-accent)" id="fmt-preview">—</strong>
            <button onclick="App.updateFmtPreview()" style="margin-left:10px;font-size:11px;color:var(--brand-accent);background:none;border:none;cursor:pointer">Refresh preview</button>
          </div>
          <button class="btn btn-primary" onclick="App.saveSettings()">Save Settings</button>
        </div>

        <div class="card" style="margin-bottom:16px">
          <h3 style="font-size:14px;font-weight:600;margin-bottom:16px">Company & Fiscal Year</h3>
          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Company Name</label>
            <input type="text" class="input-field" id="s-company" value="${_esc(s.companyName)}" placeholder="e.g. Acme Corp">
          </div>
          <div style="margin-bottom:12px">
            <label class="label-sm" style="display:block;margin-bottom:5px">Fiscal Year Label</label>
            <input type="text" class="input-field" id="s-fy" value="${_esc(s.fiscalYearLabel)}" placeholder="e.g. FY26">
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Default Reporting Period</label>
              <select class="input-field" id="s-period">
                <option value="monthly"  ${s.reportingPeriod==='monthly'?'selected':''}>Monthly</option>
                <option value="quarterly"${s.reportingPeriod==='quarterly'?'selected':''}>Quarterly</option>
                <option value="ytd"      ${s.reportingPeriod==='ytd'?'selected':''}>YTD</option>
                <option value="yearly"   ${s.reportingPeriod==='yearly'?'selected':''}>Full Year</option>
                <option value="last_fy"  ${s.reportingPeriod==='last_fy'?'selected':''}>Last FY</option>
              </select>
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Fiscal Year Start Month</label>
              <select class="input-field" id="s-fystart">
                ${ALL_MONTHS.map(m=>`<option value="${m}" ${s.fyStartMonth===m?'selected':''}>${m}</option>`).join('')}
              </select>
            </div>
          </div>
          <button class="btn btn-primary" onclick="App.saveSettings()">Save Settings</button>
        </div>

        <div class="card" style="margin-bottom:16px">
          <h3 style="font-size:14px;font-weight:600;margin-bottom:8px">White-Label Theme</h3>
          <p style="font-size:13px;color:var(--text-secondary);margin-bottom:10px">Edit <code style="font-family:var(--font-mono);color:var(--brand-accent)">css/theme.css</code> to customise colours and logo.</p>
          <div style="font-family:var(--font-mono);font-size:11px;color:var(--text-secondary);background:var(--bg-input);border-radius:8px;padding:12px;line-height:2">
            --brand-primary:  #0A2540;<br>
            --brand-accent:   #00C2A8;<br>
            --brand-logo-url: url('/assets/logo.svg');
          </div>
        </div>

        <div class="card">
          <h3 style="font-size:14px;font-weight:600;margin-bottom:8px">Data Management</h3>
          <p style="font-size:13px;color:var(--text-secondary);margin-bottom:12px">All data stored in browser localStorage. Export or reset below.</p>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn btn-ghost" onclick="App.exportData()">↓ Export JSON</button>
            <button class="btn btn-ghost" style="color:var(--rag-red)" onclick="App.confirmReset()"><i class="fa-solid fa-triangle-exclamation"></i> Reset All Data</button>
          </div>
        </div>
      </div>`;
  }

  // ── Utilities ─────────────────────────────────────────────────────────────
  function _sectionHeader(section, enc) {
    return `
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">
        <div class="label-xs" style="flex:1;padding-bottom:5px;border-bottom:1px solid var(--border-subtle)">${_esc(section)}</div>
        <button class="btn btn-ghost" style="font-size:10px;padding:3px 8px" onclick="App.promptRenameSection('${enc}')">✎ Rename</button>
        <button class="btn btn-ghost" style="font-size:10px;padding:3px 8px;color:var(--rag-red)" onclick="App.confirmRemoveSection('${enc}')"><i class="fa-solid fa-trash-can"></i> Remove</button>
      </div>`;
  }

  function _esc(s) {
    return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // ── FORMULAS PAGE ─────────────────────────────────────────────────────────
  function formulas() {
    const allKpis      = DataStore.getKpis();
    const formulaKpis  = allKpis.filter(k => k.isFormula);
    const sourceKpis   = allKpis.filter(k => !k.isFormula);
    const settings     = DataStore.getSettings();
    const mode         = settings.reportingPeriod || 'monthly';
    const fmlDefs      = DataStore.getFormulas();

    const opLabels = { sum:'Sum (+)', subtract:'Subtract (−)', multiply:'Multiply (×)', divide:'Divide (÷)', avg:'Average', min:'Minimum', max:'Maximum', custom:'Custom Expression' };

    const emptyState = formulaKpis.length === 0 ? `
      <div class="card" style="text-align:center;padding:60px 24px;margin-bottom:24px">
        <div style="font-size:48px;margin-bottom:16px">∑</div>
        <div style="font-family:var(--font-display);font-size:18px;font-weight:700;margin-bottom:8px">No Formula KPIs yet</div>
        <div style="font-size:13px;color:var(--text-secondary);max-width:420px;margin:0 auto 20px">
          Formula KPIs automatically compute their value from other KPIs — totals, averages, ratios, and custom expressions. Monthly data flows through too, so charts and trends work automatically.
        </div>
        <button class="btn btn-primary" onclick="App.openAddFormulaKpi()" style="font-size:14px;padding:10px 24px">
          ∑ Create Formula KPI
        </button>
      </div>` : '';

    const formulaCards = formulaKpis.map(kpi => {
      const fml   = fmlDefs.find(f => f.kpiId === kpi.id);
      const ps    = DataStore.getPeriodStats(kpi, mode);
      const operandNames = (fml?.operands || []).map(id => {
        const k = allKpis.find(x => x.id === id);
        return k ? k.metric : '(deleted)';
      });
      const ragCol = kpi.rag === 'green' ? 'var(--rag-green)' : kpi.rag === 'amber' ? 'var(--rag-amber)' : kpi.rag === 'red' ? 'var(--rag-red)' : 'var(--rag-neutral)';

      // Build a visual formula string
      const opSymbols = { sum:'+', subtract:'−', multiply:'×', divide:'÷', avg:'avg', min:'min', max:'max', custom:'ƒ' };
      const opSym = opSymbols[fml?.op] || '+';
      let formulaStr = fml?.op === 'custom'
        ? `<span style="font-family:var(--font-mono);font-size:11px;color:var(--text-secondary)">${_esc(fml.expression || '')}</span>`
        : operandNames.map((n,i) => `<span style="background:var(--bg-input);border-radius:4px;padding:2px 7px;font-size:11px">${_esc(n)}</span>${i<operandNames.length-1?` <span style="color:var(--brand-accent);font-weight:700">${opSym}</span> `:''}`).join('');

      return `
        <div class="card" style="position:relative;border-top:3px solid ${ragCol};margin-bottom:16px">
          <!-- Top row -->
          <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:14px">
            <div style="flex:1;min-width:0">
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:4px">
                <span style="font-family:var(--font-display);font-size:16px;font-weight:700">${_esc(kpi.metric)}</span>
                <span style="font-size:10px;padding:2px 7px;border-radius:4px;background:rgba(0,194,168,0.12);color:var(--brand-accent);font-weight:600;letter-spacing:0.06em">∑ FORMULA</span>
                ${kpi.isKey ? '<span style="font-size:10px;padding:2px 7px;border-radius:4px;background:rgba(58,134,255,0.12);color:var(--brand-accent-2);font-weight:600"><i class="fa-solid fa-star"></i> KEY</span>' : ''}
              </div>
              <div style="font-size:11px;color:var(--text-muted)">${_esc(kpi.section)} ${kpi.who ? '· ' + _esc(kpi.who) : ''}</div>
            </div>
            <div style="display:flex;gap:6px;flex-shrink:0">
              <button class="btn btn-ghost" style="font-size:11px;padding:5px 10px" onclick="App.openEditFormulaKpi('${kpi.id}')">✎ Edit</button>
              <button class="btn btn-ghost" style="font-size:11px;padding:5px 10px;color:var(--rag-red)" onclick="App.confirmRemoveFormulaKpi('${kpi.id}')">✕</button>
            </div>
          </div>

          <!-- Formula expression display -->
          <div style="background:var(--bg-input);border:1px solid var(--border-subtle);border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <span style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:var(--text-muted);white-space:nowrap">Formula:</span>
            <span style="font-size:12px;font-weight:600;padding:2px 8px;border-radius:4px;background:rgba(0,194,168,0.1);color:var(--brand-accent)">${opLabels[fml?.op] || 'Sum'}</span>
            <span style="color:var(--text-muted);font-size:11px">of</span>
            ${formulaStr || '<span style="color:var(--text-muted);font-size:11px">no operands</span>'}
          </div>

          <!-- KPI stats row -->
          <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:12px">
            <div>
              <div class="label-xs" style="margin-bottom:4px">Computed Value</div>
              <div style="font-family:var(--font-display);font-size:22px;font-weight:700;color:${ragCol}">${ps.actual !== null ? DataStore.formatValue(ps.actual, kpi) : '—'}</div>
            </div>
            <div>
              <div class="label-xs" style="margin-bottom:4px">Target</div>
              <div style="font-family:var(--font-display);font-size:16px;font-weight:600;color:var(--text-primary)">${DataStore.formatTarget(kpi)}</div>
            </div>
            <div>
              <div class="label-xs" style="margin-bottom:4px">YTD</div>
              <div style="font-family:var(--font-display);font-size:16px;font-weight:600;color:var(--text-secondary)">${kpi.ytd !== null ? DataStore.formatValue(kpi.ytd, kpi) : '—'}</div>
            </div>
            <div>
              <div class="label-xs" style="margin-bottom:4px">Status</div>
              <div class="rag-badge ${kpi.rag || 'neutral'}" style="margin-top:2px">${(kpi.rag||'neutral').charAt(0).toUpperCase()+(kpi.rag||'neutral').slice(1)}</div>
            </div>
          </div>

          <!-- Monthly breakdown sparkline-style -->
          ${_formulaMonthlyRow(kpi, DataStore.getFyMonths())}
        </div>`;
    }).join('');

    const howItWorks = `
      <div class="card" style="margin-top:24px;border:1px dashed var(--border-card)">
        <div style="font-family:var(--font-display);font-size:14px;font-weight:700;margin-bottom:12px;display:flex;align-items:center;gap:8px">
          <i class="fa-solid fa-circle-info" style="color:var(--brand-accent)"></i> How Formula KPIs work
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:16px;font-size:13px;color:var(--text-secondary);line-height:1.6">
          <div><strong style="color:var(--text-primary)">Auto-computed</strong><br>Values recalculate instantly whenever any source KPI is updated.</div>
          <div><strong style="color:var(--text-primary)">Monthly data flows through</strong><br>Each month's value is computed from source KPI monthly actuals, so charts, trends and period views all work.</div>
          <div><strong style="color:var(--text-primary)">Targets inherited</strong><br>Targets are computed by applying the same formula to the source KPI targets.</div>
          <div><strong style="color:var(--text-primary)">Dashboard-ready</strong><br>Toggle "Key KPI" to pin a formula result on your Overview dashboard like any other KPI.</div>
        </div>
      </div>`;

    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
          <div>
            <h2 class="page-title" style="margin:0 0 4px">Formula KPIs</h2>
            <div style="font-size:13px;color:var(--text-secondary)">Compute KPIs automatically from other KPIs using maths</div>
          </div>
          <button class="btn btn-primary" onclick="App.openAddFormulaKpi()" style="gap:8px">
            <span style="font-size:16px">∑</span> New Formula KPI
          </button>
        </div>

        ${sourceKpis.length === 0 ? `
          <div class="card" style="text-align:center;padding:40px;margin-bottom:24px">
            <div style="font-size:24px;margin-bottom:10px"><i class="fa-solid fa-triangle-exclamation"></i></div>
            <div style="font-size:14px;font-weight:600;margin-bottom:6px">No source KPIs yet</div>
            <div style="font-size:13px;color:var(--text-secondary);margin-bottom:14px">Add some KPIs in Data Entry first — formula KPIs compute their values from existing KPIs.</div>
            <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
          </div>` : ''}

        ${emptyState}
        ${formulaCards}
        ${formulaKpis.length > 0 || sourceKpis.length > 0 ? howItWorks : ''}
      </div>`;
  }

  function _formulaMonthlyRow(kpi, fyMonths) {
    const ma = kpi.monthlyActuals || {};
    const vals = fyMonths.map(m => ma[m] !== undefined ? parseFloat(ma[m]) : null);
    const hasData = vals.some(v => v !== null);
    if (!hasData) return '';

    const max = Math.max(...vals.filter(v => v !== null), 1);
    const bars = fyMonths.map((m, i) => {
      const v = vals[i];
      const h = v !== null ? Math.max(4, Math.round((v / max) * 36)) : 0;
      const col = v !== null ? 'var(--brand-accent)' : 'var(--border-subtle)';
      return `<div title="${m}: ${v !== null ? DataStore.formatValue(v, kpi) : '—'}"
                   style="flex:1;height:${h}px;background:${col};border-radius:2px 2px 0 0;opacity:0.75;align-self:flex-end;min-width:0;cursor:default;transition:opacity 0.1s"
                   onmouseover="this.style.opacity='1'" onmouseout="this.style.opacity='0.75'"></div>`;
    }).join('');

    const labels = fyMonths.map(m =>
      `<div style="flex:1;font-size:8px;text-align:center;color:var(--text-muted);overflow:hidden">${m}</div>`
    ).join('');

    return `
      <div style="margin-top:14px;border-top:1px solid var(--border-subtle);padding-top:12px">
        <div class="label-xs" style="margin-bottom:8px">Monthly Computed Values</div>
        <div style="display:flex;gap:2px;align-items:flex-end;height:40px">${bars}</div>
        <div style="display:flex;gap:2px;margin-top:2px">${labels}</div>
      </div>`;
  }

  return { overview, sectionPage, dataEntry, settings, formulas };
})();
/**
 * app.js v4
 * Router, event handlers, XLSX parser.
 * New: toggleAdvancedEntry, promptInsertSection, updateMonthlyActual
 * Removed: overviewKpis / toggleOverviewKpi (replaced by isKey on KPI)
 */

const App = (() => {
  let _currentPage    = 'overview';
  let _sidebarOpen    = false;
  let _dataEntryTab   = 'kpis';
  let _advancedMode   = false;
  let _kpiSearchQuery = '';   // search query for advanced data entry
  let _simpleSearchQuery = ''; // search query for simple data entry
  let _expandedKpiId  = null; // single-KPI monthly inline expand

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    DataStore.init();
    DataStore.subscribe(render);
    render();
  }

  // ── Routing ───────────────────────────────────────────────────────────────
  function navigate(page) {
    _currentPage = page;
    _sidebarOpen = false;
    if (!page.startsWith('data-entry')) { _dataEntryTab = 'kpis'; _expandedKpiId = null; _simpleSearchQuery = ''; }
    render();
    window.scrollTo(0, 0);
  }

  function setDataEntryTab(tab) { _dataEntryTab = tab; render(); }

  // ── Render ────────────────────────────────────────────────────────────────
  function render() {
    const settings = DataStore.getSettings();
    let pageTitle = 'Executive Overview';
    let pageHtml  = '';

    if (_currentPage === 'overview') {
      pageTitle = 'Executive Overview';
      pageHtml  = Pages.overview();
    } else if (_currentPage === 'formulas') {
      pageTitle = 'Formula KPIs';
      pageHtml  = Pages.formulas();
    } else if (_currentPage.startsWith('section:')) {
      const sectionName = decodeURIComponent(_currentPage.replace('section:', ''));
      pageTitle = sectionName;
      pageHtml  = Pages.sectionPage(sectionName);
    } else if (_currentPage === 'data-entry') {
      pageTitle = 'Data Entry';
      pageHtml  = Pages.dataEntry(_dataEntryTab, _advancedMode, _expandedKpiId);
    } else if (_currentPage === 'settings') {
      pageTitle = 'Settings';
      pageHtml  = Pages.settings();
    } else {
      pageTitle = 'Executive Overview';
      pageHtml  = Pages.overview();
    }

    const subtitle = settings.companyName + ' · ' + settings.fiscalYearLabel;

    document.getElementById('app-root').innerHTML = `
      <div class="app-shell">
        ${Components.sidebar(_currentPage)}
        <div class="sidebar-overlay ${_sidebarOpen?'active':''}" onclick="App.toggleSidebar()"></div>
        <div class="main-content">
          ${Components.topBar(pageTitle, subtitle)}
          <div class="page-content">${pageHtml}</div>
        </div>
      </div>`;

    if (_sidebarOpen) {
      const sb = document.getElementById('sidebar');
      if (sb) sb.classList.add('open');
    }
  }

  // ── Sidebar ───────────────────────────────────────────────────────────────
  function toggleSidebar() {
    _sidebarOpen = !_sidebarOpen;
    const sb = document.getElementById('sidebar');
    const ov = document.querySelector('.sidebar-overlay');
    if (sb) sb.classList.toggle('open', _sidebarOpen);
    if (ov) ov.classList.toggle('active', _sidebarOpen);
  }

  // ── Modals ────────────────────────────────────────────────────────────────
  function showModal(html) {
    const el = document.getElementById('modal-root') || document.createElement('div');
    el.id = 'modal-root';
    el.innerHTML = html;
    document.body.appendChild(el);
  }

  function closeModal() {
    const el = document.getElementById('modal-root');
    if (el) el.remove();
  }

  function openKpiDetail(id) {
    const kpi = DataStore.getKpiById(id);
    if (!kpi) return;
    const mode        = DataStore.getSettings().reportingPeriod || 'monthly';
    const periodStats = DataStore.getPeriodStats(kpi, mode);
    showModal(Components.kpiDetailModal(kpi, periodStats));
  }

  function openEditKpi(id) {
    const kpi = DataStore.getKpiById(id);
    if (kpi) showModal(Components.editKpiModal(kpi, false));
  }

  function openAddKpi(encodedSection) {
    const sections = DataStore.getSections();
    let defaultSection = sections[0] || 'General';
    if (encodedSection) defaultSection = decodeURIComponent(encodedSection);
    showModal(Components.editKpiModal({
      section: defaultSection, metric:'', who:'', unit:'', targetFY26:null,
      targetMonth:null, ytd:null, thresholdId:'th_manual_only',
      ragOverride:null, comment:'', isKey:false, monthlyActuals:{},
    }, true));
  }

  function openAddThreshold() {
    showModal(Components.thresholdEditModal({ name:'', description:'', type:'relative', levels:[] }, true));
  }

  function openEditThreshold(id) {
    const th = DataStore.getThresholdById(id);
    if (th) showModal(Components.thresholdEditModal(th, false));
  }

  function openAddSectionModal() {
    showModal(`
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:380px">
          <h3 style="font-family:var(--font-display);font-size:16px;font-weight:700;margin-bottom:16px">New Section</h3>
          <label class="label-sm" style="display:block;margin-bottom:6px">Section Name</label>
          <input type="text" id="new-section-name" class="input-field" placeholder="e.g. Customer Success" style="margin-bottom:16px"
                 onkeydown="if(event.key==='Enter')App.saveNewSection()">
          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" style="flex:1" onclick="App.saveNewSection()">Create Section</button>
            <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
          </div>
        </div>
      </div>`);
    setTimeout(()=>document.getElementById('new-section-name')?.focus(), 50);
  }

  function saveNewSection() {
    const name = document.getElementById('new-section-name')?.value?.trim();
    if (!name) { showToast('⚠ Enter a section name', 'amber'); return; }
    DataStore.addSection(name);
    closeModal();
    navigate('section:'+encodeURIComponent(name));
    showToast('✓ Section "'+name+'" created');
  }

  // ── Insert section break between KPIs ─────────────────────────────────────
  /**
   * Called when user clicks the insert-bar above a KPI row.
   * Prompts for a new section name, then calls DataStore.splitSectionAt
   * which re-assigns all KPIs from that KPI onwards (in the same section) to the new name.
   */
  function promptInsertSection(kpiId) {
    const kpi = DataStore.getKpiById(kpiId);
    if (!kpi) return;
    showModal(`
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:400px">
          <h3 style="font-family:var(--font-display);font-size:16px;font-weight:700;margin-bottom:8px">Insert Section Break</h3>
          <p style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">
            A new section will be created here. <strong>${kpi.metric}</strong> and all KPIs below it (within <em>${kpi.section}</em>) will move into this new section.
          </p>
          <label class="label-sm" style="display:block;margin-bottom:6px">New Section Name</label>
          <input type="text" id="insert-section-name" class="input-field" placeholder="e.g. Customer Success" style="margin-bottom:16px"
                 onkeydown="if(event.key==='Enter')App.confirmInsertSection('${kpiId}')">
          <div style="display:flex;gap:8px">
            <button class="btn btn-primary" style="flex:1" onclick="App.confirmInsertSection('${kpiId}')">Create &amp; Split</button>
            <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
          </div>
        </div>
      </div>`);
    setTimeout(()=>document.getElementById('insert-section-name')?.focus(), 50);
  }

  function confirmInsertSection(kpiId) {
    const name = document.getElementById('insert-section-name')?.value?.trim();
    if (!name) { showToast('⚠ Enter a section name', 'amber'); return; }
    DataStore.splitSectionAt(kpiId, name);
    closeModal();
    showToast('✓ Section "'+name+'" created');
  }

  // ── Monthly actual update ─────────────────────────────────────────────────
  function updateMonthlyActual(id, month, value) {
    DataStore.updateMonthlyActual(id, month, value);
  }

  // ── Auto-calc annual from monthly ─────────────────────────────────────────
  function autoCalcAnnual() {
    const mo = parseFloat(document.getElementById('e-tgt-month')?.value);
    const fy = document.getElementById('e-tgt-fy');
    if (!isNaN(mo) && mo && fy && !fy.value) fy.value = Math.round(mo * 12);
  }

  // ── Save KPI ──────────────────────────────────────────────────────────────
  function saveKpiEdit(id, isNew) {
    const metric = document.getElementById('e-metric')?.value?.trim();
    if (!metric) { showToast('⚠ KPI Name is required', 'amber'); return; }

    const section   = document.getElementById('e-section')?.value?.trim()    || 'General';
    const who       = document.getElementById('e-who')?.value?.trim()         || '';
    const tgtFY     = document.getElementById('e-tgt-fy')?.value;
    const tgtMo     = document.getElementById('e-tgt-month')?.value;
    const tgtOp     = document.getElementById('e-tgt-op')?.value              || null;
    const unit      = document.getElementById('e-unit')?.value                || '';
    const ytdVal    = document.getElementById('e-ytd')?.value;
    const thId      = document.getElementById('e-threshold')?.value;
    const ragOv     = document.getElementById('e-rag-override')?.value        || null;
    const comment   = document.getElementById('e-comment')?.value?.trim()     || '';
    const isKey     = document.getElementById('e-iskey')?.checked             || false;
    const ytdMethod = document.getElementById('e-ytdmethod')?.value           || 'sum';

    const fields = {
      metric, section, who, unit,
      targetFY26:    tgtFY !== ''    ? parseFloat(tgtFY)   : null,
      targetMonth:   tgtMo !== ''    ? parseFloat(tgtMo)   : null,
      targetFY26Op:  tgtOp || null,
      ytd:           ytdVal !== ''    ? parseFloat(ytdVal)    : null,
      thresholdId:   thId,
      ragOverride:   ragOv || null,
      comment, isKey, ytdMethod,
    };

    if (isNew) { DataStore.addKpi(fields); showToast('✓ KPI added'); }
    else       { DataStore.updateKpi(id, fields); showToast('✓ KPI saved'); }
    closeModal();
  }

  // ── Quick inline update ───────────────────────────────────────────────────
  function quickUpdate(id, field, value) {
    let v = value;
    if (field === 'isKey') {
      v = (value === true || value === 'true');
      // Keep overviewKpiIds in sync with isKey, identical logic to addKpiToOverview
      const settings = DataStore.getSettings();
      let ids = settings.overviewKpiIds
        ? [...settings.overviewKpiIds]
        : DataStore.getKeyKpis().map(k => k.id);
      if (v && !ids.includes(id)) ids.push(id);
      if (!v) ids = ids.filter(x => x !== id);
      DataStore.updateKpi(id, { isKey: v });
      DataStore.setOverviewKpiIds(ids);
      return;
    } else if (['targetFY26','targetMonth','ytd','actual'].includes(field)) {
      v = value === '' ? null : parseFloat(value);
    }
    DataStore.updateKpi(id, { [field]: v });
  }

  // ── Section management ────────────────────────────────────────────────────
  function promptRenameSection(encodedSection) {
    const oldName = decodeURIComponent(encodedSection);
    const newName = prompt('Rename section:', oldName);
    if (!newName || newName === oldName) return;
    DataStore.renameSection(oldName, newName);
    if (_currentPage === 'section:'+encodedSection) _currentPage = 'section:'+encodeURIComponent(newName);
    showToast('✓ Section renamed');
  }

  function confirmRemoveSection(encodedSection) {
    const name  = decodeURIComponent(encodedSection);
    const count = DataStore.getKpisBySection(name).length;
    if (!confirm(`Remove section "${name}" and its ${count} KPI${count!==1?'s':''}? This cannot be undone.`)) return;
    DataStore.removeSection(name);
    if (_currentPage === 'section:'+encodedSection) navigate('overview');
    showToast('✓ Section removed');
  }

  // ── Threshold management ──────────────────────────────────────────────────
  function addThresholdLevel() {
    const container = document.getElementById('th-levels-container');
    if (!container) return;
    const i = container.children.length;
    const div = document.createElement('div');
    div.id = 'level-row-'+i;
    div.style.cssText = 'display:grid;grid-template-columns:90px 60px 120px 1fr auto;gap:8px;align-items:center;margin-bottom:8px';
    div.innerHTML = `
      <select class="input-field" style="padding:6px 8px;font-size:12px" id="lv-rag-${i}">
        <option value="green">green</option><option value="amber">amber</option><option value="red">red</option>
      </select>
      <select class="input-field" style="padding:6px 8px;font-size:12px" id="lv-op-${i}">
        ${['>','>=','<','<=','='].map(op=>`<option value="${op}">${op}</option>`).join('')}
      </select>
      <input type="number" class="input-field" style="padding:6px 8px;font-size:12px" id="lv-val-${i}" value="0.90" step="0.01">
      <input type="text" class="input-field" style="padding:6px 8px;font-size:12px" id="lv-lbl-${i}" placeholder="Label">
      <button onclick="this.parentElement.remove()" style="color:var(--rag-red);background:none;border:none;cursor:pointer;font-size:16px">✕</button>`;
    container.appendChild(div);
  }

  function saveThreshold(id, isNew) {
    const name = document.getElementById('th-name')?.value?.trim();
    if (!name) { showToast('⚠ Threshold name is required', 'amber'); return; }
    const desc = document.getElementById('th-desc')?.value?.trim() || '';
    const type = document.getElementById('th-type')?.value || 'relative';
    const container = document.getElementById('th-levels-container');
    const levels = [];
    if (container) {
      Array.from(container.children).forEach((row,i) => {
        const rag = document.getElementById('lv-rag-'+i)?.value;
        const op  = document.getElementById('lv-op-'+i)?.value;
        const val = parseFloat(document.getElementById('lv-val-'+i)?.value);
        const lbl = document.getElementById('lv-lbl-'+i)?.value || '';
        if (rag && op && !isNaN(val)) levels.push({ rag, op, value:val, label:lbl });
      });
    }
    const fields = { name, description:desc, type, levels };
    if (isNew) { DataStore.addThreshold(fields); showToast('✓ Threshold created'); }
    else       { DataStore.updateThreshold(id, fields); showToast('✓ Threshold saved'); }
    closeModal();
  }

  function confirmRemoveThreshold(id) {
    const th = DataStore.getThresholdById(id);
    if (!th) return;
    const usageCount = DataStore.getKpis().filter(k=>k.thresholdId===id).length;
    if (!confirm(`Delete threshold "${th.name}"? ${usageCount} KPI${usageCount!==1?'s':''} will revert to Manual.`)) return;
    DataStore.removeThreshold(id);
    showToast('✓ Threshold deleted');
    closeModal();
  }

  // ── KPI remove ────────────────────────────────────────────────────────────
  function confirmRemoveKpi(id) {
    const kpi = DataStore.getKpiById(id);
    if (!kpi) return;
    if (!confirm(`Remove KPI "${kpi.metric}"? This cannot be undone.`)) return;
    DataStore.removeKpi(id);
    closeModal();
    showToast('✓ KPI removed');
  }

  // ── Settings ──────────────────────────────────────────────────────────────
  function saveSettings() {
    const companyName     = document.getElementById('s-company')?.value?.trim()    || 'My Company';
    const fiscalYearLabel = document.getElementById('s-fy')?.value?.trim()         || 'FY26';
    const reportingPeriod = document.getElementById('s-period')?.value             || 'monthly';
    const fyStartMonth    = document.getElementById('s-fystart')?.value            || 'Jul';
    const currencySymbol  = document.getElementById('s-currency')?.value           || '$';
    const decimals        = parseInt(document.getElementById('s-decimals')?.value  ?? '2');
    const largeNumFormat  = document.getElementById('s-largenum')?.value           || 'auto';
    const pctStorage      = document.getElementById('s-pctstorage')?.value         || 'decimal';
    DataStore.updateSettings({ companyName, fiscalYearLabel, reportingPeriod, fyStartMonth,
                               currencySymbol, decimals, largeNumFormat, pctStorage });
    showToast('✓ Settings saved');
  }

  function updateFmtPreview() {
    const el = document.getElementById('fmt-preview');
    if (!el) return;
    // show a sample currency and percent formatted with current (unsaved) selections
    const sym  = document.getElementById('s-currency')?.value   || '$';
    const dec  = parseInt(document.getElementById('s-decimals')?.value ?? '2');
    const lnf  = document.getElementById('s-largenum')?.value   || 'auto';
    const pcts = document.getElementById('s-pctstorage')?.value || 'decimal';
    const sampleKpiCurrency = { unit:'$' };
    const sampleKpiPct      = { unit:'%' };
    const origSettings = DataStore.getSettings();
    // Temporarily override for preview
    DataStore.updateSettings({ currencySymbol:sym, decimals:dec, largeNumFormat:lnf, pctStorage:pcts });
    const cv = DataStore.formatValue(1234567.89, sampleKpiCurrency);
    const pv = DataStore.formatValue(pcts==='decimal' ? 0.856 : 85.6, sampleKpiPct);
    // Restore
    DataStore.updateSettings(origSettings);
    el.textContent = `${cv}  ·  ${pv}`;
  }

  // ── XLSX / CSV Upload ─────────────────────────────────────────────────────
  function handleXlsxUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      const reader = new FileReader();
      reader.onload = e => {
        const rows = _parseCSV(e.target.result);
        const { updated, added } = DataStore.importFromRows(rows);
        showToast(`✓ ${updated} updated, ${added} added from CSV`);
      };
      reader.readAsText(file);
    } else {
      _loadSheetJs(() => _parseXlsxFile(file));
    }
  }

  function _loadSheetJs(cb) {
    if (window.XLSX) { cb(); return; }
    const s = document.createElement('script');
    s.src    = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = cb;
    document.head.appendChild(s);
  }

  function _parseXlsxFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb  = XLSX.read(e.target.result, { type:'array' });
        const ws  = wb.Sheets[wb.SheetNames[0]];

        // Read raw rows AND the raw cell objects so we can inspect format codes
        // sheet_to_json with raw:false gives formatted strings (e.g. "80%" not 0.8)
        // We read both and reconcile
        const rawRows  = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
        const fmtRows  = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:false });

        // ── Find header row ──────────────────────────────────────────────────
        let headerRow = -1, colMap = {};
        for (let i = 0; i < Math.min(8, rawRows.length); i++) {
          const row = (rawRows[i]||[]).map(c => String(c||'').toLowerCase().trim());
          if (row.some(c => c.includes('metric') || c.includes('kpi'))) {
            headerRow = i;
            row.forEach((c, idx) => {
              if      (c.includes('metric') || c === 'kpi')         colMap.metric      = idx;
              else if (c.includes('section') || c.includes('group')) colMap.section     = idx;
              else if (c.includes('who') || c.includes('owner'))     colMap.who         = idx;
              else if (c.includes('unit'))                           colMap.unit        = idx;
              else if (c.includes('label') || c.includes('display')) colMap.targetLabel = idx;
              else if (c.includes('fy') || c.includes('annual'))     colMap.targetFY26  = idx;
              else if (c.includes('month') && !c.includes('actual')) colMap.targetMonth = idx;
              else if (c.includes('ytd') || c.includes('year to'))   colMap.ytd         = idx;
              else if (c.includes('actual') || c.includes('current'))colMap.actual      = idx;
              else if (c.includes('rag') || c.includes('status'))    colMap.rag         = idx;
              else if (c.includes('comment') || c.includes('note'))  colMap.comment     = idx;
            });
            break;
          }
        }
        if (headerRow < 0) {
          headerRow = 1;
          colMap = { metric:0, who:1, targetFY26:2, ytd:3, targetMonth:4 };
        }

        // ── Cell parser — the heart of the fix ──────────────────────────────
        // Returns { value: number, op: string|null, unit: '$'|'%'|'', label: string }
        function parseCell(rawVal, fmtVal) {
          if (rawVal === null || rawVal === undefined) {
            return { value: NaN, op: null, unit: '', label: '' };
          }

          // Use the formatted string to detect operators and units from the cell text
          const fmtStr  = String(fmtVal || rawVal).trim();
          const rawStr  = String(rawVal).trim();

          // ── Step 1: detect operator prefix (>, >=, <, <=) ────────────────
          const opMatch = fmtStr.match(/^([<>]=?)\s*/);
          const op      = opMatch ? opMatch[1] : null;
          const afterOp = op ? fmtStr.slice(op.length).trim() : fmtStr;

          // ── Step 2: detect unit from the formatted string ─────────────────
          let unit = '';
          if (afterOp.startsWith('$'))      unit = '$';
          else if (afterOp.endsWith('%'))   unit = '%';
          // Also detect if the raw Excel cell was formatted as % (XLSX stores 0.8 but fmt shows 80%)
          else if (typeof rawVal === 'number' && fmtStr.includes('%')) unit = '%';

          // ── Step 3: extract numeric portion ──────────────────────────────
          const numStr  = afterOp.replace(/[$,%]/g, '').replace(/,/g, '').trim();
          let   value   = parseFloat(numStr);

          // If Excel stored a percentage as a decimal (0.85 formatted as "85%")
          // the rawVal will be e.g. 0.85 and fmtStr will be "85%" — keep rawVal
          if (unit === '%' && typeof rawVal === 'number' && rawVal <= 1 && rawVal >= -1
              && !rawStr.includes('%') && parseFloat(numStr) > 1) {
            value = rawVal; // already in decimal form (0.85), don't divide again
          } else if (unit === '%' && value > 1) {
            // Cell text was "85%" typed as text — normalise to decimal for storage
            value = value / 100;
          }

          // ── Step 4: build a human label for display ────────────────────────
          let label = '';
          if (!isNaN(value)) {
            const displayNum = unit === '%'
              ? (value * 100).toFixed(value * 100 % 1 === 0 ? 0 : 1) + '%'
              : unit === '$'
                ? '$' + (Math.abs(value) >= 1e6 ? (value/1e6).toFixed(1)+'M'
                       : Math.abs(value) >= 1000 ? (value/1000).toFixed(0)+'K'
                       : value.toLocaleString())
                : value % 1 === 0 ? value.toLocaleString() : value.toFixed(2);
            label = op ? op + ' ' + displayNum : displayNum;
          }

          return { value, op, unit, label };
        }

        // ── Build import rows ────────────────────────────────────────────────
        const rows = [];
        let currentSection = 'Imported';

        for (let i = headerRow + 1; i < rawRows.length; i++) {
          const row    = rawRows[i];
          const fmtRow = fmtRows[i];
          if (!row) continue;

          const metricRaw = String(row[colMap.metric ?? 0] || '').trim();
          if (!metricRaw) continue;

          // Detect section header rows (no target, no owner, no actuals)
          const hasTgt = colMap.targetFY26 !== undefined && row[colMap.targetFY26] !== null;
          const hasWho = colMap.who        !== undefined && row[colMap.who];
          const hasYtd = colMap.ytd        !== undefined && row[colMap.ytd] !== null;
          if (!hasTgt && !hasWho && !hasYtd) { currentSection = metricRaw; continue; }

          // Parse each numeric cell
          const tgtFY  = parseCell(row[colMap.targetFY26], fmtRow?.[colMap.targetFY26]);
          const tgtMo  = parseCell(row[colMap.targetMonth], fmtRow?.[colMap.targetMonth]);
          const ytdC   = parseCell(row[colMap.ytd],        fmtRow?.[colMap.ytd]);
          const actC   = colMap.actual !== undefined
                           ? parseCell(row[colMap.actual], fmtRow?.[colMap.actual])
                           : { value: NaN, op: null, unit: '', label: '' };

          // Unit resolution: explicit column wins, then first detected unit from cells
          let unit = '';
          if (colMap.unit !== undefined && row[colMap.unit]) {
            unit = String(row[colMap.unit]).trim();          // '$', '%', or ''
          } else {
            unit = tgtFY.unit || tgtMo.unit || ytdC.unit || actC.unit;
          }

          // Target label: explicit column wins, then auto-generated from parseCell
          let targetLabel = '';
          if (colMap.targetLabel !== undefined && row[colMap.targetLabel]) {
            targetLabel = String(row[colMap.targetLabel]).trim();
          } else if (tgtFY.label) {
            targetLabel = tgtFY.label;                       // e.g. "> 80%" or "$4.2M"
          }

          rows.push({
            section:      row[colMap.section] ? String(row[colMap.section]).trim() : currentSection,
            metric:       metricRaw,
            who:          colMap.who !== undefined && row[colMap.who] ? String(row[colMap.who]).trim() : '',
            unit,
            targetLabel,                                     // human-readable display target
            targetFY26:   tgtFY.value,                       // stored as clean number
            targetFY26Op: tgtFY.op,                          // '>', '>=', '<', '<=' or null
            targetMonth:  tgtMo.value,
            ytd:          ytdC.value,
            actual:       actC.value,
            rag:          colMap.rag !== undefined && row[colMap.rag]
                            ? String(row[colMap.rag]).toLowerCase().trim() : null,
            comment:      colMap.comment !== undefined ? (row[colMap.comment] || '') : '',
          });
        }

        const { updated, added } = DataStore.importFromRows(rows);
        showToast(`✓ ${updated} KPIs updated, ${added} new KPIs added`);
        render();
      } catch(err) {
        console.error('XLSX parse error:', err);
        showToast('⚠ Error reading file — check format', 'amber');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function _parseCSV(text) {
    const lines = text.split('\n').filter(l => l.trim());
    if (lines.length < 2) return [];
    const sep    = lines[0].includes('\t') ? '\t' : ',';
    const header = lines[0].split(sep).map(h => h.trim().toLowerCase().replace(/"/g, ''));
    const getIdx = (...keys) => { for (const k of keys) { const i = header.findIndex(h => h.includes(k)); if (i >= 0) return i; } return -1; };

    const cols = {
      metric:      getIdx('metric','kpi','name'),
      who:         getIdx('who','owner','assigned'),
      unit:        getIdx('unit'),
      targetLabel: getIdx('label','display','target_label'),
      targetFY26:  getIdx('fy26','target fy','annual target'),
      targetMonth: getIdx('month','monthly'),
      ytd:         getIdx('ytd','year to date'),
      actual:      getIdx('actual','current'),
      rag:         getIdx('rag','status'),
      comment:     getIdx('comment','note'),
      section:     getIdx('section','group','category'),
    };

    // Same cell parser logic as XLSX (text-only version)
    function parseCell(raw) {
      if (!raw) return { value: NaN, op: null, unit: '', label: '' };
      const s       = String(raw).trim();
      const opMatch = s.match(/^([<>]=?)\s*/);
      const op      = opMatch ? opMatch[1] : null;
      const afterOp = op ? s.slice(op.length).trim() : s;
      let unit = '';
      if (afterOp.startsWith('$'))    unit = '$';
      else if (afterOp.endsWith('%')) unit = '%';
      const numStr = afterOp.replace(/[$,%]/g, '').replace(/,/g, '').trim();
      let value    = parseFloat(numStr);
      if (unit === '%' && value > 1) value = value / 100;
      let label = '';
      if (!isNaN(value)) {
        const dn = unit === '%' ? (value*100).toFixed(value*100%1===0?0:1)+'%'
                 : unit === '$' ? '$'+(Math.abs(value)>=1e6?(value/1e6).toFixed(1)+'M':Math.abs(value)>=1000?(value/1000).toFixed(0)+'K':value.toLocaleString())
                 : value%1===0?value.toLocaleString():value.toFixed(2);
        label = op ? op+' '+dn : dn;
      }
      return { value, op, unit, label };
    }

    let currentSection = 'Imported';
    return lines.slice(1).map(line => {
      const c      = line.split(sep).map(c => c.trim().replace(/^"|"$/g, ''));
      const metric = cols.metric >= 0 ? c[cols.metric] : c[0];
      if (!metric) return null;

      const tgtFY  = parseCell(cols.targetFY26  >= 0 ? c[cols.targetFY26]  : c[2]);
      const tgtMo  = parseCell(cols.targetMonth >= 0 ? c[cols.targetMonth] : c[4]);
      const ytdC   = parseCell(cols.ytd         >= 0 ? c[cols.ytd]         : c[3]);
      const actC   = parseCell(cols.actual      >= 0 ? c[cols.actual]      : '');

      let unit = '';
      if (cols.unit >= 0 && c[cols.unit]) unit = c[cols.unit].trim();
      else unit = tgtFY.unit || tgtMo.unit || ytdC.unit || actC.unit;

      let targetLabel = '';
      if (cols.targetLabel >= 0 && c[cols.targetLabel]) targetLabel = c[cols.targetLabel].trim();
      else if (tgtFY.label) targetLabel = tgtFY.label;

      return {
        section:      cols.section >= 0 ? (c[cols.section] || currentSection) : currentSection,
        metric,
        who:          cols.who >= 0 ? (c[cols.who] || '') : '',
        unit,
        targetLabel,
        targetFY26:   tgtFY.value,
        targetFY26Op: tgtFY.op,
        targetMonth:  tgtMo.value,
        ytd:          ytdC.value,
        actual:       actC.value,
        rag:          cols.rag >= 0 ? (c[cols.rag] || '').toLowerCase() : null,
        comment:      cols.comment >= 0 ? (c[cols.comment] || '') : '',
      };
    }).filter(Boolean);
  }

  // ── Export / Reset ────────────────────────────────────────────────────────
  function exportData() {
    const data = {
      settings:   DataStore.getSettings(),
      thresholds: DataStore.getThresholds(),
      kpis:       DataStore.getKpis(),
      exportedAt: new Date().toISOString(),
    };
    const blob = new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href=url; a.download='exec-dashboard-export.json'; a.click(); URL.revokeObjectURL(url);
  }

  function confirmReset() {
    if (confirm('Reset all data? This will clear ALL KPIs and cannot be undone.')) {
      DataStore.resetToDefaults(); navigate('overview'); showToast('✓ Reset complete');
    }
  }

  // ── Toggle Key flag (show/hide on overview) ───────────────────────────────
  function toggleKpiKey(id) {
    const kpi = DataStore.getKpiById(id);
    if (!kpi) return;
    const willBeKey = !kpi.isKey;
    DataStore.updateKpi(id, { isKey: willBeKey });
    if (!willBeKey) {
      const s = DataStore.getSettings();
      if (s.overviewKpiIds) {
        DataStore.setOverviewKpiIds(s.overviewKpiIds.filter(x => x !== id));
      }
    }
  }

  // ── Toggle single-KPI inline full expand ─────────────────────────────────
  function toggleKpiMonthly(id) {
    _expandedKpiId = (_expandedKpiId === id) ? null : id;
    render();
    if (_expandedKpiId) {
      setTimeout(() => {
        const el = document.querySelector(`[data-kpi-id="${id}"]`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, 80);
    }
  }

  // ── KPI search (advanced mode) ────────────────────────────────────────────
  function setKpiSearch(q) {
    _kpiSearchQuery = q;
    render();
    setTimeout(() => {
      const box = document.getElementById('kpi-search-adv');
      if (box) { box.focus(); box.setSelectionRange(q.length, q.length); }
    }, 10);
  }

  function setSimpleKpiSearch(q) {
    _simpleSearchQuery = q;
    // Filter rows in-place without full re-render
    const tables = document.querySelectorAll('[data-section-table]');
    const query = q.toLowerCase().trim();
    tables.forEach(table => {
      // Each KPI occupies a spacer tr + a data tr (+ optional monthRow tr)
      // We target data rows by their data-kpi-id attribute
      const rows = table.querySelectorAll('tr[data-kpi-id]');
      rows.forEach(row => {
        const name = row.querySelector('td:nth-child(2)')?.textContent?.toLowerCase() || '';
        const match = !query || name.includes(query);
        // Hide/show the data row
        row.style.display = match ? '' : 'none';
        // Hide/show its preceding spacer row
        const spacer = row.previousElementSibling;
        if (spacer?.classList.contains('insert-spacer')) spacer.style.display = match ? '' : 'none';
        // Hide/show its following month-expand row if present
        const next = row.nextElementSibling;
        if (next && !next.classList.contains('insert-spacer') && !next.hasAttribute('data-kpi-id')) {
          next.style.display = match ? '' : 'none';
        }
      });
    });
    // Keep search box focused
    const box = document.getElementById('kpi-search-simple');
    if (box) { box.focus(); box.setSelectionRange(q.length, q.length); }
  }

  function toggleAdvancedEntry() {
    _advancedMode = !_advancedMode;
    if (!_advancedMode) _kpiSearchQuery = '';
    render();
  }

  // ── Navigate KPI card → advanced data entry ───────────────────────────────
  function openKpiInDataEntry(id) {
    _advancedMode = true;
    _dataEntryTab = 'kpis';
    navigate('data-entry');
    // Scroll to the card after render
    setTimeout(() => {
      const el = document.querySelector(`[data-kpi-id="${id}"]`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  }

  // ── Manage KPIs dropdown (Overview) ───────────────────────────────────────
  function toggleManageKpiDropdown() {
    const dd = document.getElementById('manage-kpi-dropdown');
    if (!dd) return;
    const isOpen = dd.style.display !== 'none';
    dd.style.display = isOpen ? 'none' : 'block';
    if (!isOpen) {
      setTimeout(() => document.getElementById('manage-kpi-search')?.focus(), 50);
      // Close on outside click
      setTimeout(() => {
        const handler = (e) => {
          const wrap = document.getElementById('manage-kpi-wrap');
          if (wrap && !wrap.contains(e.target)) {
            dd.style.display = 'none';
            document.removeEventListener('click', handler);
          }
        };
        document.addEventListener('click', handler);
      }, 0);
    }
  }

  function toggleOverviewKpi(id, checked) {
    const settings = DataStore.getSettings();
    let ids = settings.overviewKpiIds ? [...settings.overviewKpiIds] : DataStore.getKeyKpis().map(k => k.id);
    if (checked && !ids.includes(id)) ids.push(id);
    if (!checked) ids = ids.filter(x => x !== id);
    DataStore.setOverviewKpiIds(ids);
    const badge = document.querySelector('#manage-kpi-wrap button span');
    if (badge) badge.textContent = `${DataStore.getOverviewKpis().length}/${DataStore.getKpis().length}`;
  }

  // Called from any overview checkbox — syncs both isKey and overviewKpiIds atomically
  function addKpiToOverview(id, checked) {
    const kpi = DataStore.getKpiById(id);
    if (!kpi) return;
    // Sync isKey flag
    if (checked && !kpi.isKey) DataStore.updateKpi(id, { isKey: true });
    if (!checked && kpi.isKey) DataStore.updateKpi(id, { isKey: false });
    // Sync overviewKpiIds
    const settings = DataStore.getSettings();
    let ids = settings.overviewKpiIds
      ? [...settings.overviewKpiIds]
      : DataStore.getKeyKpis().map(k => k.id);
    if (checked && !ids.includes(id)) ids.push(id);
    if (!checked) ids = ids.filter(x => x !== id);
    DataStore.setOverviewKpiIds(ids);
  }

  function setAllOverviewKpis(selected) {
    const allKpis = DataStore.getKpis();
    if (selected) {
      allKpis.forEach(k => { if (!k.isKey) DataStore.updateKpi(k.id, { isKey: true }); });
      DataStore.setOverviewKpiIds(allKpis.map(k => k.id));
    } else {
      allKpis.forEach(k => { if (k.isKey) DataStore.updateKpi(k.id, { isKey: false }); });
      DataStore.setOverviewKpiIds([]);
    }
    render();
  }

  function filterManageKpiList(q) {
    const list = document.getElementById('manage-kpi-list');
    if (!list) return;
    const items = list.querySelectorAll('div[onmouseover]');
    const query = q.toLowerCase().trim();
    items.forEach(div => {
      // KPI name is in the first child div's first child div (font-size:12px name element)
      const nameEl = div.querySelector('div > div:first-child');
      const name = nameEl?.textContent?.toLowerCase() || '';
      div.style.display = (!query || name.includes(query)) ? '' : 'none';
    });
  }

  // ── Formula KPI management ────────────────────────────────────────────────
  function openAddFormulaKpi() {
    const kpis = DataStore.getKpis().filter(k => !k.isFormula);
    showModal(Components.formulaKpiModal(null, null, kpis, false));
  }

  function openEditFormulaKpi(kpiId) {
    const kpi = DataStore.getKpiById(kpiId);
    const fml = DataStore.getFormulaByKpiId(kpiId);
    const kpis = DataStore.getKpis().filter(k => !k.isFormula && k.id !== kpiId);
    if (kpi && fml) showModal(Components.formulaKpiModal(kpi, fml, kpis, true));
  }

  function saveFormulaKpi(kpiId, isEdit) {
    const metric  = document.getElementById('fml-metric')?.value?.trim();
    if (!metric) { showToast('⚠ Name is required', 'amber'); return; }
    const section = document.getElementById('fml-section')?.value?.trim() || 'Formulas';
    const unit    = document.getElementById('fml-unit')?.value || '';
    const isKey   = document.getElementById('fml-iskey')?.checked || false;
    const thId    = document.getElementById('fml-threshold')?.value || 'th_manual_only';
    const op      = document.getElementById('fml-op')?.value || 'sum';
    const comment = document.getElementById('fml-comment')?.value?.trim() || '';

    // Collect selected operand KPI ids (checkboxes)
    const operands = Array.from(document.querySelectorAll('.fml-operand-cb:checked')).map(el => el.value);
    if (operands.length === 0 && op !== 'custom') {
      showToast('⚠ Select at least one source KPI', 'amber'); return;
    }

    let expression = '';
    if (op === 'custom') {
      expression = document.getElementById('fml-expression')?.value?.trim() || '';
      if (!expression) { showToast('⚠ Enter a custom expression', 'amber'); return; }
    }

    const kpiFields  = { metric, section, unit, isKey, thresholdId: thId, comment };
    const formulaDef = { op, operands, expression };

    if (isEdit) {
      DataStore.updateFormulaKpi(kpiId, kpiFields, formulaDef);
      showToast('✓ Formula KPI updated');
    } else {
      DataStore.addFormulaKpi(kpiFields, formulaDef);
      showToast('✓ Formula KPI created');
    }
    closeModal();
  }

  function confirmRemoveFormulaKpi(kpiId) {
    const kpi = DataStore.getKpiById(kpiId);
    if (!kpi) return;
    if (!confirm(`Remove formula KPI "${kpi.metric}"? This cannot be undone.`)) return;
    DataStore.removeFormulaKpi(kpiId);
    showToast('✓ Formula KPI removed');
  }

  // ── Toast ─────────────────────────────────────────────────────────────────
  function showToast(msg, type='green') {
    const old = document.querySelector('.toast');
    if (old) old.remove();
    const el = document.createElement('div');
    el.className = 'toast';
    el.style.borderLeft = `3px solid var(--rag-${type})`;
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(()=>el?.remove(), 3000);
  }

  return {
    init, navigate, setDataEntryTab, toggleAdvancedEntry, render, toggleSidebar,
    get _kpiSearchQuery()    { return _kpiSearchQuery; },
    get _simpleSearchQuery() { return _simpleSearchQuery; },
    get _expandedKpiId()     { return _expandedKpiId; },
    setKpiSearch, setSimpleKpiSearch, openKpiInDataEntry, toggleKpiKey, toggleKpiMonthly,
    toggleManageKpiDropdown, toggleOverviewKpi, addKpiToOverview, setAllOverviewKpis, filterManageKpiList,
    updateFmtPreview,
    openKpiDetail, openEditKpi, openAddKpi,
    openAddThreshold, openEditThreshold, openAddSectionModal, saveNewSection,
    promptInsertSection, confirmInsertSection,
    updateMonthlyActual,
    autoCalcAnnual, saveKpiEdit, quickUpdate,
    confirmRemoveKpi, promptRenameSection, confirmRemoveSection,
    addThresholdLevel, saveThreshold, confirmRemoveThreshold,
    openAddFormulaKpi, openEditFormulaKpi, saveFormulaKpi, confirmRemoveFormulaKpi,
    saveSettings, handleXlsxUpload, exportData, confirmReset, showToast, closeModal,
  };
})();
