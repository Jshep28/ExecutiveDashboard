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
