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
    // Migrate: stamp isFormula on any KPI that has a matching formula definition
    const fmlKpiIds = new Set(_formulas.map(f => f.kpiId));
    let migrated = false;
    _kpis.forEach(k => {
      if (fmlKpiIds.has(k.id) && !k.isFormula) { k.isFormula = true; migrated = true; }
    });
    if (migrated) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(_kpis)); } catch(e) {} }
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
      isFormula:      fields.isFormula       || false,
      monthlyActuals: fields.monthlyActuals  || {},
      ytdMethod:      fields.ytdMethod       || 'sum',   // 'sum' | 'avg' | 'last' | 'max' | 'min'
      autoCalcTargets: fields.autoCalcTargets || 'none', // 'none' | 'fy_from_month' | 'month_from_fy'
    };
    _kpis.push(kpi);
    _computeRag(); _save(); _notify();
    return id;
  }

  function updateKpi(id, fields) {
    const kpi = _kpis.find(k=>k.id===id);
    if (!kpi) return;
    Object.assign(kpi, fields);
    // Per-KPI auto-calc: only when the user has opted in via autoCalcTargets
    const act   = kpi.autoCalcTargets || 'none';
    const isPct = kpi.unit === '%';
    if (act === 'fy_from_month' && kpi.targetMonth !== null && fields.targetFY26 === undefined)
      kpi.targetFY26 = isPct ? kpi.targetMonth : kpi.targetMonth * 12;
    else if (act === 'month_from_fy' && kpi.targetFY26 !== null && fields.targetMonth === undefined)
      kpi.targetMonth = isPct ? kpi.targetFY26 : kpi.targetFY26 / 12;
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

  /** Recompute ytd/actual for a KPI from its existing monthlyActuals using its current ytdMethod */
  function recomputeKpiYtd(id) {
    const kpi = _kpis.find(k=>k.id===id);
    if (!kpi) return;
    const fyMonths = getFyMonths();
    const entered = fyMonths.filter(m => kpi.monthlyActuals?.[m] !== undefined);
    const vals = entered.map(m => kpi.monthlyActuals[m]);
    if (vals.length === 0) { kpi.ytd = null; kpi.actual = null; }
    else {
      const method = kpi.ytdMethod || 'sum';
      if      (method === 'sum')  kpi.ytd = vals.reduce((a,b)=>a+b, 0);
      else if (method === 'avg')  kpi.ytd = vals.reduce((a,b)=>a+b, 0) / vals.length;
      else if (method === 'last') kpi.ytd = vals[vals.length - 1];
      else if (method === 'max')  kpi.ytd = Math.max(...vals);
      else if (method === 'min')  kpi.ytd = Math.min(...vals);
      else                        kpi.ytd = vals.reduce((a,b)=>a+b, 0);
      kpi.actual = kpi.monthlyActuals[entered[entered.length - 1]];
    }
    _applyAllFormulas(); _computeRag(); _save(); _notify();
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

    // Helper: derive ytd + actual from a monthlyActuals map
    function _deriveFromMonthly(monthlyActuals, ytdMethod) {
      const fyMs   = getFyMonths();
      const entered = fyMs.filter(m => monthlyActuals[m] !== undefined);
      if (!entered.length) return {};
      const vals = entered.map(m => monthlyActuals[m]);
      let ytd;
      if      (ytdMethod === 'avg')  ytd = vals.reduce((a,b)=>a+b,0)/vals.length;
      else if (ytdMethod === 'last') ytd = vals[vals.length-1];
      else if (ytdMethod === 'max')  ytd = Math.max(...vals);
      else if (ytdMethod === 'min')  ytd = Math.min(...vals);
      else                           ytd = vals.reduce((a,b)=>a+b,0); // sum default
      return { ytd, actual: monthlyActuals[entered[entered.length-1]] };
    }

    rows.forEach(row => {
      if (!row.metric) return;
      const existing = _kpis.find(k => k.metric.toLowerCase().trim() === row.metric.toLowerCase().trim());
      if (existing) {
        const u = {};
        if (row.who        !== undefined)                    u.who          = row.who;
        if (row.unit)                                        u.unit         = row.unit;
        if (row.targetLabel)                                 u.targetLabel  = row.targetLabel;
        if (row.targetFY26Op)                                u.targetFY26Op = row.targetFY26Op;
        // Always include targetFY26 explicitly — even when empty — so updateKpi's
        // auto-calc (targetMonth * 12) never fires and overwrites a blank import cell.
        u.targetFY26 = !isNaN(parseFloat(row.targetFY26))
          ? parseFloat(row.targetFY26)
          : existing.targetFY26 ?? null;
        if (!isNaN(parseFloat(row.ytd)))                     u.ytd          = parseFloat(row.ytd);
        if (!isNaN(parseFloat(row.targetMonth)))             u.targetMonth  = parseFloat(row.targetMonth);
        if (!isNaN(parseFloat(row.actual)))                  u.actual       = parseFloat(row.actual);
        if (row.rag)                                         u.ragOverride  = row.rag.toLowerCase().trim();
        if (row.comment)                                     u.comment      = row.comment;
        // Merge monthly actuals and re-derive ytd/actual
        if (row.monthlyActuals && Object.keys(row.monthlyActuals).length > 0) {
          u.monthlyActuals = { ...(existing.monthlyActuals || {}), ...row.monthlyActuals };
          Object.assign(u, _deriveFromMonthly(u.monthlyActuals, existing.ytdMethod || 'sum'));
        }
        updateKpi(existing.id, u); updated++;
      } else {
        const kpiFields = {
          section:      row.section      || 'Imported',
          metric:       row.metric,
          who:          row.who          || '',
          unit:         row.unit         || '',
          targetLabel:  row.targetLabel  || '',
          targetFY26Op: row.targetFY26Op || null,
          targetFY26:   isNaN(parseFloat(row.targetFY26))  ? null : parseFloat(row.targetFY26),
          targetMonth:  isNaN(parseFloat(row.targetMonth)) ? null : parseFloat(row.targetMonth),
          ytd:          isNaN(parseFloat(row.ytd))         ? null : parseFloat(row.ytd),
        };
        if (row.monthlyActuals && Object.keys(row.monthlyActuals).length > 0) {
          kpiFields.monthlyActuals = row.monthlyActuals;
          Object.assign(kpiFields, _deriveFromMonthly(row.monthlyActuals, 'sum'));
        }
        addKpi(kpiFields); added++;
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
      const curQ     = getCurrentFyQuarter();
      // Try current quarter first; if empty, walk back to find the most recent quarter with data
      let q = curQ;
      let qMonths = getFyQuarterMonths(q);
      let qActual = aggregateMonths(qMonths);
      if (qActual === null && q > 1) {
        for (let tryQ = q - 1; tryQ >= 1; tryQ--) {
          const tryMonths = getFyQuarterMonths(tryQ);
          const tryActual = aggregateMonths(tryMonths);
          if (tryActual !== null) { q = tryQ; qMonths = tryMonths; qActual = tryActual; break; }
        }
      }
      actual         = qActual;
      // Point-in-time methods compare against monthly target; additive (sum) scales up
      target         = isPointInTime
        ? tgtMonth
        : (tgtMonth !== null ? tgtMonth * 3 : (tgtFY !== null ? tgtFY / 4 : null));
      label          = 'Q' + q + ' ' + fyLabel + (q !== curQ ? ' (latest)' : '');
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

  /** Get KPIs to show on the Overview page (filtered and ordered by overviewKpiIds if set). */
  function getOverviewKpis() {
    const ids = _settings.overviewKpiIds;
    const keyKpis = _kpis.filter(k => k.isKey);
    if (!ids) return keyKpis;
    return keyKpis.filter(k => ids.includes(k.id))
                  .sort((a, b) => ids.indexOf(a.id) - ids.indexOf(b.id));
  }

  /** Persist the set of KPI ids selected for the overview panel. null = all key KPIs. */
  function setOverviewKpiIds(ids) {
    _settings.overviewKpiIds = ids;
    _save(); _notify();
  }
  function _notify() { _subscribers.forEach(fn=>fn()); }

  function restoreFromBackup(data) {
    if (data.kpis)       { _kpis = data.kpis; _nextId = _kpis.reduce((m,k)=>Math.max(m,parseInt(k.id?.replace('kpi_','')||0)),0)+1; }
    if (data.thresholds) { _thresholds = data.thresholds; }
    if (data.formulas)   { _formulas = data.formulas; _saveFormulas(); }
    if (data.settings)   { Object.assign(_settings, data.settings); }
    _applyAllFormulas(); _computeRag(); _save(); _notify();
  }

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
    importFromRows, formatValue, formatTarget, progressPct, resetToDefaults, restoreFromBackup, subscribe, getOverviewKpis, setOverviewKpiIds, recomputeKpiYtd,
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
        ${kpi.isKey?'★':'☆'}
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
      const pageId  = 'section:' + encodeURIComponent(section);
      const sKpis   = DataStore.getKpisBySection(section);
      const _mode   = settings.reportingPeriod || 'monthly';
      const counts  = {green:0,amber:0,red:0,neutral:0};
      sKpis.forEach(k=>{
        const ps = DataStore.getPeriodStats(k, _mode);
        counts[ps.rag || k.rag || 'neutral']++;
      });
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
  function ragSummaryBar(kpis, mode) {
    const _mode = mode || DataStore.getSettings().reportingPeriod || 'monthly';
    const counts = {green:0,amber:0,red:0,neutral:0};
    kpis.forEach(k=>{
      const ps = DataStore.getPeriodStats(k, _mode);
      counts[ps.rag || k.rag || 'neutral']++;
    });
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
            <button class="btn btn-ghost" onclick="App.openEditKpi('${kpi.id}')">✎ Edit KPI</button>
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

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:6px">
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Target / Year</label>
              <input type="number" id="e-tgt-fy" class="input-field" value="${kpi.targetFY26??''}" placeholder="Annual target" oninput="App.autoCalcAnnual('fy')">
            </div>
            <div>
              <label class="label-sm" style="display:block;margin-bottom:5px">Target / Month</label>
              <input type="number" id="e-tgt-month" class="input-field" value="${kpi.targetMonth??''}" placeholder="Monthly target" oninput="App.autoCalcAnnual('month')">
            </div>
          </div>
          <div style="margin-bottom:12px;padding:8px 10px;background:var(--bg-input);border-radius:8px;display:flex;align-items:center;gap:16px;flex-wrap:wrap">
            <span style="font-size:11px;color:var(--text-muted);font-weight:500">Auto-calculate:</span>
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:12px;color:var(--text-secondary)">
              <input type="radio" name="e-autocalc" id="e-autocalc-none" value="none" ${(kpi.autoCalcTargets||'none')==='none'?'checked':''}
                     style="accent-color:var(--brand-accent)" onchange="App.autoCalcAnnual('mode')"> Off
            </label>
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:12px;color:var(--text-secondary)">
              <input type="radio" name="e-autocalc" id="e-autocalc-fy" value="fy_from_month" ${kpi.autoCalcTargets==='fy_from_month'?'checked':''}
                     style="accent-color:var(--brand-accent)" onchange="App.autoCalcAnnual('mode')"> Year = Month × 12
            </label>
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:12px;color:var(--text-secondary)">
              <input type="radio" name="e-autocalc" id="e-autocalc-mo" value="month_from_fy" ${kpi.autoCalcTargets==='month_from_fy'?'checked':''}
                     style="accent-color:var(--brand-accent)" onchange="App.autoCalcAnnual('mode')"> Month = Year ÷ 12
            </label>
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
        <div style="font-size:28px;margin-bottom:10px">📊</div>
        <div style="font-family:var(--font-display);font-size:15px;font-weight:600;margin-bottom:6px">Import XLSX / CSV</div>
        <div style="font-size:12px;color:var(--text-secondary);margin-bottom:14px;line-height:1.6">
          Upload a spreadsheet with columns:<br>
          <strong>A: Metric &nbsp;·&nbsp; B: Who &nbsp;·&nbsp; C: Target FY &nbsp;·&nbsp; D: Target / Month &nbsp;·&nbsp; E: YTD</strong><br>
          Matching KPIs are updated. New ones are added.
        </div>
        <input type="file" id="xlsx-file-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="App.handleXlsxUpload(event)">
        <button class="btn btn-primary" onclick="document.getElementById('xlsx-file-input').click()">↑ Choose File</button>
        <div style="font-size:10px;color:var(--text-muted);margin-top:10px">Supports .xlsx, .xls, .csv</div>
      </div>`;
  }

  // ── Shared utility ────────────────────────────────────────────────────────
  function _esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
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
      : sourceKpis.map((k, i) => {
          const ps = DataStore.getPeriodStats(k, DataStore.getSettings().reportingPeriod || 'monthly');
          return `
            <label class="fml-kpi-row" style="display:flex;align-items:center;gap:10px;padding:8px 12px;cursor:pointer;transition:background 0.1s;${i>0?'border-top:1px solid var(--border-subtle)':''}"
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
                <label for="fml-iskey" class="label-sm" style="cursor:pointer">★ Pin to Overview (Key KPI)</label>
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
            <div style="position:relative;margin-bottom:8px">
              <span style="position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none">⌕</span>
              <input type="text" id="fml-kpi-search" class="input-field"
                     style="padding-left:30px;font-size:12px"
                     placeholder="Search source KPIs…"
                     oninput="OverviewPanel._fmlFilterKpis(this.value)">
            </div>
            <div id="fml-kpi-list" style="max-height:220px;overflow-y:auto;border:1px solid var(--border-subtle);border-radius:8px;background:var(--bg-card)">
              ${kpiCheckboxes}
            </div>
            <div id="fml-kpi-empty" style="display:none;padding:12px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs match your search</div>
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
            <div style="position:relative;margin-bottom:8px">
              <span style="position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none">⌕</span>
              <input type="text" id="fml-ref-search" class="input-field"
                     style="padding-left:30px;font-size:12px"
                     placeholder="Search KPI reference…"
                     oninput="OverviewPanel._fmlFilterRef(this.value)">
            </div>
            <div style="border:1px solid var(--border-subtle);border-radius:8px;overflow:hidden">
              <table style="width:100%;border-collapse:collapse" id="fml-ref-table">
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
 * overview-panel.js  v3
 *
 * Changes from v2:
 *  - Per surrounding-card visualization type: 'card' | 'odometer' | 'bar' | 'line'
 *  - Each KPI card has a ••• dropdown to switch its own visualization
 *  - Card viz types stored in overviewPanel.cardVizTypes: { kpiId: 'card'|'odometer'|'bar'|'line' }
 *  - All card viz types render within a single card shell (same size)
 *  - mountCharts() now also mounts per-card mini charts
 */

const OverviewPanel = (() => {

  // ── State ─────────────────────────────────────────────────────────────────
  function _getState() {
    const s = DataStore.getSettings();
    return s.overviewPanel || {
      chartType:     'odometer',
      heroKpiIds:    [],
      heroKpiOrder:  [],
      stackedKpiIds: {},
      heroPosition:  'center',
      kpiPickerOpen: false,
      menuOpen:      false,
      cardVizTypes:  {},   // { [kpiId]: 'card' | 'odometer' | 'bar' | 'line' }
      cardMenuOpen:  null, // kpiId whose ••• menu is open, or null
    };
  }

  function _setState(patch) {
    const cur = _getState();
    DataStore.updateSettings({ overviewPanel: { ...cur, ...patch } });
  }

  // ── Drag-and-drop state (module-level, survives re-renders) ──────────────
  let _dragId     = null;   // id of KPI being dragged
  let _dragOverId = null;   // id of card currently hovered

  // ── Public API ────────────────────────────────────────────────────────────
  function setChartType(type) { _setState({ chartType: type, menuOpen: false }); }
  function toggleMenu()       { _setState({ menuOpen: !_getState().menuOpen }); setTimeout(_fixHeroDropdowns, 0); }
  function closeMenu()        { _setState({ menuOpen: false }); }
  function setHeroPosition(p) { _setState({ heroPosition: p, menuOpen: false }); }
  function toggleKpiPicker()  { _setState({ kpiPickerOpen: !_getState().kpiPickerOpen }); setTimeout(_fixHeroDropdowns, 0); }
  function closeCollapsedPicker() { _setState({ kpiPickerOpen: false }); }

  // Reposition hero card dropdowns to stay within the viewport
  function _fixHeroDropdowns() {
    function _clampDd(ddId, anchorSel, menuW, rightAlign) {
      const dd  = document.getElementById(ddId);
      const btn = document.querySelector(anchorSel);
      if (!dd || !btn) return;
      const rect = btn.getBoundingClientRect();
      const w = Math.min(menuW, window.innerWidth - 16);
      let left = rightAlign ? (rect.right - w) : rect.left;
      if (left < 8) left = 8;
      if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8;
      dd.style.position  = 'fixed';
      dd.style.left      = left + 'px';
      dd.style.right     = 'auto';
      dd.style.top       = (rect.bottom + 6) + 'px';
      dd.style.width     = w + 'px';
      dd.style.transform = 'none';
    }
    _clampDd('hero-menu-dropdown',   '#hero-menu-wrap button',   230, true);
    _clampDd('hero-picker-dropdown', '#hero-picker-wrap button', 320, false);
    // Collapsed picker: centre on the "+" button
    const collDd = document.getElementById('hero-collapsed-picker');
    const collBtn = document.querySelector('.overview-hero-collapsed button');
    if (collDd && collBtn) {
      const rect = collBtn.getBoundingClientRect();
      const w = Math.min(300, window.innerWidth - 16);
      let left = rect.left + rect.width / 2 - w / 2;
      if (left < 8) left = 8;
      if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8;
      collDd.style.position  = 'fixed';
      collDd.style.left      = left + 'px';
      collDd.style.right     = 'auto';
      collDd.style.transform = 'none';
      collDd.style.top       = (rect.bottom + 8) + 'px';
      collDd.style.width     = w + 'px';
    }
  }

  // ── Per-card visualization type ───────────────────────────────────────────
  function setCardVizType(kpiId, type) {
    const st  = _getState();
    const map = { ...(st.cardVizTypes || {}) };
    map[kpiId] = type;
    _closeCardMenuPortal();
    _setState({ cardVizTypes: map, cardMenuOpen: null });
  }

  function _closeCardMenuPortal() {
    const existing = document.getElementById('_card-viz-portal');
    if (existing) existing.remove();
    document.querySelectorAll('.card-viz-menu-btn.open').forEach(b => b.classList.remove('open'));
  }

  function toggleCardMenu(kpiId, btnEl) {
    const existing = document.getElementById('_card-viz-portal');
    if (existing) {
      const wasKpi = existing.dataset.forKpi === kpiId;
      _closeCardMenuPortal();
      if (wasKpi) return;
    }

    const btn = btnEl || document.querySelector(`[data-kpi-id="${kpiId}"].card-viz-menu-btn`);
    if (!btn) return;
    btn.classList.add('open');

    const st = _getState();
    const cardVizTypes = st.cardVizTypes || {};
    const vizType = cardVizTypes[kpiId] || 'card';

    const vizOptions = [
      { v:'card',     l:'KPI Card',    i:'⊡' },
      { v:'odometer', l:'Odometer',    i:'◎' },
      { v:'bar',      l:'Bar Chart',   i:'▊' },
      { v:'line',     l:'Line Chart',  i:'∿' },
    ];

    const rect = btn.getBoundingClientRect();
    const portal = document.createElement('div');
    portal.id = '_card-viz-portal';
    portal.dataset.forKpi = kpiId;

    // Viewport-safe positioning
    const menuW = 168, menuH = 190;
    const spaceRight = window.innerWidth - rect.right;
    const spaceLeft  = rect.left;
    let hPos;
    if (spaceRight >= menuW + 8 || spaceRight >= spaceLeft) {
      // Align right edge of menu to right edge of button, clamped
      hPos = `right:${Math.max(8, spaceRight)}px`;
    } else {
      // Align left edge of menu to left edge of button, clamped
      hPos = `left:${Math.max(8, Math.min(rect.left, window.innerWidth - menuW - 8))}px`;
    }
    const topVal  = (rect.bottom + 4 + menuH > window.innerHeight - 8)
      ? Math.max(8, rect.top - menuH - 4)
      : rect.bottom + 4;

    portal.style.cssText = [
      'position:fixed',
      `top:${topVal}px`,
      hPos,
      'width:168px',
      'background:var(--bg-modal)',
      'border:1px solid var(--border-card)',
      'border-radius:8px',
      'box-shadow:var(--shadow-modal)',
      'z-index:99999',
      'overflow:hidden',
    ].join(';');

    portal.innerHTML = `
      <div style="padding:7px 12px;border-bottom:1px solid var(--border-subtle);font-size:10px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted)">Visualization</div>
      ${vizOptions.map(o => `
        <div class="card-viz-option ${vizType===o.v?'active':''}"
             onclick="event.stopPropagation();OverviewPanel.setCardVizType('${kpiId}','${o.v}')">
          <span style="font-size:14px;width:16px;text-align:center">${o.i}</span>
          <span>${o.l}</span>
          ${vizType===o.v?'<span style="margin-left:auto;font-size:10px;color:var(--brand-accent)">✓</span>':''}
        </div>`).join('')}`;

    document.body.appendChild(portal);

    const outsideHandler = (e) => {
      if (!portal.contains(e.target) && !btn.contains(e.target)) {
        _closeCardMenuPortal();
        document.removeEventListener('click', outsideHandler, true);
      }
    };
    setTimeout(() => document.addEventListener('click', outsideHandler, true), 0);
  }

  function closeCardMenu() {
    _closeCardMenuPortal();
  }

  function toggleHeroKpi(id, checked) {
    const st    = _getState();
    let ids     = [...(st.heroKpiIds || [])];
    let order   = [...(st.heroKpiOrder?.length ? st.heroKpiOrder : ids)];
    if (checked && !ids.includes(id))  { ids.push(id); order.push(id); }
    if (!checked) { ids = ids.filter(x => x !== id); order = order.filter(x => x !== id); }
    _setState({ heroKpiIds: ids, heroKpiOrder: order, kpiPickerOpen: true });
  }

  function setHeroAllKpis(selected) {
    const all = DataStore.getKpis();
    const ids = selected ? all.map(k => k.id) : [];
    _setState({ heroKpiIds: ids, heroKpiOrder: ids, kpiPickerOpen: true });
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

  // ── Drag-and-drop handlers (reorder within Key Metrics grid) ─────────────
  function dndStart(id, ev) {
    _dragId = id;
    _dragOverId = null;
    ev.dataTransfer.effectAllowed = 'move';
    ev.dataTransfer.setData('text/plain', id);
    setTimeout(() => {
      document.querySelector(`[data-dnd-id="${id}"]`)?.classList.add('dnd-dragging');
    }, 0);
  }

  function dndEnd() {
    document.querySelectorAll('.dnd-dragging').forEach(el => el.classList.remove('dnd-dragging'));
    document.querySelectorAll('.dnd-drop-before').forEach(el => el.classList.remove('dnd-drop-before'));
    document.getElementById('dnd-zone-grid')?.classList.remove('dnd-zone-active');
    _dragId = null; _dragOverId = null;
  }

  function dndOver(targetId, ev) {
    ev.preventDefault();
    ev.dataTransfer.dropEffect = 'move';
    if (_dragOverId === targetId) return;
    document.querySelector(`[data-dnd-id="${_dragOverId}"]`)?.classList.remove('dnd-drop-before');
    const gridEl = document.getElementById('dnd-zone-grid');
    _dragOverId = targetId;
    if (targetId) {
      document.querySelector(`[data-dnd-id="${targetId}"]`)?.classList.add('dnd-drop-before');
      gridEl?.classList.remove('dnd-zone-active');
    } else {
      gridEl?.classList.add('dnd-zone-active');
    }
  }

  function dndDrop(targetId, ev) {
    ev.preventDefault();
    ev.stopPropagation();
    const draggedId = _dragId;
    dndEnd();
    if (!draggedId || !targetId || draggedId === targetId) return;

    const s = DataStore.getSettings();
    let ids = s.overviewKpiIds
      ? [...s.overviewKpiIds]
      : DataStore.getOverviewKpis().map(k => k.id);

    const fromIdx = ids.indexOf(draggedId);
    const toIdx   = ids.indexOf(targetId);
    if (fromIdx === -1 || toIdx === -1) return;
    // Swap the two positions
    [ids[fromIdx], ids[toIdx]] = [ids[toIdx], ids[fromIdx]];
    DataStore.setOverviewKpiIds(ids);
  }

  function moveHeroKpi(id, dir) {
    const st    = _getState();
    const ids   = [...(st.heroKpiIds || [])];
    const order = [...(st.heroKpiOrder?.length ? st.heroKpiOrder : ids)];
    const i     = order.indexOf(id);
    if (i < 0) return;
    const j = dir === 'left' ? i - 1 : i + 1;
    if (j < 0 || j >= order.length) return;
    [order[i], order[j]] = [order[j], order[i]];
    _setState({ heroKpiOrder: order });
  }

  function openStackPicker(primaryId) {
    const st      = _getState();
    const allKpis = DataStore.getKpis();
    const primary = allKpis.find(k => k.id === primaryId);
    if (!primary) return;
    const curStack  = (st.stackedKpiIds || {})[primaryId];
    const available = allKpis.filter(k => k.id !== primaryId);

    const items = available.length === 0
      ? '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No other KPIs available</div>'
      : available.map(k => `
          <div class="hero-kpi-item" style="cursor:pointer" onclick="OverviewPanel.setStack('${primaryId}','${k.id}')">
            <div style="flex:1;min-width:0">
              <div class="hki-name">${_esc(k.metric)}</div>
              <div class="hki-section">${_esc(k.section)}</div>
            </div>
            ${curStack === k.id ? '<span style="color:var(--brand-accent);font-size:14px">✓</span>' : ''}
          </div>`).join('');

    App.showModal(`
      <div class="modal-overlay" onclick="if(event.target===this)App.closeModal()">
        <div class="modal" style="max-width:360px">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
            <div>
              <div style="font-family:var(--font-display);font-size:15px;font-weight:700">Stack a KPI</div>
              <div style="font-size:11px;color:var(--text-muted);margin-top:2px">Stacked below: ${_esc(primary.metric)}</div>
            </div>
            <button onclick="App.closeModal()" style="color:var(--text-muted);font-size:18px;background:none;border:none;cursor:pointer">✕</button>
          </div>
          <div id="stack-picker-list" style="max-height:280px;overflow-y:auto;border:1px solid var(--border-subtle);border-radius:8px;background:var(--bg-card)">
            ${items}
          </div>
          ${curStack ? `
          <button onclick="OverviewPanel.setStack('${primaryId}',null)" class="btn btn-ghost"
                  style="width:100%;margin-top:10px;font-size:12px;color:var(--rag-red)">
            ✕ Remove stacked KPI
          </button>` : ''}
        </div>
      </div>`);
  }

  function setStack(primaryId, stackedId) {
    const st  = _getState();
    const map = { ...(st.stackedKpiIds || {}) };
    if (!stackedId) delete map[primaryId];
    else map[primaryId] = stackedId;
    _setState({ stackedKpiIds: map });
    App.closeModal();
  }

  function _filterStackList(q) {
    const list = document.getElementById('stack-picker-list');
    if (!list) return;
    const query = q.toLowerCase().trim();
    list.querySelectorAll('.hero-kpi-item').forEach(item => {
      const name = item.querySelector('.hki-name')?.textContent?.toLowerCase() || '';
      item.style.display = (!query || name.includes(query)) ? '' : 'none';
    });
  }

  // ── Mount dispatcher ──────────────────────────────────────────────────────
  function mountCharts() {
    const st         = _getState();
    const allKpis    = DataStore.getKpis();
    const heroKpiIds = st.heroKpiIds;
    const rawKpis    = heroKpiIds === undefined || heroKpiIds === null
      ? DataStore.getKeyKpis()
      : heroKpiIds.length > 0
        ? allKpis.filter(k => heroKpiIds.includes(k.id))
        : [];

    const order = st.heroKpiOrder || [];
    const kpis  = order.length
      ? [...rawKpis].sort((a, b) => {
          const ai = order.indexOf(a.id), bi = order.indexOf(b.id);
          return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi);
        })
      : rawKpis;

    const settings = DataStore.getSettings();
    const mode     = settings.reportingPeriod || 'monthly';
    const stacks   = st.stackedKpiIds || {};

    if (kpis.length > 0) {
      switch (st.chartType) {
        case 'odometer': _mountOdometers(kpis, mode, settings, stacks, allKpis); break;
        case 'bar':      _mountBar(kpis, mode);       break;
        case 'line':     _mountLine(kpis, mode);      break;
      }
    }

    // Mount per-card mini visualizations
    _mountCardVizs(st, allKpis, mode);
  }

  // ── Per-card mini visualization mounter ──────────────────────────────────
  function _mountCardVizs(st, allKpis, mode) {
    const cardVizTypes = st.cardVizTypes || {};
    Object.entries(cardVizTypes).forEach(([kpiId, vizType]) => {
      if (vizType === 'card') return; // standard card, nothing to mount
      const container = document.getElementById(`card-viz-${kpiId}`);
      if (!container) return;
      const kpi = allKpis.find(k => k.id === kpiId);
      if (!kpi) return;

      switch (vizType) {
        case 'odometer': _mountCardOdometer(container, kpi, mode); break;
        case 'bar':      _mountCardBar(container, kpi, mode);      break;
        case 'line':     _mountCardLine(container, kpi, mode);     break;
      }
    });
  }

  // ── Helpers ───────────────────────────────────────────────────────────────
  function _cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
  const RAG_COLOR = {
    green:   () => _cssVar('--rag-green')   || '#00C48C',
    amber:   () => _cssVar('--rag-amber')   || '#FF9F1C',
    red:     () => _cssVar('--rag-red')     || '#FF4444',
    neutral: () => _cssVar('--rag-neutral') || '#6B7A99',
  };
  function _fmt(v, kpi)  { return v === null || v === undefined ? '—' : DataStore.formatValue(v, kpi); }
  function _shortNum(n)  {
    const abs = Math.abs(n);
    if (abs >= 1e9) return (n/1e9).toFixed(1)+'B';
    if (abs >= 1e6) return (n/1e6).toFixed(1)+'M';
    if (abs >= 1e3) return (n/1e3).toFixed(0)+'K';
    return Math.round(n).toString();
  }
  function _esc(s) {
    return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // ── Arc geometry helpers ──────────────────────────────────────────────────
  function _arcPath(cx, cy, r) {
    const pathD = `M ${cx - r} ${cy} A ${r} ${r} 0 0 1 ${cx + r} ${cy}`;
    const circ  = Math.PI * r;
    return { pathD, circ };
  }

  // ── CARD MINI: Odometer ───────────────────────────────────────────────────
  function _mountCardOdometer(container, kpi, mode) {
    const ps     = DataStore.getPeriodStats(kpi, mode);
    const actual = ps.actual ?? 0;
    const target = ps.target ?? (kpi.targetFY26 ?? 1);
    const rag    = ps.rag || kpi.rag || 'neutral';
    const color  = RAG_COLOR[rag]?.() || '#6B7A99';
    const maxV   = Math.max(1, target) / 0.75;
    const pct    = Math.min(1, Math.max(0, actual / maxV));

    // Compact layout: value row + SVG gauge, total ~130px
    const cx = 80, cy = 72, r = 58, sw = 13;
    const svgW = 160, svgH = 74;
    const { pathD, circ } = _arcPath(cx, cy, r);

    const tRad = Math.PI * 0.25;
    const tmx1 = cx + r * Math.cos(tRad),      tmy1 = cy - r * Math.sin(tRad);
    const tmx2 = cx + (r+8) * Math.cos(tRad),  tmy2 = cy - (r+8) * Math.sin(tRad);
    const tlx  = cx + (r+17) * Math.cos(tRad), tly  = cy - (r+17) * Math.sin(tRad);
    const ragBgMap = { green:'rgba(0,196,140,0.15)', amber:'rgba(255,159,28,0.15)', red:'rgba(255,68,68,0.15)', neutral:'rgba(107,122,153,0.15)' };

    // RAG + target text sit just below the arc baseline inside SVG
    const labelH = 32;
    const ragY   = cy + 10;

    container.innerHTML = `
      <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:2px;padding:0 2px">
        <div style="font-family:var(--font-display);font-size:1.15rem;font-weight:700;color:${color};line-height:1">${_fmt(ps.actual, kpi)}</div>
        <div style="font-size:9px;color:var(--text-muted)">/ ${_fmt(target, kpi)}</div>
      </div>
      <svg viewBox="0 0 ${svgW} ${svgH + labelH}" style="width:100%;display:block;flex:1;min-height:0;overflow:hidden">
        <path d="${pathD}" fill="none" stroke="rgba(255,255,255,0.05)" stroke-width="${sw}" stroke-linecap="round"/>
        <path d="${pathD}" fill="none" stroke="rgba(255,255,255,0.07)" stroke-width="${sw+4}" stroke-dasharray="2 8" stroke-linecap="butt"/>
        <path id="codo-fill-${kpi.id}" d="${pathD}" fill="none" stroke="${color}" stroke-width="${sw}" stroke-linecap="round"
              stroke-dasharray="${circ}" stroke-dashoffset="${circ}"
              style="transition:stroke-dashoffset 1s cubic-bezier(0.4,0,0.2,1);filter:drop-shadow(0 0 4px ${color}66)"/>
        <line x1="${tmx1.toFixed(1)}" y1="${tmy1.toFixed(1)}" x2="${tmx2.toFixed(1)}" y2="${tmy2.toFixed(1)}"
              stroke="rgba(255,255,255,0.55)" stroke-width="1.5" stroke-linecap="round"/>
        <rect x="${(tlx-10).toFixed(1)}" y="${(tly-5).toFixed(1)}" width="20" height="10" rx="2" fill="rgba(0,0,0,0.5)"/>
        <text x="${tlx.toFixed(1)}" y="${tly.toFixed(1)}" fill="rgba(255,255,255,0.9)" font-size="7" font-weight="700" text-anchor="middle" dominant-baseline="middle">${_shortNum(target)}</text>
        <line id="codo-needle-${kpi.id}" x1="${cx}" y1="${cy}" x2="${cx - (r+2)}" y2="${cy}"
              stroke="rgba(255,255,255,0.9)" stroke-width="2" stroke-linecap="round"
              style="transform-origin:${cx}px ${cy}px;transition:transform 1s cubic-bezier(0.4,0,0.2,1)"/>
        <circle cx="${cx}" cy="${cy}" r="6" fill="var(--bg-page)" stroke="rgba(255,255,255,0.18)" stroke-width="1.5"/>
        <circle cx="${cx}" cy="${cy}" r="3" fill="${color}" opacity="0.9"/>
        <rect x="${(cx-16).toFixed(1)}" y="${(ragY-5).toFixed(1)}" width="32" height="10" rx="2" fill="${ragBgMap[rag]||ragBgMap.neutral}"/>
        <text x="${cx}" y="${ragY.toFixed(1)}" fill="${color}" font-size="7" font-weight="700" text-anchor="middle" dominant-baseline="middle" letter-spacing="0.08em">${rag.toUpperCase()}</text>
      </svg>`;

    requestAnimationFrame(() => requestAnimationFrame(() => {
      const fill   = document.getElementById(`codo-fill-${kpi.id}`);
      const needle = document.getElementById(`codo-needle-${kpi.id}`);
      if (fill)   fill.style.strokeDashoffset = circ * (1 - pct);
      if (needle) needle.style.transform       = `rotate(${pct * 180}deg)`;
    }));
  }

  // ── CARD MINI: Bar Chart ──────────────────────────────────────────────────
  function _mountCardBar(container, kpi, mode) {
    const fyMonths = DataStore.getFyMonths();
    const ma       = kpi.monthlyActuals || {};
    const ps       = DataStore.getPeriodStats(kpi, mode);
    const rag      = ps.rag || kpi.rag || 'neutral';
    const color    = RAG_COLOR[rag]?.() || '#6B7A99';
    const tgtMonth = kpi.targetMonth ?? (kpi.targetFY26 ? kpi.targetFY26 / 12 : null);

    const vals    = fyMonths.map(m => ma[m] !== undefined ? parseFloat(ma[m]) : null);
    const hasData = vals.some(v => v !== null);

    if (!hasData) {
      container.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:4px">
          <div style="font-family:var(--font-display);font-size:1.15rem;font-weight:700;color:${color}">${_fmt(ps.actual, kpi)}</div>
          <div style="font-size:10px;color:var(--text-muted)">No monthly data yet</div>
        </div>`;
      return;
    }

    // Compact: value+target row (20px) + SVG (fills remaining ~110px)
    const W = 240, H = 96;
    const pad = { top:4, right:6, bottom:18, left:24 };
    const cW = W - pad.left - pad.right;
    const cH = H - pad.top - pad.bottom;
    const maxV  = Math.max(...vals.filter(v=>v!==null), tgtMonth || 0, 1);
    const barW  = Math.max(3, (cW / fyMonths.length) - 2);
    const xStep = cW / fyMonths.length;
    const yS    = v => cH - (v / maxV) * cH;

    const tLine = tgtMonth !== null ? (() => {
      const ty = pad.top + yS(tgtMonth);
      return `<line x1="${pad.left}" y1="${ty.toFixed(1)}" x2="${W-pad.right}" y2="${ty.toFixed(1)}"
                    stroke="${color}" stroke-width="1.5" stroke-dasharray="4 3" opacity="0.6"/>`;
    })() : '';

    const bars = fyMonths.map((m, i) => {
      const v = vals[i];
      if (v === null) return '';
      const bH = Math.max(2, (v / maxV) * cH);
      const bX = pad.left + i * xStep + (xStep - barW) / 2;
      const bY = pad.top + yS(v);
      return `<rect x="${bX.toFixed(1)}" y="${bY.toFixed(1)}" width="${barW.toFixed(1)}" height="${bH.toFixed(1)}"
                    rx="2" fill="${color}" opacity="0.8"><title>${m}: ${_fmt(v, kpi)}</title></rect>`;
    }).join('');

    const xLabels = fyMonths.map((m, i) => {
      if (i % 3 !== 0) return '';
      const x = pad.left + i * xStep + xStep / 2;
      return `<text x="${x.toFixed(1)}" y="${(H-3).toFixed(1)}" fill="rgba(255,255,255,0.3)" font-size="6" text-anchor="middle">${m}</text>`;
    }).join('');

    const yTicks = [0, 1].map(f => {
      const v = maxV * f;
      const y = pad.top + yS(v);
      return `<text x="${(pad.left-2).toFixed(1)}" y="${(y+3).toFixed(1)}" fill="rgba(255,255,255,0.25)" font-size="6" text-anchor="end">${_shortNum(v)}</text>
              <line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${W-pad.right}" y2="${y.toFixed(1)}" stroke="rgba(255,255,255,0.04)" stroke-width="1"/>`;
    }).join('');

    container.innerHTML = `
      <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:2px;padding:0 2px;flex-shrink:0">
        <div style="font-family:var(--font-display);font-size:1.15rem;font-weight:700;color:${color}">${_fmt(ps.actual, kpi)}</div>
        ${tgtMonth !== null ? `<div style="font-size:9px;color:var(--text-muted)">Target ${_fmt(tgtMonth, kpi)}/mo</div>` : ''}
      </div>
      <svg viewBox="0 0 ${W} ${H}" style="width:100%;flex:1;min-height:0;display:block;overflow:hidden">
        ${yTicks}${tLine}${bars}${xLabels}
      </svg>`;
  }

  // ── CARD MINI: Line Chart ─────────────────────────────────────────────────
  function _mountCardLine(container, kpi, mode) {
    const fyMonths = DataStore.getFyMonths();
    const ma       = kpi.monthlyActuals || {};
    const ps       = DataStore.getPeriodStats(kpi, mode);
    const rag      = ps.rag || kpi.rag || 'neutral';
    const color    = RAG_COLOR[rag]?.() || '#6B7A99';
    const tgtMonth = kpi.targetMonth ?? (kpi.targetFY26 ? kpi.targetFY26 / 12 : null);

    const vals   = fyMonths.map(m => ma[m] !== undefined ? parseFloat(ma[m]) : null);
    const points = fyMonths.map((m, i) => ({ m, v: vals[i], i })).filter(p => p.v !== null);

    if (points.length < 2) {
      container.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:4px">
          <div style="font-family:var(--font-display);font-size:1.15rem;font-weight:700;color:${color}">${_fmt(ps.actual, kpi)}</div>
          <div style="font-size:10px;color:var(--text-muted)">Need ≥2 months of data</div>
        </div>`;
      return;
    }

    const W = 240, H = 96;
    const pad  = { top:4, right:6, bottom:18, left:24 };
    const cW   = W - pad.left - pad.right;
    const cH   = H - pad.top - pad.bottom;
    const allV = [...points.map(p => p.v), tgtMonth !== null ? tgtMonth : 0];
    const minV = Math.min(0, ...allV);
    const maxV = Math.max(...allV, 1);
    const range = maxV - minV || 1;
    const xStep = cW / (fyMonths.length - 1);
    const yS = v => cH - ((v - minV) / range) * cH;

    const coordPts = points.map(p => ({ x: pad.left + p.i * xStep, y: pad.top + yS(p.v) }));
    const lineD = coordPts.map((p, i) => `${i===0?'M':'L'} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(' ');
    const areaD = coordPts.length > 1
      ? `${lineD} L ${coordPts[coordPts.length-1].x.toFixed(1)} ${(pad.top+cH).toFixed(1)} L ${coordPts[0].x.toFixed(1)} ${(pad.top+cH).toFixed(1)} Z`
      : '';

    const tLine = tgtMonth !== null ? (() => {
      const ty = pad.top + yS(tgtMonth);
      return `<line x1="${pad.left}" y1="${ty.toFixed(1)}" x2="${W-pad.right}" y2="${ty.toFixed(1)}"
                    stroke="${color}" stroke-width="1.5" stroke-dasharray="4 3" opacity="0.5"/>`;
    })() : '';

    const xLabels = fyMonths.map((m, i) => {
      if (i % 3 !== 0) return '';
      const x = pad.left + i * xStep;
      return `<text x="${x.toFixed(1)}" y="${(H-3).toFixed(1)}" fill="rgba(255,255,255,0.3)" font-size="6" text-anchor="middle">${m}</text>`;
    }).join('');

    const yTicks = [0, 1].map(f => {
      const v = minV + range * f;
      const y = pad.top + yS(v);
      return `<text x="${(pad.left-2).toFixed(1)}" y="${(y+3).toFixed(1)}" fill="rgba(255,255,255,0.25)" font-size="6" text-anchor="end">${_shortNum(v)}</text>
              <line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${W-pad.right}" y2="${y.toFixed(1)}" stroke="rgba(255,255,255,0.04)" stroke-width="1"/>`;
    }).join('');

    const dots = coordPts.map(p => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2" fill="${color}" opacity="0.9"/>`).join('');

    container.innerHTML = `
      <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:2px;padding:0 2px;flex-shrink:0">
        <div style="font-family:var(--font-display);font-size:1.15rem;font-weight:700;color:${color}">${_fmt(ps.actual, kpi)}</div>
        ${tgtMonth !== null ? `<div style="font-size:9px;color:var(--text-muted)">Target ${_fmt(tgtMonth, kpi)}/mo</div>` : ''}
      </div>
      <svg viewBox="0 0 ${W} ${H}" style="width:100%;flex:1;min-height:0;display:block;overflow:hidden">
        <defs>
          <linearGradient id="clg-${kpi.id}" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="${color}" stop-opacity="0.20"/>
            <stop offset="100%" stop-color="${color}" stop-opacity="0"/>
          </linearGradient>
        </defs>
        ${yTicks}
        ${areaD ? `<path d="${areaD}" fill="url(#clg-${kpi.id})"/>` : ''}
        <path d="${lineD}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        ${dots}
        ${tLine}${xLabels}
      </svg>`;
  }

  // ── HERO: Concentric gauge ────────────────────────────────────────────────
  function _buildConcentricGauge(primaryKpi, stackedKpi, mode) {
    // Fixed SVG canvas — arc perfectly centred, all elements inside viewBox, nothing clips
    const svgW = 280, svgH = 220;
    const cx = svgW / 2;  // 140 — horizontally centred
    const cy = 140;       // arc baseline Y — labels go below

    const rOuter = 98, rInner = rOuter - 28;
    const swOuter = 20,  swInner = 16;

    const psP     = DataStore.getPeriodStats(primaryKpi, mode);
    const actP    = psP.actual  ?? 0;
    const tgtP    = psP.target  ?? (primaryKpi.targetFY26 ?? 1);
    const ragP    = psP.rag || primaryKpi.rag || 'neutral';
    const colP    = RAG_COLOR[ragP]?.() || '#6B7A99';
    const maxP    = Math.max(1, tgtP) / 0.75;
    const pctP    = Math.min(1, Math.max(0, actP / maxP));
    const { pathD: outerPath, circ: circOuter } = _arcPath(cx, cy, rOuter);

    let actS = 0, tgtS = 0, ragS = 'neutral', colS = '#6B7A99', pctS = 0, circInner = 0, innerPath = '';
    if (stackedKpi) {
      const psS = DataStore.getPeriodStats(stackedKpi, mode);
      actS  = psS.actual ?? 0;
      tgtS  = psS.target ?? (stackedKpi.targetFY26 ?? 1);
      ragS  = psS.rag || stackedKpi.rag || 'neutral';
      colS  = RAG_COLOR[ragS]?.() || '#6B7A99';
      const maxS = Math.max(1, tgtS) / 0.75;
      pctS  = Math.min(1, Math.max(0, actS / maxS));
      const inner = _arcPath(cx, cy, rInner);
      circInner = inner.circ;
      innerPath = inner.pathD;
    }

    const needleLenP = rOuter + 4;
    const needleLenS = rInner - 4;
    const tRad = Math.PI * 0.25;
    const tmx1 = cx + rOuter * Math.cos(tRad),        tmy1 = cy - rOuter * Math.sin(tRad);
    const tmx2 = cx + (rOuter + 10) * Math.cos(tRad), tmy2 = cy - (rOuter + 10) * Math.sin(tRad);
    const tlx  = cx + (rOuter + 22) * Math.cos(tRad), tly  = cy - (rOuter + 22) * Math.sin(tRad);
    const tmix1 = cx + rInner * Math.cos(tRad),         tmiy1 = cy - rInner * Math.sin(tRad);
    const tmix2 = cx + (rInner - 10) * Math.cos(tRad), tmiy2 = cy - (rInner - 10) * Math.sin(tRad);
    const tlix  = cx + (rInner + 22) * Math.cos(tRad), tliy  = cy - (rInner + 22) * Math.sin(tRad);

    const hubR = 16, hubInnerR = 5;
    const ragColor = { green:'#00C48C', amber:'#FF9F1C', red:'#FF4444', neutral:'#6B7A99' };
    const ragBg    = { green:'rgba(0,196,140,0.15)', amber:'rgba(255,159,28,0.15)', red:'rgba(255,68,68,0.15)', neutral:'rgba(107,122,153,0.15)' };
    const colPRag = ragColor[ragP] || '#6B7A99';
    const bgPRag  = ragBg[ragP]   || 'rgba(107,122,153,0.15)';
    const colSRag = ragColor[ragS] || '#6B7A99';
    const bgSRag  = ragBg[ragS]   || 'rgba(107,122,153,0.15)';
    const ragW = 52, ragH = 16, ragRr = 4;
    const ragY1      = cy + hubR + 12;
    const ragY2name  = ragY1 + 22;
    const ragY2val   = ragY1 + 40;
    const ragY2badge = ragY1 + 56;

    const stackBtnSvg = `
      <g class="odo-hub-btn" onclick="OverviewPanel.openStackPicker('${primaryKpi.id}')"
         style="cursor:pointer" title="${stackedKpi ? 'Change stacked KPI' : 'Stack a KPI'}">
        <circle cx="${cx}" cy="${cy}" r="${hubR}" fill="var(--bg-page)"
                stroke="${stackedKpi ? colS : 'rgba(255,255,255,0.15)'}" stroke-width="1.5"
                class="odo-hub-ring"/>
        <circle cx="${cx}" cy="${cy}" r="${hubInnerR}" fill="${colP}" opacity="0.9"/>
        ${stackedKpi ? `<circle cx="${cx}" cy="${cy}" r="2.5" fill="${colS}" opacity="0.95"/>` : ''}
        <text x="${cx}" y="${cy}" fill="rgba(255,255,255,0.55)" font-size="11"
              font-weight="700" text-anchor="middle" dominant-baseline="middle"
              class="odo-hub-plus" style="pointer-events:none;user-select:none">⊕</text>
      </g>`;

    return `
      <div class="odo-block" id="odo-wrap-${primaryKpi.id}">
        <div style="text-align:center;margin-bottom:4px;flex-shrink:0;padding:0 8px">
          <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-secondary);margin-bottom:2px;white-space:normal;word-break:break-word;line-height:1.3">${_esc(primaryKpi.metric)}</div>
          <div style="font-family:var(--font-display);font-weight:700;font-size:1.8rem;color:${colP};letter-spacing:-0.5px;line-height:1" id="odo-val-primary-${primaryKpi.id}">${_fmt(psP.actual, primaryKpi)}</div>
        </div>
        <div style="flex:1;min-height:0;display:flex;align-items:center;justify-content:center">
          <svg viewBox="0 0 ${svgW} ${svgH}" style="width:100%;max-height:100%;overflow:visible;display:block">
            <path d="${outerPath}" fill="none" stroke="rgba(255,255,255,0.05)" stroke-width="${swOuter}" stroke-linecap="round"/>
            <path d="${outerPath}" fill="none" stroke="rgba(255,255,255,0.08)" stroke-width="${swOuter+4}" stroke-dasharray="2 10" stroke-linecap="butt"/>
            <path id="odo-fill-${primaryKpi.id}" d="${outerPath}" fill="none" stroke="${colP}" stroke-width="${swOuter}" stroke-linecap="round"
                  stroke-dasharray="${circOuter}" stroke-dashoffset="${circOuter}"
                  style="transition:stroke-dashoffset 1s cubic-bezier(0.4,0,0.2,1);filter:drop-shadow(0 0 6px ${colP}66)"/>
            <path id="odo-glow-${primaryKpi.id}" d="${outerPath}" fill="none" stroke="${colP}" stroke-width="3" stroke-linecap="round" opacity="0.15"
                  stroke-dasharray="${circOuter}" stroke-dashoffset="${circOuter}"
                  style="transition:stroke-dashoffset 1s cubic-bezier(0.4,0,0.2,1)"/>
            <line x1="${tmx1.toFixed(1)}" y1="${tmy1.toFixed(1)}" x2="${tmx2.toFixed(1)}" y2="${tmy2.toFixed(1)}"
                  stroke="rgba(255,255,255,0.7)" stroke-width="2.5" stroke-linecap="round"/>
            <rect x="${(tlx-14).toFixed(1)}" y="${(tly-7).toFixed(1)}" width="28" height="14" rx="3" fill="rgba(0,0,0,0.55)"/>
            <text x="${tlx.toFixed(1)}" y="${tly.toFixed(1)}" fill="rgba(255,255,255,0.95)" font-size="9" font-weight="700" text-anchor="middle" dominant-baseline="middle">${_shortNum(tgtP)}</text>
            ${stackedKpi ? `
            <path d="${innerPath}" fill="none" stroke="rgba(255,255,255,0.04)" stroke-width="${swInner}" stroke-linecap="round"/>
            <path d="${innerPath}" fill="none" stroke="rgba(255,255,255,0.07)" stroke-width="${swInner+4}" stroke-dasharray="2 9" stroke-linecap="butt"/>
            <path id="odo-fill-${stackedKpi.id}" d="${innerPath}" fill="none" stroke="${colS}" stroke-width="${swInner}" stroke-linecap="round"
                  stroke-dasharray="${circInner}" stroke-dashoffset="${circInner}"
                  style="transition:stroke-dashoffset 1s cubic-bezier(0.4,0,0.2,1);filter:drop-shadow(0 0 5px ${colS}55)"/>
            <line x1="${tmix1.toFixed(1)}" y1="${tmiy1.toFixed(1)}" x2="${tmix2.toFixed(1)}" y2="${tmiy2.toFixed(1)}"
                  stroke="rgba(255,255,255,0.6)" stroke-width="2" stroke-linecap="round"/>
            <rect x="${(tlix-13).toFixed(1)}" y="${(tliy-6).toFixed(1)}" width="26" height="12" rx="3" fill="rgba(0,0,0,0.55)"/>
            <text x="${tlix.toFixed(1)}" y="${tliy.toFixed(1)}" fill="rgba(255,255,255,0.95)" font-size="8" font-weight="700" text-anchor="middle" dominant-baseline="middle">${_shortNum(tgtS)}</text>
            ` : ''}
            <line id="odo-needle-${primaryKpi.id}" x1="${cx}" y1="${cy}" x2="${cx - needleLenP}" y2="${cy}"
                  stroke="rgba(255,255,255,0.92)" stroke-width="3" stroke-linecap="round"
                  style="transform-origin:${cx}px ${cy}px;transition:transform 1s cubic-bezier(0.4,0,0.2,1)"/>
            ${stackedKpi ? `
            <line id="odo-needle-stacked-${stackedKpi.id}" x1="${cx}" y1="${cy}" x2="${cx - needleLenS}" y2="${cy}"
                  stroke="${colS}" stroke-width="2.5" stroke-linecap="round" opacity="0.9"
                  style="transform-origin:${cx}px ${cy}px;transition:transform 1s cubic-bezier(0.4,0,0.2,1)"/>
            ` : ''}
            ${stackBtnSvg}
            <rect x="${(cx - ragW/2).toFixed(1)}" y="${(ragY1 - ragH/2).toFixed(1)}" width="${ragW}" height="${ragH}" rx="${ragRr}" fill="${bgPRag}"/>
            <text x="${cx}" y="${ragY1}" fill="${colPRag}" font-size="9" font-weight="700" text-anchor="middle" dominant-baseline="middle" letter-spacing="0.08em">${ragP.toUpperCase()}</text>
            ${stackedKpi ? `
            <text x="${cx}" y="${ragY2name.toFixed(1)}" fill="rgba(255,255,255,0.4)" font-size="8" font-weight="700" text-anchor="middle" dominant-baseline="middle" letter-spacing="0.08em">${_esc(stackedKpi.metric).toUpperCase().slice(0,22)}</text>
            <text x="${cx}" y="${ragY2val.toFixed(1)}" fill="${colS}" font-size="14" font-weight="700" text-anchor="middle" dominant-baseline="middle" letter-spacing="-0.3px">${_fmt(actS, stackedKpi)}</text>
            <rect x="${(cx - ragW/2).toFixed(1)}" y="${(ragY2badge - ragH/2).toFixed(1)}" width="${ragW}" height="${ragH}" rx="${ragRr}" fill="${bgSRag}"/>
            <text x="${cx}" y="${ragY2badge.toFixed(1)}" fill="${colSRag}" font-size="9" font-weight="700" text-anchor="middle" dominant-baseline="middle" letter-spacing="0.08em">${ragS.toUpperCase()}</text>
            ` : ''}
          </svg>
        </div>
      </div>`;
  }

  // ── HERO: Odometer mount ──────────────────────────────────────────────────
  function _mountOdometers(kpis, mode, settings, stacks, allKpis) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;
    const n = kpis.length;
    const cols = kpis.map((kpi, idx) => {
      const stackedId = stacks[kpi.id];
      const stacked   = stackedId ? allKpis.find(k => k.id === stackedId) : null;
      const canLeft   = idx > 0, canRight = idx < n - 1;
      return `
        <div class="odo-col" data-kpi-id="${kpi.id}">
          <div class="odo-reorder-btns">
            <button class="odo-move-btn" ${canLeft?'':'disabled'} onclick="OverviewPanel.moveHeroKpi('${kpi.id}','left')">◂</button>
            <span class="odo-pos-label">${idx+1}/${n}</span>
            <button class="odo-move-btn" ${canRight?'':'disabled'} onclick="OverviewPanel.moveHeroKpi('${kpi.id}','right')">▸</button>
          </div>
          ${_buildConcentricGauge(kpi, stacked, mode)}
        </div>`;
    }).join('');
    container.innerHTML = `<div class="odo-row">${cols}</div>`;
    requestAnimationFrame(() => requestAnimationFrame(() => {
      kpis.forEach(kpi => {
        const stackedId = stacks[kpi.id];
        const stacked   = stackedId ? allKpis.find(k => k.id === stackedId) : null;
        const psP    = DataStore.getPeriodStats(kpi, mode);
        const actP   = psP.actual ?? 0;
        const tgtP   = psP.target ?? (kpi.targetFY26 ?? 1);
        const maxP   = Math.max(1, tgtP) / 0.75;
        const pctP   = Math.min(1, Math.max(0, actP / maxP));
        const fillP  = document.getElementById(`odo-fill-${kpi.id}`);
        const glowP  = document.getElementById(`odo-glow-${kpi.id}`);
        const needle = document.getElementById(`odo-needle-${kpi.id}`);
        if (fillP)  { const c = parseFloat(fillP.getAttribute('stroke-dasharray')); fillP.style.strokeDashoffset = c * (1 - pctP); }
        if (glowP)  { const c = parseFloat(glowP.getAttribute('stroke-dasharray')); glowP.style.strokeDashoffset = c * (1 - pctP); }
        if (needle) needle.style.transform = `rotate(${pctP * 180}deg)`;
        if (stacked) {
          const psS = DataStore.getPeriodStats(stacked, mode);
          const actS = psS.actual ?? 0; const tgtS = psS.target ?? (stacked.targetFY26 ?? 1);
          const maxS = Math.max(1, tgtS) / 0.75; const pctS = Math.min(1, Math.max(0, actS / maxS));
          const fillS   = document.getElementById(`odo-fill-${stacked.id}`);
          const needleS = document.getElementById(`odo-needle-stacked-${stacked.id}`);
          if (fillS) { const c = parseFloat(fillS.getAttribute('stroke-dasharray')); fillS.style.strokeDashoffset = c * (1 - pctS); }
          if (needleS) needleS.style.transform = `rotate(${pctS * 180}deg)`;
        }
      });
    }));
  }

  // ── HERO: Bar Chart ───────────────────────────────────────────────────────
  function _mountBar(kpis, mode) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;
    const W = container.clientWidth || 700, H = 300;
    const pad = { top:24, right:20, bottom:76, left:52 };
    const cW = W - pad.left - pad.right, cH = H - pad.top - pad.bottom;
    const data = kpis.map(kpi => {
      const ps = DataStore.getPeriodStats(kpi, mode);
      // Split label into up to 2 lines of ~16 chars each
      const words = kpi.metric.split(' ');
      const lines = []; let cur = '';
      words.forEach(w => {
        if ((cur + (cur?' ':'')+w).length <= 16) { cur += (cur?' ':'')+w; }
        else { if (cur) lines.push(cur); cur = w.length>16?w.slice(0,15)+'…':w; }
      });
      if (cur) lines.push(cur);
      return { lines: lines.slice(0,2), actual: ps.actual??0, target: ps.target??kpi.targetFY26??0, rag: ps.rag||kpi.rag||'neutral' };
    });
    const maxVal = Math.max(...data.flatMap(d=>[d.actual,d.target]), 1);
    const groupW = cW / data.length, barW = Math.min(44, groupW * 0.5);
    const yS = v => cH - (v / maxVal) * cH;
    const ticks = Array.from({length:6},(_,i)=>{
      const v=(maxVal/5)*i, y=pad.top+yS(v);
      return `<line x1="${pad.left}" y1="${y}" x2="${W-pad.right}" y2="${y}" stroke="rgba(255,255,255,0.05)" stroke-width="1"/>
              <text x="${pad.left-6}" y="${y+4}" fill="rgba(255,255,255,0.3)" font-size="10" text-anchor="end">${_shortNum(v)}</text>`;
    }).join('');
    const bars = data.map((d,i)=>{
      const cx=pad.left+groupW*i+groupW/2, color=RAG_COLOR[d.rag]?.()||'#6B7A99';
      const aH=Math.max(2,(d.actual/maxVal)*cH), ay=pad.top+yS(d.actual), ty=pad.top+yS(d.target);
      const labelY = pad.top+cH+16;
      const labelLines = d.lines.map((ln, li) =>
        `<text x="${cx.toFixed(1)}" y="${(labelY + li*13).toFixed(1)}" fill="rgba(255,255,255,0.55)" font-size="10" text-anchor="middle">${_esc(ln)}</text>`
      ).join('');
      return `<rect x="${(cx-barW/2).toFixed(1)}" y="${ay.toFixed(1)}" width="${barW}" height="${aH.toFixed(1)}" rx="3" fill="${color}" opacity="0.85"/>
              <rect x="${(cx-barW*0.8).toFixed(1)}" y="${(ty-1.5).toFixed(1)}" width="${(barW*1.6).toFixed(1)}" height="3" rx="1.5" fill="rgba(255,255,255,0.6)"/>
              ${labelLines}`;
    }).join('');
    container.innerHTML = `<svg width="100%" height="${H}" viewBox="0 0 ${W} ${H}" style="overflow:visible">${ticks}${bars}</svg>`;
  }

  // ── HERO: Line Chart ──────────────────────────────────────────────────────
  function _mountLine(kpis, mode) {
    const container = document.getElementById('hero-chart-area');
    if (!container) return;
    const fyMonths = DataStore.getFyMonths();
    const W = container.clientWidth || 700, H = 260;
    const pad = { top:28, right:90, bottom:52, left:58 };
    const cW = W - pad.left - pad.right, cH = H - pad.top - pad.bottom;
    const allVals = kpis.flatMap(kpi => {
      const mvs = fyMonths.map(m=>kpi.monthlyActuals?.[m]).filter(v=>v!=null).map(Number);
      const tgt = kpi.targetMonth??kpi.targetFY26;
      return tgt!=null?[...mvs,Number(tgt)]:mvs;
    });
    if (allVals.length === 0) {
      container.innerHTML = `<div style="text-align:center;padding:48px;color:var(--text-muted);font-size:13px">No monthly data yet</div>`;
      return;
    }
    const minVal=Math.min(0,...allVals), maxVal=Math.max(...allVals,1), range=maxVal-minVal||1;
    const yS = v => cH - ((v-minVal)/range)*cH;
    const xStep = cW / Math.max(fyMonths.length-1,1);
    const grid = Array.from({length:6},(_,i)=>{
      const v=minVal+(range/5)*i, y=pad.top+yS(v);
      return `<line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${W-pad.right}" y2="${y.toFixed(1)}" stroke="rgba(255,255,255,0.05)" stroke-width="1"/>
              <text x="${(pad.left-8).toFixed(1)}" y="${(y+4).toFixed(1)}" fill="rgba(255,255,255,0.3)" font-size="10" text-anchor="end">${_shortNum(v)}</text>`;
    }).join('');
    const xLabels = fyMonths.map((m,i)=>{
      const x=pad.left+i*xStep;
      return `<text x="${x.toFixed(1)}" y="${(pad.top+cH+16).toFixed(1)}" fill="rgba(255,255,255,0.35)" font-size="10" text-anchor="middle">${m}</text>`;
    }).join('');
    const PALETTE = ['#00C2A8','#3A86FF','#FF9F1C','#FF4444','#A78BFA','#34D399'];
    const series = kpis.map((kpi,ki)=>{
      const color=PALETTE[ki%PALETTE.length];
      const points=fyMonths.map((m,i)=>{const v=kpi.monthlyActuals?.[m];return v!=null?{x:pad.left+i*xStep,y:pad.top+yS(Number(v))}:null;}).filter(Boolean);
      const pathD=points.length>1?points.map((p,i)=>`${i===0?'M':'L'} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(' '):'';
      const areaD=pathD?`${pathD} L ${points[points.length-1].x.toFixed(1)} ${(pad.top+cH).toFixed(1)} L ${points[0].x.toFixed(1)} ${(pad.top+cH).toFixed(1)} Z`:'';
      const tgtVal=kpi.targetMonth??kpi.targetFY26;
      let targetSvg='';
      if(tgtVal!=null){const ty=pad.top+yS(Number(tgtVal));const tLabel=DataStore.formatValue(Number(tgtVal),kpi);targetSvg=`<line x1="${pad.left}" y1="${ty.toFixed(1)}" x2="${(W-pad.right).toFixed(1)}" y2="${ty.toFixed(1)}" stroke="${color}" stroke-width="1.5" stroke-dasharray="6 4" opacity="0.65"/><text x="${(W-pad.right+9).toFixed(1)}" y="${(ty+5).toFixed(1)}" fill="${color}" font-size="10" font-weight="600" opacity="0.95">${_esc(tLabel)}</text>`;}
      const dots=points.map(p=>`<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="3.5" fill="${color}" opacity="0.9"/>`).join('');
      return `<defs><linearGradient id="lg-${ki}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${color}" stop-opacity="0.18"/><stop offset="100%" stop-color="${color}" stop-opacity="0"/></linearGradient></defs>${areaD?`<path d="${areaD}" fill="url(#lg-${ki})"/>`:''}${pathD?`<path d="${pathD}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>`:''} ${dots}${targetSvg}`;
    }).join('');
    const legend = kpis.map((kpi, ki) => {
      const color = PALETTE[ki % PALETTE.length];
      // Split into up to 2 lines matching bar chart logic
      const words = kpi.metric.split(' ');
      const lines = []; let cur = '';
      words.forEach(w => {
        if ((cur + (cur?' ':'')+w).length <= 16) { cur += (cur?' ':'')+w; }
        else { if (cur) lines.push(cur); cur = w.length>16?w.slice(0,15)+'…':w; }
      });
      if (cur) lines.push(cur);
      const linesHtml = lines.slice(0,2).map(ln => `<div>${_esc(ln)}</div>`).join('');
      return `<div style="display:flex;align-items:flex-start;gap:5px;font-size:11px;color:rgba(255,255,255,0.6);min-width:0;flex-shrink:1">
        <div style="width:14px;height:3px;border-radius:2px;background:${color};flex-shrink:0;margin-top:5px"></div>
        <div style="line-height:1.4">${linesHtml}</div>
      </div>`;
    }).join('');
    container.innerHTML = `
      <div style="display:flex;flex-direction:column;width:100%;height:100%">
        <svg width="100%" height="${H}" viewBox="0 0 ${W} ${H}" style="overflow:visible;flex-shrink:0">${grid}${xLabels}${series}</svg>
        <div style="display:flex;flex-wrap:wrap;gap:8px 16px;padding:4px 8px 2px;justify-content:center">${legend}</div>
      </div>`;
  }

  // ── RENDER ────────────────────────────────────────────────────────────────
  function render() {
    const settings     = DataStore.getSettings();
    const overviewKpis = DataStore.getOverviewKpis();
    const allKpis      = DataStore.getKpis();
    const sections     = DataStore.getSections();
    const st           = _getState();
    const mode         = settings.reportingPeriod || 'monthly';

    const ragCounts = {green:0,amber:0,red:0,neutral:0};
    allKpis.forEach(k => {
      const ps = DataStore.getPeriodStats(k, mode);
      ragCounts[ps.rag || k.rag || 'neutral']++;
    });

    const heroKpiIds = st.heroKpiIds;
    const heroKpis   = heroKpiIds === undefined || heroKpiIds === null
      ? DataStore.getKeyKpis()
      : heroKpiIds.length > 0
        ? allKpis.filter(k => heroKpiIds.includes(k.id))
        : [];

    const header   = _renderHeader(settings, overviewKpis, allKpis);
    const ragStrip = _renderRagStrip(ragCounts, allKpis.length);

    const isCollapsed = heroKpis.length === 0;

    // All overview KPI cards render uniformly — the grid auto-places them
    // Hero takes cols 1-3 × rows 1-2 (expanded) or 1×1 (collapsed)
    // KPI cards are always 1 col × 1 row = 200px, auto-placed by CSS grid
    const allKpiCards = overviewKpis.map(kpi =>
      `<div class="overview-kpi-cell dnd-grid-card" draggable="true" data-dnd-id="${kpi.id}"
           ondragstart="OverviewPanel.dndStart('${kpi.id}',event)"
           ondragend="OverviewPanel.dndEnd()"
           ondragover="event.stopPropagation();OverviewPanel.dndOver('${kpi.id}',event)"
           ondrop="event.stopPropagation();OverviewPanel.dndDrop('${kpi.id}',event)">
        <div class="dnd-grip-overlay" title="Drag to reorder">⠿</div>
        ${_renderKpiCard(kpi, st, mode)}
      </div>`
    ).join('');

    // Collapsed hero picker dropdown — rendered inline inside collapsed card
    const pickerOpen = st.kpiPickerOpen || false;
    const heroKpiIdsSafe = Array.isArray(heroKpiIds) ? heroKpiIds : [];
    const collapsedPickerDropdown = pickerOpen ? `
      <div id="hero-collapsed-picker" style="position:absolute;left:50%;transform:translateX(-50%);top:calc(100% + 8px);width:300px;z-index:200;
                  background:var(--bg-modal);border:1px solid var(--border-card);
                  border-radius:10px;box-shadow:var(--shadow-modal);overflow:hidden"
           onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
        <div style="padding:10px 12px;border-bottom:1px solid var(--border-subtle);display:flex;align-items:center;justify-content:space-between">
          <div style="font-size:12px;font-weight:600;color:var(--text-primary)">Select KPIs for Hero Card</div>
          <button onclick="event.stopPropagation();OverviewPanel.closeCollapsedPicker()"
                  style="color:var(--text-muted);font-size:16px;background:none;border:none;cursor:pointer;padding:0 2px">✕</button>
        </div>
        <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
          <input type="text" class="input-field" style="padding:6px 10px;font-size:12px"
                 placeholder="Search KPIs…"
                 onclick="event.stopPropagation()" onmousedown="event.stopPropagation()"
                 oninput="OverviewPanel.filterHeroKpiSearch(this.value)">
        </div>
        <div id="hero-kpi-list" style="max-height:240px;overflow-y:auto">
          ${allKpis.length === 0
            ? '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs yet — add some in Data Entry first</div>'
            : allKpis.map(k => `
                <div class="hero-kpi-item" onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
                  <div style="flex:1;min-width:0">
                    <div class="hki-name">${_esc(k.metric)}</div>
                    <div class="hki-section">${_esc(k.section)}</div>
                  </div>
                  <input type="checkbox" ${heroKpiIdsSafe.includes(k.id)?'checked':''}
                         style="accent-color:var(--brand-accent);width:15px;height:15px;flex-shrink:0;cursor:pointer"
                         onclick="event.stopPropagation()" onmousedown="event.stopPropagation()"
                         onchange="event.stopPropagation();OverviewPanel.toggleHeroKpi('${k.id}',this.checked)">
                </div>`).join('')}
        </div>
        <div style="padding:8px 12px;border-top:1px solid var(--border-subtle);display:flex;gap:8px"
             onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
          <button onclick="event.stopPropagation();OverviewPanel.setHeroAllKpis(true)"  class="btn btn-ghost"   style="flex:1;font-size:11px;padding:5px">All</button>
          <button onclick="event.stopPropagation();OverviewPanel.setHeroAllKpis(false)" class="btn btn-ghost"   style="flex:1;font-size:11px;padding:5px">None</button>
          <button onclick="event.stopPropagation();OverviewPanel.closeCollapsedPicker()"  class="btn btn-primary" style="flex:1;font-size:11px;padding:5px">Done</button>
        </div>
      </div>` : '';

    const collapsedHeroCard = `
      <div class="overview-hero-slot overview-hero-collapsed${pickerOpen ? ' picker-open' : ''}"
           style="position:relative;">
        <button onclick="event.stopPropagation();OverviewPanel.toggleKpiPicker()"
                style="background:none;border:none;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:8px;padding:16px;
                       width:100%;height:100%;justify-content:center">
          <div style="width:40px;height:40px;border-radius:50%;
                      border:2px solid ${pickerOpen ? 'var(--brand-accent)' : 'var(--border-card)'};
                      display:flex;align-items:center;justify-content:center;font-size:22px;
                      color:${pickerOpen ? 'var(--brand-accent)' : 'var(--text-muted)'};
                      transition:all 0.15s">+</div>
          <div style="font-size:11px;color:${pickerOpen ? 'var(--brand-accent)' : 'var(--text-muted)'};font-weight:500;transition:color 0.15s">Hero chart</div>
        </button>
        ${collapsedPickerDropdown}
      </div>`;

    const expandedHeroCard = `
      <div class="overview-hero-slot overview-hero-expanded">
        ${_renderHeroCard(st, heroKpis, allKpis, settings)}
      </div>`;

    const emptyState = allKpis.length === 0 ? `
      <div style="grid-column:span 4">
        <div class="card" style="text-align:center;padding:40px">
          <div style="font-size:32px;margin-bottom:12px">◈</div>
          <div style="font-size:16px;font-weight:600;margin-bottom:6px">No Key Metrics yet</div>
          <div style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">In Data Entry, tick the <strong>Ovw</strong> checkbox to pin KPIs here.</div>
          <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
        </div>
      </div>` : '';

    const heroWithSide = `
      <div class="overview-kpi-grid" id="dnd-zone-grid" style="margin-bottom:24px"
           ondragover="OverviewPanel.dndOver(null,event)"
           ondrop="OverviewPanel.dndDrop(null,event)">
        ${isCollapsed ? collapsedHeroCard : expandedHeroCard}
        ${allKpiCards}
        ${emptyState}
      </div>`;

    const sectionGrid = sections.length > 0 ? `
      <div style="margin-top:8px">
        <h3 style="font-family:var(--font-display);font-size:15px;font-weight:600;margin-bottom:12px">All Sections</h3>
        <div class="grid-3">
          ${sections.map(section => {
            const skpis  = DataStore.getKpisBySection(section);
            const counts = {green:0,amber:0,red:0,neutral:0};
            skpis.forEach(k=>{
              const ps = DataStore.getPeriodStats(k, mode);
              counts[ps.rag || k.rag || 'neutral']++;
            });
            const pid = 'section:'+encodeURIComponent(section);
            return `
              <div class="card" style="cursor:pointer;transition:all 0.15s" onclick="App.navigate('${pid}')"
                   onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform=''">
                <div style="font-family:var(--font-display);font-size:14px;font-weight:600;margin-bottom:8px">${_esc(section)}</div>
                <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px">
                  ${Object.entries(counts).filter(([,v])=>v>0).map(([r,c])=>`<span class="rag-badge ${r}" style="font-size:10px;padding:2px 7px">${c} ${r}</span>`).join('')}
                </div>
                <div class="label-xs">${skpis.length} KPIs →</div>
              </div>`;
          }).join('')}
        </div>
      </div>` : '';

    return `
      <div style="margin-bottom:24px">
        ${header}${ragStrip}${heroWithSide}${sectionGrid}
      </div>

      <style>
        /* ══════════════════════════════════════════════════════
           OVERVIEW GRID — fixed row height, strict card sizing
           Hero = 3 cols × 2 rows (416px incl gap). Every KPI cell = 1 col × 1 row = 200px.
           Grid auto-places KPI cards around the hero naturally.
           ══════════════════════════════════════════════════════ */

        .overview-kpi-grid {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          grid-auto-rows: 200px;
          gap: 16px;
          align-items: stretch;
          overflow: visible;
        }

        /* Hero slot base */
        .overview-hero-slot {
          border-radius: 16px;
          box-shadow: var(--shadow-card);
        }
        .overview-hero-expanded {
          grid-column: 1 / span 3;
          grid-row: 1 / span 2;
          background: var(--bg-card);
          border: 1px solid var(--border-card);
          padding: 14px 20px 16px;
          display: flex;
          flex-direction: column;
          overflow: hidden;
        }
        .overview-hero-collapsed {
          grid-column: 1 / span 1;
          grid-row: 1 / span 1;
          background: var(--bg-card);
          border: 2px dashed var(--border-card);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 8px;
          overflow: visible;
          position: relative;
        }
        .overview-hero-collapsed.picker-open {
          border-color: var(--brand-accent);
        }

        /* KPI card cell — grid places it, overflow visible so dropdown escapes */
        .overview-kpi-cell {
          min-height: 0;
          min-width: 0;
          box-sizing: border-box;
          overflow: visible;
          position: relative;
          z-index: 1;
        }
        .overview-kpi-cell:has(.card-viz-dropdown) { z-index: 100; }

        /* ── KPI viz card shell — fills its cell exactly ── */
        .kpi-viz-card {
          background: var(--bg-card);
          border: 1px solid var(--border-card);
          border-radius: var(--card-radius);
          padding: 12px 14px 10px;
          box-shadow: var(--shadow-card);
          width: 100%;
          height: 100%;
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
          position: relative;
          overflow: visible;   /* must be visible so dropdown isn't clipped */
          cursor: pointer;
        }
        /* Clip only the interior content, not the whole card */
        .kpi-viz-card > *:not(.card-viz-dropdown) { position:relative; z-index:1 }
        .kpi-viz-card:hover { border-color: var(--brand-accent); }

        /* ── Viz content area — takes remaining height after header ── */
        .card-viz-content {
          flex: 1;
          min-height: 0;
          overflow: hidden;         /* critical: clips charts to card bounds */
          display: flex;
          flex-direction: column;
          justify-content: center;
        }

        /* Tighten value font in cards so it fits */
        .kpi-viz-card .value-xl { font-size: 22px; }

        /* ── Card viz ••• menu ── */
        .card-viz-menu-btn { opacity:0;transition:opacity 0.15s;background:var(--bg-input);border:1px solid var(--border-card);border-radius:5px;padding:2px 7px;cursor:pointer;color:var(--text-muted);font-size:12px;letter-spacing:1.5px;line-height:1.4 }
        .kpi-viz-card:hover .card-viz-menu-btn { opacity:1 }
        .card-viz-menu-btn.open { opacity:1;color:var(--brand-accent);border-color:var(--brand-accent) }
        /* Dropdown: absolute within card header's relative container, z-index above grid */
        .card-viz-dropdown { position:absolute;right:0;top:calc(100% + 2px);width:168px;background:var(--bg-modal);border:1px solid var(--border-card);border-radius:8px;box-shadow:var(--shadow-modal);z-index:9999;overflow:hidden }
        .card-viz-option { display:flex;align-items:center;gap:9px;padding:8px 12px;cursor:pointer;font-size:12px;font-weight:500;color:var(--text-secondary);transition:background 0.1s }
        .card-viz-option:hover { background:var(--bg-card-hover);color:var(--text-primary) }
        .card-viz-option.active { color:var(--brand-accent);background:rgba(0,194,168,0.08) }

        /* ── Odometer layout (hero chart area) ── */
        .odo-row  { display:grid;grid-auto-columns:1fr;grid-auto-flow:column;gap:16px;width:100%;height:100%;align-items:stretch }
        .odo-col  { display:flex;flex-direction:column;align-items:center;position:relative;min-width:0;min-height:0 }
        .odo-block { display:flex;flex-direction:column;align-items:center;width:100%;height:100%;min-height:0 }
        .odo-main  { width:100% }
        .odo-small { width:100% }
        .odo-svg   { width:100%;display:block }
        /* Hub ⊕ button — lives inside the SVG, hover styles via CSS */
        .odo-hub-ring { transition:stroke 0.15s,stroke-width 0.15s }
        .odo-hub-btn:hover .odo-hub-ring { stroke:var(--brand-accent) !important;stroke-width:2 !important }
        .odo-hub-plus { opacity:0;transition:opacity 0.15s }
        .odo-hub-btn:hover .odo-hub-plus { opacity:1 }
        .odo-reorder-btns { display:flex;align-items:center;gap:4px;margin-bottom:6px;opacity:0;transition:opacity 0.15s;flex-shrink:0 }
        .odo-col:hover .odo-reorder-btns { opacity:1 }
        .odo-move-btn { background:var(--bg-input);border:1px solid var(--border-card);border-radius:4px;color:var(--text-muted);font-size:12px;padding:1px 7px;cursor:pointer;transition:all 0.1s }
        .odo-move-btn:hover:not([disabled]) { color:var(--brand-accent);border-color:var(--brand-accent) }
        .odo-move-btn[disabled] { opacity:0.2;cursor:default }
        .odo-pos-label { font-size:10px;color:var(--text-muted);min-width:28px;text-align:center }
        .hero-menu-btn { opacity:0.6;transition:opacity 0.15s }
        .hero-menu-btn:hover { opacity:1 }
        .hero-kpi-item { display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--border-subtle);transition:background 0.1s }
        .hero-kpi-item:hover { background:var(--bg-card-hover) }
        .hki-name    { font-size:12px;font-weight:500;color:var(--text-primary);line-height:1.35 }
        .hki-section { font-size:10px;color:var(--text-muted);margin-top:2px }

        /* ── Responsive: collapse to 2-col on tablet ── */
        @media (max-width: 900px) {
          .overview-kpi-grid { grid-template-columns: repeat(2, 1fr); }
          .overview-hero-expanded { grid-column: 1 / span 2; grid-row: 1 / span 2; }
        }
        /* ── Mobile: single column, vertical hero metrics, auto heights ── */
        @media (max-width: 560px) {
          .overview-kpi-grid {
            grid-template-columns: 1fr;
            grid-auto-rows: auto;
            gap: 12px;
          }
          .overview-hero-expanded {
            grid-column: 1;
            grid-row: auto;
            min-height: 0;
            padding: 12px;
          }
          .overview-hero-collapsed { min-height: 100px; }
          /* Stack each hero metric below the previous one */
          .odo-row {
            display: flex;
            flex-direction: column;
            height: auto;
            gap: 20px;
          }
          .odo-col { min-height: 220px; }
          /* Hide reorder controls — not useful on mobile */
          .odo-reorder-btns { display: none !important; }
          /* KPI viz cards: natural height instead of stretching to fixed row */
          .kpi-viz-card { height: auto; min-height: 150px; }
          /* Tighten card viz option row */
          .card-viz-option { padding: 10px 12px; }
        }

        /* ── Drag-and-drop ── */
        .dnd-grid-card[draggable="true"] { cursor: grab; }
        .dnd-grid-card[draggable="true"]:active { cursor: grabbing; }

        .dnd-dragging { opacity: 0.25 !important; }

        /* Swap target indicator: teal border all around */
        .dnd-drop-before .kpi-viz-card {
          border-color: var(--brand-accent) !important;
          box-shadow: 0 0 0 2px rgba(0,194,168,0.35) !important;
        }

        /* Zone-level drop highlight */
        #dnd-zone-grid.dnd-zone-active {
          outline: 2px dashed var(--brand-accent);
          outline-offset: 6px;
          border-radius: 12px;
        }

        /* Grip handle overlay — shows on hover */
        .dnd-grip-overlay {
          position: absolute;
          top: 7px; left: 7px;
          z-index: 20;
          font-size: 13px;
          color: var(--text-muted);
          opacity: 0;
          transition: opacity 0.15s;
          pointer-events: none;
          user-select: none;
          line-height: 1;
        }
        .dnd-grid-card:hover .dnd-grip-overlay { opacity: 0.55; }
      </style>`;
  }

  // ── Per-card render (with viz type switcher) ──────────────────────────────
  function _renderKpiCard(kpi, st, mode) {
    const cardVizTypes = st.cardVizTypes || {};
    const vizType      = cardVizTypes[kpi.id] || 'card';
    const menuOpen     = st.cardMenuOpen === kpi.id;
    const rag          = kpi.rag || 'neutral';

    // Always use period-aware RAG so the line colour matches the data shown
    const ps      = DataStore.getPeriodStats(kpi, mode);
    const cardRag = ps.rag || kpi.rag || 'neutral';
    const ragColor = { green:'var(--rag-green)', amber:'var(--rag-amber)', red:'var(--rag-red)', neutral:'var(--rag-neutral)' }[cardRag];

    // Inline RAG line — sits inside the card padding, fully contained, never overlaps border
    const ragLine = `<div style="height:2px;border-radius:2px;background:${ragColor};margin-bottom:10px;opacity:0.9"></div>`;

    const vizOptions = [
      { v:'card',     l:'KPI Card',    i:'⊡' },
      { v:'odometer', l:'Odometer',    i:'◎' },
      { v:'bar',      l:'Bar Chart',   i:'▊' },
      { v:'line',     l:'Line Chart',  i:'∿' },
    ];


    const _metricLen = kpi.metric.length;
    const _metricSize = _metricLen > 60 ? '10px' : _metricLen > 40 ? '11px' : _metricLen > 25 ? '12px' : '13px';
    const cardHeader = `
      <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:4px;position:relative">
        <div style="flex:1;min-width:0">
          <div style="font-size:9px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:var(--text-muted);line-height:1.2">${kpi.section}</div>
          <div style="font-size:${_metricSize};font-weight:500;color:var(--text-primary);margin-top:2px;line-height:1.3;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical">${kpi.metric}</div>
        </div>
        <div style="display:flex;align-items:center;gap:4px;flex-shrink:0;margin-left:6px">
          ${kpi.who?`<div class="label-xs" style="color:var(--brand-accent)">${kpi.who}</div>`:''}
          <button class="card-viz-menu-btn ${menuOpen?'open':''}"
                  onclick="event.stopPropagation();OverviewPanel.toggleCardMenu('${kpi.id}',this)"
                  title="Change visualization">•••</button>
        </div>
      </div>`;

    if (vizType === 'card') {
      // Standard KPI card content
      const val        = DataStore.formatValue(ps.actual, kpi);
      const targetDisp = ps.target !== null ? DataStore.formatValue(ps.target, kpi) : '—';
      const pct        = ps.progressPct;

      return `
        <div class="kpi-viz-card rag-${cardRag}" onclick="App.openKpiDetail('${kpi.id}')">
          ${ragLine}
          ${cardHeader}
          <div class="card-viz-content">
            <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:4px">
              <div class="value-xl">${val}</div>
              <div class="label-xs" style="color:var(--text-muted)">${ps.label}</div>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:6px;flex-wrap:wrap">
              <span class="label-xs">Target: ${targetDisp}</span>
              <span class="rag-badge ${cardRag}" style="font-size:10px;padding:2px 7px">${cardRag==='green'?'● On Track':cardRag==='amber'?'▲ At Risk':cardRag==='red'?'✕ Off Track':'○ No Data'}</span>
            </div>
            <div class="progress-bar">
              <div class="progress-bar-fill" style="width:${pct}%;background:${ragColor}"></div>
            </div>
            ${pct>0?`<div class="label-xs" style="margin-top:4px;text-align:right">${pct}% of target</div>`:''}
            ${kpi.comment?`<div style="margin-top:6px;font-size:11px;color:var(--rag-amber);font-style:italic">${kpi.comment}</div>`:''}
          </div>
        </div>`;
    }

    // Non-card viz types: render a card shell with a viz content area
    return `
      <div class="kpi-viz-card rag-${rag}" onclick="App.openKpiDetail('${kpi.id}')">
        ${ragLine}
        ${cardHeader}
        <div class="card-viz-content" id="card-viz-${kpi.id}" onclick="event.stopPropagation()">
          <div style="text-align:center;color:var(--text-muted);font-size:11px;padding:8px">Loading…</div>
        </div>
      </div>`;
  }

  // ── Header ────────────────────────────────────────────────────────────────
  function _renderHeader(settings, overviewKpis, allKpis) {
    const overviewIds = settings.overviewKpiIds;
    const manageList  = allKpis.length === 0 ? '' : allKpis.map((k, i) => {
      const isVisible = k.isKey && (!overviewIds || overviewIds.includes(k.id));
      return `
        <div class="mkpi-row" style="display:flex;align-items:center;gap:10px;padding:8px 12px;${i>0?'border-top:1px solid var(--border-subtle)':''}transition:background 0.1s"
             onmouseover="this.style.background='var(--bg-card-hover)'" onmouseout="this.style.background=''">
          <div style="flex:1;min-width:0">
            <div style="font-size:12px;font-weight:500;color:var(--text-primary)">${_esc(k.metric)}</div>
            <div style="font-size:10px;color:var(--text-muted)">${_esc(k.section)}</div>
          </div>
          <input type="checkbox" ${isVisible?'checked':''}
                 style="accent-color:var(--brand-accent);width:15px;height:15px;cursor:pointer"
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
              <span style="background:rgba(0,194,168,0.15);color:var(--brand-accent);font-size:10px;padding:1px 6px;border-radius:10px;font-weight:700">${overviewKpis.length}/${allKpis.length}</span>
            </button>
            <div id="manage-kpi-dropdown" style="display:none;position:absolute;right:0;top:calc(100% + 6px);
                 width:320px;background:var(--bg-modal);border:1px solid var(--border-card);
                 border-radius:10px;box-shadow:var(--shadow-modal);z-index:150;overflow:hidden">
              <div style="padding:10px 12px;border-bottom:1px solid var(--border-subtle);display:flex;align-items:center;justify-content:space-between">
                <div style="font-size:12px;font-weight:600">Manage Overview KPIs</div>
                <button onclick="App.toggleManageKpiDropdown()" style="color:var(--text-muted);font-size:16px;background:none;border:none;cursor:pointer">✕</button>
              </div>
              <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
                <input type="text" id="manage-kpi-search" class="input-field" style="padding:6px 10px;font-size:12px"
                       placeholder="Search KPIs…" oninput="App.filterManageKpiList(this.value)">
              </div>
              <div id="manage-kpi-list" style="max-height:260px;overflow-y:auto">
                ${manageList||'<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs yet</div>'}
              </div>
              <div style="padding:8px 12px;border-top:1px solid var(--border-subtle);display:flex;gap:8px">
                <button onclick="App.setAllOverviewKpis(true)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">All</button>
                <button onclick="App.setAllOverviewKpis(false)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">None</button>
              </div>
            </div>
          </div>` : ''}
          <select class="input-field" style="width:auto;padding:6px 28px 6px 10px;font-size:12px"
                  onchange="DataStore.updateSettings({reportingPeriod:this.value})">
            <option value="monthly"   ${settings.reportingPeriod==='monthly'?'selected':''}>Monthly</option>
            <option value="quarterly" ${settings.reportingPeriod==='quarterly'?'selected':''}>Quarterly</option>
            <option value="ytd"       ${settings.reportingPeriod==='ytd'?'selected':''}>YTD</option>
            <option value="yearly"    ${settings.reportingPeriod==='yearly'?'selected':''}>Full Year</option>
          </select>
        </div>
      </div>`;
  }

  // ── RAG Strip ─────────────────────────────────────────────────────────────
  function _renderRagStrip(ragCounts, total) {
    return `
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px">
        ${Object.entries(ragCounts).map(([rag,cnt]) => cnt > 0 ? `
          <div style="background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;gap:8px">
            <div class="rag-badge ${rag}">${rag.charAt(0).toUpperCase()+rag.slice(1)}</div>
            <span style="font-family:var(--font-display);font-size:20px;font-weight:700">${cnt}</span>
          </div>` : '').join('')}
        <div style="flex:1;background:var(--bg-card);border:1px solid var(--border-card);border-radius:8px;padding:10px 16px;display:flex;align-items:center;justify-content:flex-end;gap:6px">
          <span class="label-xs">Total KPIs:</span>
          <span style="font-size:14px;font-weight:700;color:var(--brand-accent)">${total}</span>
        </div>
      </div>`;
  }

  // ── Hero Card ─────────────────────────────────────────────────────────────
  function _renderHeroCard(st, heroKpis, allKpis, settings) {
    const chartType   = st.chartType || 'odometer';
    const heroPosition= st.heroPosition || 'center';
    const menuOpen    = st.menuOpen || false;
    const pickerOpen  = st.kpiPickerOpen || false;
    const heroKpiIds  = st.heroKpiIds || [];

    const chartOpts = [
      { v:'odometer', l:'Odometer Gauges', i:'◎' },
      { v:'bar',      l:'Bar Chart',       i:'▊' },
      { v:'line',     l:'Line Chart',      i:'∿' },
    ];
    const curOpt = chartOpts.find(o => o.v === chartType) || chartOpts[0];

    const menuDropdown = menuOpen ? `
      <div id="hero-menu-dropdown" style="position:absolute;right:0;top:calc(100% + 6px);width:230px;
                  background:var(--bg-modal);border:1px solid var(--border-card);
                  border-radius:10px;box-shadow:var(--shadow-modal);z-index:160;overflow:hidden">
        <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
          <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-muted)">Chart Type</div>
        </div>
        ${chartOpts.map(o => `
          <div onclick="OverviewPanel.setChartType('${o.v}')"
               style="display:flex;align-items:center;gap:10px;padding:9px 14px;cursor:pointer;
                      background:${o.v===chartType?'rgba(0,194,168,0.08)':'transparent'};
                      color:${o.v===chartType?'var(--brand-accent)':'var(--text-secondary)'};transition:background 0.1s"
               onmouseover="this.style.background='var(--bg-card-hover)'"
               onmouseout="this.style.background='${o.v===chartType?'rgba(0,194,168,0.08)':'transparent'}'">
            <span style="font-size:15px;width:18px;text-align:center">${o.i}</span>
            <span style="font-size:13px;font-weight:500">${o.l}</span>
            ${o.v===chartType?'<span style="margin-left:auto;font-size:11px;color:var(--brand-accent)">✓</span>':''}
          </div>`).join('')}
      </div>` : '';

    const pickerDropdown = pickerOpen ? `
      <div id="hero-picker-dropdown" style="position:absolute;left:0;top:calc(100% + 6px);width:320px;
                  background:var(--bg-modal);border:1px solid var(--border-card);
                  border-radius:10px;box-shadow:var(--shadow-modal);z-index:160;overflow:hidden">
        <div style="padding:10px 12px;border-bottom:1px solid var(--border-subtle);display:flex;align-items:center;justify-content:space-between">
          <div style="font-size:12px;font-weight:600">Select KPIs for Hero Card</div>
          <button onclick="OverviewPanel.toggleKpiPicker()" style="color:var(--text-muted);font-size:16px;background:none;border:none;cursor:pointer">✕</button>
        </div>
        <div style="padding:8px 12px;border-bottom:1px solid var(--border-subtle)">
          <input type="text" class="input-field" style="padding:6px 10px;font-size:12px"
                 placeholder="Search KPIs…" oninput="OverviewPanel.filterHeroKpiSearch(this.value)">
        </div>
        <div id="hero-kpi-list" style="max-height:240px;overflow-y:auto">
          ${allKpis.length === 0
            ? '<div style="padding:16px;text-align:center;color:var(--text-muted);font-size:12px">No KPIs yet</div>'
            : allKpis.map(k => `
                <div class="hero-kpi-item">
                  <div style="flex:1;min-width:0">
                    <div class="hki-name">${_esc(k.metric)}</div>
                    <div class="hki-section">${_esc(k.section)}</div>
                  </div>
                  <input type="checkbox" ${heroKpiIds.includes(k.id)?'checked':''}
                         style="accent-color:var(--brand-accent);width:15px;height:15px;flex-shrink:0;cursor:pointer"
                         onchange="OverviewPanel.toggleHeroKpi('${k.id}',this.checked)">
                </div>`).join('')}
        </div>
        <div style="padding:8px 12px;border-top:1px solid var(--border-subtle);display:flex;gap:8px">
          <button onclick="OverviewPanel.setHeroAllKpis(true)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">All</button>
          <button onclick="OverviewPanel.setHeroAllKpis(false)" class="btn btn-ghost" style="flex:1;font-size:11px;padding:5px">None</button>
          <button onclick="OverviewPanel.closeCollapsedPicker()" class="btn btn-primary" style="flex:1;font-size:11px;padding:5px">Done</button>
        </div>
      </div>` : '';

    const heroCount = heroKpiIds.length > 0 ? heroKpiIds.length : (DataStore.getKeyKpis().length || allKpis.length);
    const isEmpty   = allKpis.length === 0;

    const cardContent = isEmpty
      ? `<div style="text-align:center;padding:48px 24px;color:var(--text-muted)">
           <div style="font-size:36px;margin-bottom:12px">◈</div>
           <div style="font-size:15px;font-weight:600;margin-bottom:6px;color:var(--text-secondary)">No KPIs configured</div>
           <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
         </div>`
      : `<div id="hero-chart-area" style="flex:1;min-height:0;overflow:hidden;display:flex;align-items:stretch">
           <div style="text-align:center;padding:32px;color:var(--text-muted);font-size:12px;width:100%;display:flex;align-items:center;justify-content:center">Loading…</div>
         </div>`;

    return `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;gap:12px;flex-wrap:wrap">
        <div style="position:relative" id="hero-picker-wrap">
          <button onclick="OverviewPanel.toggleKpiPicker()"
                  style="display:flex;align-items:center;gap:7px;padding:6px 13px;border-radius:8px;
                         font-size:12px;font-weight:600;border:1px solid var(--border-card);
                         background:var(--bg-input);color:var(--text-secondary);cursor:pointer;transition:all 0.15s"
                  onmouseover="this.style.borderColor='var(--brand-accent)';this.style.color='var(--text-primary)'"
                  onmouseout="this.style.borderColor='var(--border-card)';this.style.color='var(--text-secondary)'">
            <span style="font-size:15px">◉</span>
            ${heroCount} KPI${heroCount!==1?'s':''} selected
            <span style="font-size:9px;opacity:0.5">▾</span>
          </button>
          ${pickerDropdown}
        </div>
        <div style="display:flex;align-items:center;gap:7px;font-size:12px;font-weight:600;
                    color:var(--brand-accent);background:rgba(0,194,168,0.08);
                    border:1px solid rgba(0,194,168,0.2);border-radius:8px;padding:5px 13px">
          <span>${curOpt.i}</span><span>${curOpt.l}</span>
        </div>
        <div style="position:relative" id="hero-menu-wrap">
          <button class="hero-menu-btn" onclick="OverviewPanel.toggleMenu()"
                  style="background:var(--bg-input);border:1px solid var(--border-card);
                         border-radius:7px;padding:5px 10px;cursor:pointer;
                         color:var(--text-secondary);font-size:16px;letter-spacing:2px">•••</button>
          ${menuDropdown}
        </div>
      </div>
      <div style="height:2px;background:linear-gradient(90deg,var(--brand-accent),var(--brand-accent-2),transparent);border-radius:1px;margin-bottom:12px;opacity:0.6"></div>
      ${cardContent}`;
  }

  // ── Formula modal helpers ─────────────────────────────────────────────────
  function _fmlFilterKpis(q) {
    const list  = document.getElementById('fml-kpi-list');
    const empty = document.getElementById('fml-kpi-empty');
    if (!list) return;
    const query = q.toLowerCase().trim();
    const labels = list.querySelectorAll('label.fml-kpi-row');
    let visible = 0;
    labels.forEach(lbl => {
      const name = lbl.querySelector('div > div:first-child')?.textContent?.toLowerCase() || '';
      const sect = lbl.querySelector('div > div:last-child')?.textContent?.toLowerCase() || '';
      const show = !query || name.includes(query) || sect.includes(query);
      // Use 'flex' not '' to preserve flex layout ('' would erase the inline display:flex)
      lbl.style.display = show ? 'flex' : 'none';
      if (show) {
        lbl.style.borderTop = visible === 0 ? 'none' : '1px solid var(--border-subtle)';
        visible++;
      }
    });
    if (empty) empty.style.display = visible===0 && labels.length>0 ? '' : 'none';
    if (list)  list.style.display  = visible===0 && labels.length>0 ? 'none' : '';
  }

  function _fmlFilterRef(q) {
    const tbody = document.querySelector('#fml-ref-table tbody');
    if (!tbody) return;
    const query = q.toLowerCase().trim();
    tbody.querySelectorAll('tr').forEach(row => {
      row.style.display = !query || row.textContent.toLowerCase().includes(query) ? '' : 'none';
    });
  }

  function _fmlSelectOp(op) {
    const hidden = document.getElementById('fml-op');
    if (!hidden) return;
    hidden.value = op;
    ['sum','subtract','multiply','divide','avg','min','max','custom'].forEach(o => {
      const btn = document.getElementById('fml-op-btn-'+o);
      const chk = document.getElementById('fml-op-chk-'+o);
      if (!btn) return;
      const active = o === op;
      btn.style.borderColor = active ? 'var(--brand-accent)' : 'var(--border-card)';
      btn.style.background  = active ? 'rgba(0,194,168,0.08)' : 'var(--bg-input)';
      const lbl = btn.querySelector('div:first-child > div:first-child');
      if (lbl) lbl.style.color = active ? 'var(--brand-accent)' : 'var(--text-primary)';
      if (chk) {
        chk.innerHTML = active ? '<div style="width:8px;height:8px;border-radius:50%;background:var(--brand-accent)"></div>' : '';
        chk.style.borderColor = active ? 'var(--brand-accent)' : 'var(--border-card)';
      }
    });
    const srcSec = document.getElementById('fml-source-section');
    const cstSec = document.getElementById('fml-custom-section');
    if (srcSec) srcSec.style.display = op==='custom' ? 'none' : '';
    if (cstSec) cstSec.style.display = op==='custom' ? '' : 'none';
  }

  // ─────────────────────────────────────────────────────────────────────────
  return {
    render, mountCharts,
    setChartType, toggleMenu, closeMenu, setHeroPosition,
    toggleKpiPicker, closeCollapsedPicker, toggleHeroKpi, setHeroAllKpis, filterHeroKpiSearch,
    moveHeroKpi, openStackPicker, setStack, _filterStackList,
    setCardVizType, toggleCardMenu, closeCardMenu,
    dndStart, dndEnd, dndOver, dndDrop,
    _fmlSelectOp, _fmlFilterKpis, _fmlFilterRef,
  };
})();
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
    const mode = DataStore.getSettings().reportingPeriod || 'monthly';
    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:10px">
          <div>
            <div class="label-xs" style="margin-bottom:4px;cursor:pointer;color:var(--brand-accent)" onclick="App.navigate('overview')">← Overview</div>
            <h2 class="section-title" style="margin:0">${sectionName}</h2>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn btn-ghost" style="font-size:12px" onclick="App.promptRenameSection('${encodeURIComponent(sectionName)}')">✎ Rename</button>
            <button class="btn btn-ghost" style="font-size:12px;color:var(--rag-red)" onclick="App.confirmRemoveSection('${encodeURIComponent(sectionName)}')">🗑 Remove</button>
          </div>
        </div>
        ${Components.ragSummaryBar(kpis, mode)}
        ${kpis.length>0?`
          <div class="grid-auto">${kpis.map(kpi=>Components.kpiCard(kpi, DataStore.getPeriodStats(kpi, DataStore.getSettings().reportingPeriod||'monthly'), false)).join('')}</div>`:`
          <div class="card" style="text-align:center;padding:40px">
            <div style="font-size:28px;margin-bottom:10px">📋</div>
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
          <div class="de-add-btns" style="display:flex;gap:8px">
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
    const allKpis    = DataStore.getKpis().filter(k => !k.isFormula);
    const sections   = [...new Set(allKpis.map(k => k.section))];
    const thresholds = DataStore.getThresholds();
    const fyMonths   = DataStore.getFyMonths();

    if (allKpis.length === 0) {
      return `
        <div class="card" style="text-align:center;padding:56px 40px;border:2px dashed var(--border-card)">
          <div style="font-size:40px;margin-bottom:14px">📋</div>
          <div style="font-family:var(--font-display);font-size:18px;font-weight:700;margin-bottom:8px">No KPIs yet</div>
          <div style="font-size:13px;color:var(--text-secondary);max-width:360px;margin:0 auto 20px">
            Start by creating a section, then add your KPIs. You can also bulk-import from XLSX or CSV.
          </div>
          <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
            <button class="btn btn-primary" onclick="App.openAddKpi()">+ Add First KPI</button>
            <button class="btn btn-ghost" onclick="App.openAddSectionModal()">+ New Section</button>
          </div>
          <div style="margin-top:28px">${Components.xlsxImportPanel()}</div>
        </div>
        <button class="mobile-fab" onclick="App.openAddKpi()" title="Add new KPI">+</button>`;
    }

    return `
      <!-- Bulk Import -->
      <div style="margin-bottom:20px">
        <details style="background:var(--bg-card);border:1px solid var(--border-card);border-radius:10px;overflow:hidden">
          <summary style="padding:12px 16px;cursor:pointer;font-size:13px;font-weight:500;color:var(--text-secondary);list-style:none;display:flex;align-items:center;gap:8px">
            <span style="font-size:16px">📊</span> Bulk Import via XLSX / CSV
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
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none">⌕</span>
            </div>
            <div style="position:relative" id="sec-filter-adv-wrap">
              <button onclick="App.toggleSectionFilterDropdown('adv')"
                      style="display:flex;align-items:center;gap:5px;padding:7px 12px;border-radius:8px;font-size:12px;font-weight:500;
                             border:1px solid ${App._kpiSectionFilter?'var(--brand-accent)':'var(--border-card)'};
                             background:${App._kpiSectionFilter?'rgba(0,194,168,0.1)':'var(--bg-card)'};
                             color:${App._kpiSectionFilter?'var(--brand-accent)':'var(--text-secondary)'};cursor:pointer;white-space:nowrap">
                ⊞ ${App._kpiSectionFilter||'All Sections'}
              </button>
            </div>` : `
            <div style="position:relative;flex:1;max-width:280px">
              <input type="text" id="kpi-search-simple" class="input-field"
                     style="padding:7px 12px 7px 32px;font-size:12px"
                     placeholder="Search KPIs…"
                     value="${App._simpleSearchQuery||''}"
                     oninput="App.setSimpleKpiSearch(this.value)">
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none">⌕</span>
            </div>
            <div style="position:relative" id="sec-filter-simple-wrap">
              <button onclick="App.toggleSectionFilterDropdown('simple')"
                      style="display:flex;align-items:center;gap:5px;padding:7px 12px;border-radius:8px;font-size:12px;font-weight:500;
                             border:1px solid ${App._simpleSectionFilter?'var(--brand-accent)':'var(--border-card)'};
                             background:${App._simpleSectionFilter?'rgba(0,194,168,0.1)':'var(--bg-card)'};
                             color:${App._simpleSectionFilter?'var(--brand-accent)':'var(--text-secondary)'};cursor:pointer;white-space:nowrap">
                ⊞ ${App._simpleSectionFilter||'All Sections'}
              </button>
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
        // Re-apply simple search + section filter after re-render
        (function(){
          setTimeout(function(){
            App.setSimpleKpiSearch(App._simpleSearchQuery||'');
            App.applySimpleSectionFilter();
          }, 0);
        })();
      </script>
      <button class="mobile-fab" onclick="App.openAddKpi()" title="Add new KPI">+</button>`;
  }

  // ══════════════════════════════════════════════════════════════════════════
  //  SIMPLE MODE — compact table with compound target cells
  // ══════════════════════════════════════════════════════════════════════════
  function _buildSimpleSection(section, sKpis, thresholds, expandedKpiId, fyMonths) {
    const enc = encodeURIComponent(section);

    // Mobile card list — one row per KPI, tap to edit
    const RAG_COL = { green:'var(--rag-green)', amber:'var(--rag-amber)', red:'var(--rag-red)', neutral:'var(--rag-neutral)' };
    const mobileCards = sKpis.length === 0
      ? `<div style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;font-style:italic">No KPIs yet — tap + to add one.</div>`
      : sKpis.map(kpi => {
          const rag      = kpi.rag || 'neutral';
          const ragColor = RAG_COL[rag];
          const ytdFmt   = kpi.ytd != null ? DataStore.formatValue(kpi.ytd, kpi) : null;
          return `
            <div class="mobile-kpi-card" onclick="App.openEditKpi('${kpi.id}')">
              <div style="width:3px;background:${ragColor};align-self:stretch;flex-shrink:0;border-radius:2px 0 0 2px"></div>
              <div style="flex:1;padding:12px 14px;min-width:0;overflow:hidden">
                <div style="font-size:14px;font-weight:600;color:var(--text-primary);line-height:1.4;word-break:break-word;white-space:normal">${_esc(kpi.metric)}</div>
                <div style="font-size:11px;color:var(--text-muted);margin-top:4px;display:flex;gap:6px;flex-wrap:wrap;align-items:center">
                  ${kpi.who ? `<span style="color:var(--brand-accent);font-weight:500">${_esc(kpi.who)}</span>` : ''}
                  ${ytdFmt  ? `<span>YTD: <strong style="color:var(--text-secondary)">${ytdFmt}</strong></span>` : `<span style="font-style:italic">No data yet</span>`}
                </div>
                ${kpi.comment ? `<div style="font-size:11px;color:var(--rag-amber);margin-top:4px;line-height:1.4;word-break:break-word">${_esc(kpi.comment)}</div>` : ''}
              </div>
              <div style="display:flex;flex-direction:column;align-items:flex-end;justify-content:center;padding:12px 14px;gap:6px;flex-shrink:0;align-self:flex-start;padding-top:14px">
                <span class="rag-badge ${rag}" style="font-size:10px;padding:2px 7px;white-space:nowrap">${rag.charAt(0).toUpperCase()+rag.slice(1)}</span>
                <span style="color:var(--text-muted);font-size:20px;line-height:1;font-weight:300">›</span>
              </div>
            </div>`;
        }).join('');

    return `
      <div style="margin-bottom:24px" data-section-name="${section.replace(/"/g,'&quot;')}">
        ${_sectionHeader(section, enc)}
        <div class="desktop-data-view" style="overflow-x:auto;border:1px solid var(--border-card);border-radius:10px;background:var(--bg-card)">
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
        <div class="mobile-data-view">
          <div style="border:1px solid var(--border-card);border-radius:10px;overflow:hidden;background:var(--bg-card)">
            ${mobileCards}
          </div>
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
    // Filter by section filter
    if (App._kpiSectionFilter && section !== App._kpiSectionFilter) return '';
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
          💡 <strong style="color:var(--text-primary)">Target vs Colour Rule:</strong>
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
                <option value="last_fy"  ${s.reportingPeriod==='last_fy'?'selected':''}>Last FY ⏳</option>
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
            <button class="btn btn-primary" onclick="App.exportCSV()">↓ Export CSV</button>
            <button class="btn btn-ghost" onclick="App.exportData()">↓ Backup JSON</button>
            <button class="btn btn-ghost" onclick="document.getElementById('json-restore-input').click()">↑ Restore JSON</button>
            <input type="file" id="json-restore-input" accept=".json" style="display:none" onchange="App.handleJsonRestore(event)">
            <button class="btn btn-ghost" style="color:var(--rag-red)" onclick="App.confirmReset()">⚠ Reset All Data</button>
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
        <button class="btn btn-ghost" style="font-size:10px;padding:3px 8px;color:var(--rag-red)" onclick="App.confirmRemoveSection('${enc}')">🗑 Remove</button>
      </div>`;
  }

  function _esc(s) {
    return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // ── FORMULAS PAGE ────────────────────────────────────────────────────────
  function formulas() {
    const allKpis     = DataStore.getKpis();
    const formulaKpis = allKpis.filter(k => k.isFormula);
    const sourceKpis  = allKpis.filter(k => !k.isFormula);
    const fmlDefs     = DataStore.getFormulas();
    const fyMonths    = DataStore.getFyMonths();
    const settings    = DataStore.getSettings();
    const mode        = settings.reportingPeriod || 'monthly';

    const opSymbols = { sum:'+', subtract:'−', multiply:'×', divide:'÷', avg:'avg', min:'min', max:'max', custom:'ƒ' };
    const opLabels  = { sum:'Sum', subtract:'Subtract', multiply:'Multiply', divide:'Divide', avg:'Average', min:'Min', max:'Max', custom:'Custom' };

    if (sourceKpis.length === 0) return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
          <div><h2 class="page-title" style="margin:0 0 4px">Formula KPIs</h2>
               <div style="font-size:13px;color:var(--text-secondary)">Computed KPIs built from other KPIs</div></div>
        </div>
        <div class="card" style="text-align:center;padding:48px">
          <div style="font-size:40px;margin-bottom:12px">⚠</div>
          <div style="font-size:15px;font-weight:600;margin-bottom:6px">No source KPIs yet</div>
          <div style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">Add KPIs in Data Entry first — formulas compute their values from existing KPIs.</div>
          <button class="btn btn-primary" onclick="App.navigate('data-entry')">✎ Go to Data Entry</button>
        </div>
      </div>`;

    // ── Search bar state (module-level var for filter)
    const tableRows = formulaKpis.map(kpi => {
      const fml         = fmlDefs.find(f => f.kpiId === kpi.id);
      const ps          = DataStore.getPeriodStats(kpi, mode);
      const ragCol      = { green:'var(--rag-green)', amber:'var(--rag-amber)', red:'var(--rag-red)', neutral:'var(--rag-neutral)' }[kpi.rag||'neutral'];
      const opSym       = opSymbols[fml?.op] || '+';
      const operandNames= (fml?.operands||[]).map(id=>{const k=allKpis.find(x=>x.id===id);return k?k.metric:'(deleted)';});

      const formulaDisplay = fml?.op === 'custom'
        ? `<span style="font-family:var(--font-mono);font-size:10px;color:var(--text-secondary);word-break:break-all">${_esc(fml.expression||'')}</span>`
        : operandNames.map((n,i)=>`<span style="background:var(--bg-input);border-radius:3px;padding:1px 5px;font-size:10px;white-space:nowrap">${_esc(n)}</span>${i<operandNames.length-1?` <span style="color:var(--brand-accent);font-weight:700;font-size:11px">${opSym}</span> `:''}`).join('');

      // Monthly sparkline mini-bars inline
      const ma   = kpi.monthlyActuals || {};
      const mVals= fyMonths.map(m => ma[m]!==undefined ? parseFloat(ma[m]) : null);
      const mMax = Math.max(...mVals.filter(v=>v!==null), 1);
      const sparkline = `
        <div style="display:flex;gap:1px;align-items:flex-end;height:24px;width:120px">
          ${fyMonths.map((m,i)=>{
            const v = mVals[i];
            const h = v!==null ? Math.max(2,Math.round((v/mMax)*22)) : 0;
            return `<div title="${m}: ${v!==null?DataStore.formatValue(v,kpi):'—'}"
                         style="flex:1;height:${h}px;background:${v!==null?ragCol:'var(--border-subtle)'};border-radius:1px 1px 0 0;
                                opacity:0.8;min-width:0;cursor:default"
                         onmouseover="this.style.opacity='1'" onmouseout="this.style.opacity='0.8'"></div>`;
          }).join('')}
        </div>`;

      return `
        <tr data-fml-name="${_esc(kpi.metric.toLowerCase())}" data-fml-section="${_esc(kpi.section.toLowerCase())}">
          <!-- RAG dot -->
          <td style="padding:10px 8px 10px 16px;width:4px;vertical-align:middle">
            <div style="width:8px;height:8px;border-radius:50%;background:${ragCol};flex-shrink:0"></div>
          </td>
          <!-- Name + formula -->
          <td style="padding:10px 8px;vertical-align:middle;min-width:160px">
            <div style="font-size:13px;font-weight:600;color:var(--text-primary);margin-bottom:3px">${_esc(kpi.metric)}</div>
            <div style="font-size:10px;color:var(--text-muted);margin-bottom:4px">${_esc(kpi.section)}${kpi.unit?' · '+kpi.unit:''}</div>
            <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
              <span style="font-size:10px;font-weight:700;padding:1px 6px;border-radius:3px;
                           background:rgba(0,194,168,0.1);color:var(--brand-accent)">∑ ${opLabels[fml?.op]||'Sum'}</span>
              ${formulaDisplay}
            </div>
          </td>
          <!-- Computed value -->
          <td style="padding:10px 8px;vertical-align:middle;text-align:right;white-space:nowrap">
            <div style="font-family:var(--font-display);font-size:16px;font-weight:700;color:${ragCol}">
              ${ps.actual!==null ? DataStore.formatValue(ps.actual,kpi) : '—'}
            </div>
            <div style="font-size:10px;color:var(--text-muted);margin-top:2px">computed</div>
          </td>
          <!-- Target -->
          <td style="padding:10px 8px;vertical-align:middle;text-align:right;white-space:nowrap">
            <div style="font-family:var(--font-display);font-size:14px;font-weight:600;color:var(--text-primary)">
              ${DataStore.formatTarget(kpi)}
            </div>
            <div style="font-size:10px;color:var(--text-muted);margin-top:2px">target</div>
          </td>
          <!-- YTD -->
          <td style="padding:10px 8px;vertical-align:middle;text-align:right;white-space:nowrap">
            <div style="font-family:var(--font-display);font-size:14px;font-weight:600;color:var(--text-secondary)">
              ${kpi.ytd!==null ? DataStore.formatValue(kpi.ytd,kpi) : '—'}
            </div>
            <div style="font-size:10px;color:var(--text-muted);margin-top:2px">YTD</div>
          </td>
          <!-- Sparkline -->
          <td style="padding:10px 8px;vertical-align:middle">
            ${sparkline}
          </td>
          <!-- RAG badge -->
          <td style="padding:10px 8px;vertical-align:middle;text-align:center">
            <span class="rag-badge ${kpi.rag||'neutral'}">${(kpi.rag||'neutral').charAt(0).toUpperCase()+(kpi.rag||'neutral').slice(1)}</span>
            ${kpi.isKey?'<div style="font-size:9px;color:var(--brand-accent-2);margin-top:3px;font-weight:600">★ KEY</div>':''}
          </td>
          <!-- Actions -->
          <td style="padding:10px 16px 10px 8px;vertical-align:middle;white-space:nowrap">
            <div style="display:flex;gap:4px;justify-content:flex-end">
              <button class="btn btn-ghost" style="font-size:11px;padding:4px 9px"
                      onclick="App.openEditFormulaKpi('${kpi.id}')">✎ Edit</button>
              <button class="btn btn-ghost" style="font-size:11px;padding:4px 9px;color:var(--rag-red)"
                      onclick="App.confirmRemoveFormulaKpi('${kpi.id}')">✕</button>
            </div>
          </td>
        </tr>`;
    }).join('');

    const emptyRow = `
      <tr id="fml-empty-row" style="display:none">
        <td colspan="8" style="padding:24px;text-align:center;color:var(--text-muted);font-size:13px">
          ⌕ No formula KPIs match your search
        </td>
      </tr>`;

    const noFormulas = formulaKpis.length === 0 ? `
      <div class="card" style="text-align:center;padding:60px 24px">
        <div style="font-size:48px;margin-bottom:16px">∑</div>
        <div style="font-family:var(--font-display);font-size:18px;font-weight:700;margin-bottom:8px">No Formula KPIs yet</div>
        <div style="font-size:13px;color:var(--text-secondary);max-width:420px;margin:0 auto 20px;line-height:1.6">
          Formula KPIs automatically compute from other KPIs — totals, averages, ratios, and custom expressions.
          Monthly data flows through so charts and trends work automatically.
        </div>
        <button class="btn btn-primary" onclick="App.openAddFormulaKpi()" style="font-size:14px;padding:10px 24px">
          ∑ Create Formula KPI
        </button>
      </div>` : '';

    // Mobile card list
    const mobileList = formulaKpis.length > 0 ? `
      <div class="mobile-data-view card" style="padding:0;overflow:hidden">
        ${formulaKpis.map(kpi => {
          const fml = fmlDefs.find(f => f.kpiId === kpi.id);
          const ps  = DataStore.getPeriodStats(kpi, mode);
          const ragColor = { green:'var(--rag-green)', amber:'var(--rag-amber)', red:'var(--rag-red)', neutral:'var(--rag-neutral)' }[kpi.rag||'neutral'];
          const rag  = kpi.rag || 'neutral';
          const opSym = opSymbols[fml?.op] || '+';
          const operandNames = (fml?.operands||[]).map(id=>{const k=allKpis.find(x=>x.id===id);return k?k.metric:'(deleted)';});
          const formulaShort = fml?.op === 'custom'
            ? `ƒ custom`
            : `${opLabels[fml?.op]||'Sum'}: ${operandNames.slice(0,2).join(` ${opSym} `)}${operandNames.length>2?' …':''}`;
          const valFmt = ps.actual !== null ? DataStore.formatValue(ps.actual, kpi) : '—';
          return `
            <div class="mobile-kpi-card">
              <div style="width:3px;background:${ragColor};align-self:stretch;flex-shrink:0;border-radius:2px 0 0 2px"></div>
              <div style="flex:1;padding:12px 14px;min-width:0;overflow:hidden">
                <div style="font-size:14px;font-weight:600;color:var(--text-primary);line-height:1.4;word-break:break-word">${_esc(kpi.metric)}</div>
                <div style="font-size:11px;color:var(--text-muted);margin-top:3px">${_esc(kpi.section)} · <span style="color:var(--brand-accent);font-weight:500">${_esc(formulaShort)}</span></div>
                <div style="font-size:12px;font-weight:700;color:${ragColor};margin-top:4px;font-family:var(--font-display)">${valFmt} <span style="font-size:10px;font-weight:400;color:var(--text-muted)">computed</span></div>
              </div>
              <div style="display:flex;flex-direction:column;align-items:flex-end;justify-content:center;padding:12px 10px;gap:8px;flex-shrink:0">
                <span class="rag-badge ${rag}" style="font-size:10px;padding:2px 7px">${rag.charAt(0).toUpperCase()+rag.slice(1)}</span>
                <div style="display:flex;gap:6px">
                  <button class="btn btn-ghost" style="font-size:11px;padding:3px 9px"
                          onclick="App.openEditFormulaKpi('${kpi.id}')">✎</button>
                  <button class="btn btn-ghost" style="font-size:11px;padding:3px 9px;color:var(--rag-red)"
                          onclick="App.confirmRemoveFormulaKpi('${kpi.id}')">✕</button>
                </div>
              </div>
            </div>`;
        }).join('')}
      </div>` : '';

    const table = formulaKpis.length > 0 ? `
      <div class="desktop-data-view card" style="padding:0;overflow:hidden">
        <table style="width:100%;border-collapse:collapse" id="fml-table">
          <thead>
            <tr style="border-bottom:1px solid var(--border-subtle)">
              <th style="width:4px;padding:10px 8px 10px 16px"></th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:left">Formula KPI</th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:right">Value</th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:right">Target</th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:right">YTD</th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted)">Trend</th>
              <th style="padding:10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:center">Status</th>
              <th style="padding:10px 16px 10px 8px;font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-muted);text-align:right">Actions</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
            ${emptyRow}
          </tbody>
        </table>
      </div>` : '';

    const howItWorks = `
      <div class="card" style="margin-top:20px;border:1px dashed var(--border-card)">
        <div style="font-family:var(--font-display);font-size:14px;font-weight:700;margin-bottom:12px;display:flex;align-items:center;gap:8px">
          <span style="color:var(--brand-accent)">ⓘ</span> How Formula KPIs work
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:16px;font-size:13px;color:var(--text-secondary);line-height:1.6">
          <div><strong style="color:var(--text-primary)">Auto-computed</strong><br>Values recalculate whenever any source KPI is updated.</div>
          <div><strong style="color:var(--text-primary)">Monthly data flows</strong><br>Each month is computed from source KPI monthly actuals — charts work automatically.</div>
          <div><strong style="color:var(--text-primary)">Targets inherited</strong><br>Targets use the same formula applied to source KPI targets.</div>
          <div><strong style="color:var(--text-primary)">Dashboard-ready</strong><br>Toggle ★ Key KPI to pin a formula result on your Overview.</div>
        </div>
      </div>`;

    return `
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
          <div>
            <h2 class="page-title" style="margin:0 0 4px">Formula KPIs</h2>
            <div style="font-size:13px;color:var(--text-secondary)">Computed KPIs built from other KPIs · ${formulaKpis.length} total</div>
          </div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
            ${formulaKpis.length > 0 ? `
            <div style="position:relative">
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none">⌕</span>
              <input type="text" id="fml-page-search" class="input-field"
                     style="padding-left:30px;width:220px;font-size:13px"
                     placeholder="Search formulas…"
                     oninput="App.filterFormulaPage(this.value)">
            </div>` : ''}
            <button class="btn btn-primary" onclick="App.openAddFormulaKpi()" style="gap:8px">
              <span style="font-size:16px">∑</span> New Formula KPI
            </button>
          </div>
        </div>
        ${noFormulas}
        ${mobileList}
        ${table}
        ${howItWorks}
      </div>
      <button class="mobile-fab" onclick="App.openAddFormulaKpi()" title="New Formula KPI" style="font-size:22px">∑</button>`;
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
  let _manageKpiDropdownOpen = false;
  let _kpiSectionFilter      = '';  // section filter for advanced data entry
  let _simpleSectionFilter   = '';  // section filter for simple data entry

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
  function autoCalcAnnual(changedField) {
    const mode  = document.querySelector('input[name="e-autocalc"]:checked')?.value || 'none';
    const fyEl  = document.getElementById('e-tgt-fy');
    const moEl  = document.getElementById('e-tgt-month');
    const unit  = document.getElementById('e-unit')?.value || '';
    if (!fyEl || !moEl || mode === 'none') return;
    const isPct = unit === '%';
    function fmt(n) { return n % 1 === 0 ? String(n) : parseFloat(n.toFixed(6)).toString(); }
    // Only calc in one direction based on which field was changed or which mode was just selected
    if (mode === 'fy_from_month' && (changedField === 'month' || changedField === 'mode')) {
      const mo = parseFloat(moEl.value);
      if (!isNaN(mo)) fyEl.value = isPct ? fmt(mo) : fmt(mo * 12);
    } else if (mode === 'month_from_fy' && (changedField === 'fy' || changedField === 'mode')) {
      const fy = parseFloat(fyEl.value);
      if (!isNaN(fy)) moEl.value = isPct ? fmt(fy) : fmt(fy / 12);
    }
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
    const isKey          = document.getElementById('e-iskey')?.checked                               || false;
    const ytdMethod      = document.getElementById('e-ytdmethod')?.value                             || 'sum';
    const autoCalcTargets = document.querySelector('input[name="e-autocalc"]:checked')?.value       || 'none';

    const fields = {
      metric, section, who, unit,
      targetFY26:      tgtFY  !== '' ? parseFloat(tgtFY)  : null,
      targetMonth:     tgtMo  !== '' ? parseFloat(tgtMo)  : null,
      targetFY26Op:    tgtOp  || null,
      ytd:             ytdVal !== '' ? parseFloat(ytdVal)  : null,
      thresholdId:     thId,
      ragOverride:     ragOv  || null,
      comment, isKey, ytdMethod, autoCalcTargets,
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
    } else if (field === 'ytdMethod') {
      DataStore.updateKpi(id, { ytdMethod: v });
      DataStore.recomputeKpiYtd(id);
      return;
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

        // ── Month name lookup ────────────────────────────────────────────────
        const MONTH_MAP = {
          'jan':'Jan','january':'Jan','feb':'Feb','february':'Feb',
          'mar':'Mar','march':'Mar','apr':'Apr','april':'Apr','may':'May',
          'jun':'Jun','june':'Jun','jul':'Jul','july':'Jul',
          'aug':'Aug','august':'Aug','sep':'Sep','sept':'Sep','september':'Sep',
          'oct':'Oct','october':'Oct','nov':'Nov','november':'Nov',
          'dec':'Dec','december':'Dec',
        };

        // ── Find header row ──────────────────────────────────────────────────
        let headerRow = -1, colMap = {}, monthCols = [];
        for (let i = 0; i < Math.min(8, rawRows.length); i++) {
          const row = (rawRows[i]||[]).map(c => String(c||'').toLowerCase().trim());
          if (row.some(c => c.includes('metric') || c.includes('kpi'))) {
            headerRow = i;
            row.forEach((c, idx) => {
              if      (c.includes('metric') || c === 'kpi')          colMap.metric      = idx;
              else if (c.includes('section') || c.includes('group')) colMap.section     = idx;
              else if (c.includes('who') || c.includes('owner'))     colMap.who         = idx;
              else if (c.includes('unit'))                            colMap.unit        = idx;
              else if (c.includes('label') || c.includes('display')) colMap.targetLabel = idx;
              // ytd must be checked before 'fy' so "ytd fy26" maps to ytd, not targetFY26
              else if (c.includes('ytd') || c.includes('year to'))   colMap.ytd         = idx;
              else if (c.includes('actual') || c.includes('current'))colMap.actual      = idx;
              else if ((c.includes('fy') || c.includes('annual')) && !c.includes('ytd')) colMap.targetFY26  = idx;
              else if ((c.includes('target') && c.includes('month')) || (c === 'target / month') || (c.includes('monthly') && !c.includes('actual'))) colMap.targetMonth = idx;
              else if (c.includes('rag') || c.includes('status'))    colMap.rag         = idx;
              else if (c.includes('comment') || c.includes('note'))  colMap.comment     = idx;
              if (MONTH_MAP[c]) monthCols.push({ colIdx: idx, month: MONTH_MAP[c] });
            });
            break;
          }
        }
        if (headerRow < 0) {
          headerRow = 1;
          colMap = { metric:0, who:1, targetFY26:2, targetMonth:3, ytd:4 };
        }
        // If no explicit month headers found, assume FY order from column F (index 5)
        if (monthCols.length === 0) {
          DataStore.getFyMonths().forEach((m, i) => monthCols.push({ colIdx: 5 + i, month: m }));
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
          const hasTgt  = colMap.targetFY26  !== undefined && row[colMap.targetFY26]  !== null && row[colMap.targetFY26]  !== '';
          const hasTgtM = colMap.targetMonth !== undefined && row[colMap.targetMonth] !== null && row[colMap.targetMonth] !== '';
          const hasWho  = colMap.who         !== undefined && row[colMap.who];
          const hasYtd  = colMap.ytd         !== undefined && row[colMap.ytd]         !== null && row[colMap.ytd]         !== '';
          if (!hasTgt && !hasTgtM && !hasWho && !hasYtd) { currentSection = metricRaw; continue; }

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

          // Parse month columns (F onwards)
          const monthlyActuals = {};
          monthCols.forEach(({ colIdx, month }) => {
            if (row[colIdx] !== null && row[colIdx] !== undefined) {
              const cell = parseCell(row[colIdx], fmtRow?.[colIdx]);
              if (!isNaN(cell.value)) monthlyActuals[month] = cell.value;
            }
          });

          rows.push({
            section:      row[colMap.section] ? String(row[colMap.section]).trim() : currentSection,
            metric:       metricRaw,
            who:          colMap.who !== undefined && row[colMap.who] ? String(row[colMap.who]).trim() : '',
            unit,
            targetLabel,
            targetFY26:   tgtFY.value,
            targetFY26Op: tgtFY.op,
            targetMonth:  tgtMo.value,
            ytd:          ytdC.value,
            actual:       actC.value,
            rag:          colMap.rag !== undefined && row[colMap.rag]
                            ? String(row[colMap.rag]).toLowerCase().trim() : null,
            comment:      colMap.comment !== undefined ? (row[colMap.comment] || '') : '',
            monthlyActuals,
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

    const MONTH_MAP = {
      'jan':'Jan','january':'Jan','feb':'Feb','february':'Feb',
      'mar':'Mar','march':'Mar','apr':'Apr','april':'Apr','may':'May',
      'jun':'Jun','june':'Jun','jul':'Jul','july':'Jul',
      'aug':'Aug','august':'Aug','sep':'Sep','sept':'Sep','september':'Sep',
      'oct':'Oct','october':'Oct','nov':'Nov','november':'Nov',
      'dec':'Dec','december':'Dec',
    };

    // Detect month columns from header; fall back to FY order from col F (index 5)
    let monthCols = [];
    header.forEach((h, idx) => { if (MONTH_MAP[h]) monthCols.push({ colIdx: idx, month: MONTH_MAP[h] }); });
    if (monthCols.length === 0) {
      DataStore.getFyMonths().forEach((m, i) => monthCols.push({ colIdx: 5 + i, month: m }));
    }

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
      const tgtMo  = parseCell(cols.targetMonth >= 0 ? c[cols.targetMonth] : c[3]);
      const ytdC   = parseCell(cols.ytd         >= 0 ? c[cols.ytd]         : c[4]);
      const actC   = parseCell(cols.actual      >= 0 ? c[cols.actual]      : '');

      let unit = '';
      if (cols.unit >= 0 && c[cols.unit]) unit = c[cols.unit].trim();
      else unit = tgtFY.unit || tgtMo.unit || ytdC.unit || actC.unit;

      let targetLabel = '';
      if (cols.targetLabel >= 0 && c[cols.targetLabel]) targetLabel = c[cols.targetLabel].trim();
      else if (tgtFY.label) targetLabel = tgtFY.label;

      // Parse month columns
      const monthlyActuals = {};
      monthCols.forEach(({ colIdx, month }) => {
        if (c[colIdx] !== undefined && c[colIdx] !== '') {
          const cell = parseCell(c[colIdx]);
          if (!isNaN(cell.value)) monthlyActuals[month] = cell.value;
        }
      });

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
        monthlyActuals,
      };
    }).filter(Boolean);
  }

  // ── Export / Reset ────────────────────────────────────────────────────────
  function exportCSV() {
    const kpis      = DataStore.getKpis();
    const fyMonths  = DataStore.getFyMonths();
    const settings  = DataStore.getSettings();
    const fyLabel   = settings.fiscalYearLabel || 'FY';

    // Escape a CSV cell value
    const esc = v => {
      const s = v === null || v === undefined ? '' : String(v);
      return (s.includes(',') || s.includes('"') || s.includes('\n'))
        ? `"${s.replace(/"/g,'""')}"` : s;
    };

    // Header: A-E fixed, then one column per FY month
    const headerRow = ['Metric','Who',`Target ${fyLabel}`,'Target Month','YTD',...fyMonths];

    const rows = [headerRow];

    // Group by section, emit section header rows as recognised by the importer
    const sections = DataStore.getSections();
    sections.forEach(section => {
      const sKpis = DataStore.getKpisBySection(section);
      if (!sKpis.length) return;
      // Section header row (col A = section name, rest empty — importer uses this as section signal)
      rows.push([section, ...Array(4 + fyMonths.length).fill('')]);
      sKpis.forEach(kpi => {
        rows.push([
          kpi.metric,
          kpi.who || '',
          kpi.targetFY26  ?? '',
          kpi.targetMonth ?? '',
          kpi.ytd         ?? '',
          ...fyMonths.map(m => kpi.monthlyActuals?.[m] ?? ''),
        ]);
      });
    });

    const csv  = rows.map(r => r.map(esc).join(',')).join('\r\n');
    const blob = new Blob([csv], { type:'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = `exec-dashboard-${fyLabel}.csv`; a.click();
    URL.revokeObjectURL(url);
  }

  function exportData() {
    const data = {
      settings:   DataStore.getSettings(),
      thresholds: DataStore.getThresholds(),
      formulas:   DataStore.getFormulas(),
      kpis:       DataStore.getKpis(),
      exportedAt: new Date().toISOString(),
    };
    const blob = new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href=url; a.download='exec-dashboard-backup.json'; a.click(); URL.revokeObjectURL(url);
  }

  function handleJsonRestore(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    event.target.value = '';
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = JSON.parse(e.target.result);
        if (!data.kpis && !data.settings && !data.thresholds && !data.formulas) {
          showToast('⚠ Invalid backup file', 'amber'); return;
        }
        if (!confirm(`Restore backup from ${data.exportedAt ? new Date(data.exportedAt).toLocaleString() : 'unknown date'}?\nThis will overwrite all current data.`)) return;
        DataStore.restoreFromBackup(data);
        navigate('overview');
        showToast('✓ Backup restored');
      } catch {
        showToast('⚠ Could not read backup file', 'amber');
      }
    };
    reader.readAsText(file);
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

  // ── Section filter ────────────────────────────────────────────────────────
  function toggleSectionFilterDropdown(mode) {
    const wrId = mode === 'adv' ? 'sec-filter-adv-wrap' : 'sec-filter-simple-wrap';
    const ddId = 'sec-filter-dd-' + mode;
    let dd = document.getElementById(ddId);
    if (dd) { dd.remove(); return; }
    const wrap = document.getElementById(wrId);
    if (!wrap) return;
    const sections = DataStore.getSections();
    const current  = mode === 'adv' ? _kpiSectionFilter : _simpleSectionFilter;
    dd = document.createElement('div');
    dd.id = ddId;
    dd.style.cssText = 'position:fixed;z-index:300;background:var(--bg-modal);border:1px solid var(--border-card);border-radius:10px;box-shadow:var(--shadow-modal);min-width:200px;overflow:hidden';
    const rect = wrap.getBoundingClientRect();
    dd.style.top  = (rect.bottom + 4) + 'px';
    dd.style.left = rect.left + 'px';
    const makeItem = (label, val) => {
      const el = document.createElement('div');
      const active = current === val;
      el.style.cssText = `padding:9px 14px;font-size:12px;cursor:pointer;display:flex;align-items:center;gap:8px;color:${active?'var(--brand-accent)':'var(--text-secondary)'};background:${active?'rgba(0,194,168,0.07)':''}`;
      el.onmouseover = () => { if (!active) el.style.background='var(--bg-card-hover)'; };
      el.onmouseout  = () => { if (!active) el.style.background=''; };
      el.innerHTML   = `<span style="width:13px;text-align:center">${active?'✓':''}</span>${label}`;
      el.onclick = () => {
        dd.remove();
        if (mode === 'adv') setSectionFilter(val);
        else                setSimpleSectionFilter(val);
      };
      return el;
    };
    dd.appendChild(makeItem('All Sections', ''));
    const divider = document.createElement('div');
    divider.style.cssText = 'height:1px;background:var(--border-subtle);margin:2px 0';
    dd.appendChild(divider);
    sections.forEach(s => dd.appendChild(makeItem(s, s)));
    document.body.appendChild(dd);
    setTimeout(() => {
      const close = (e) => { if (!dd.contains(e.target) && !wrap.contains(e.target)) { dd.remove(); document.removeEventListener('click', close); } };
      document.addEventListener('click', close);
    }, 0);
  }

  function setSectionFilter(section) {
    _kpiSectionFilter = section;
    render();
  }

  function setSimpleSectionFilter(section) {
    _simpleSectionFilter = section;
    applySimpleSectionFilter();
    // Re-render to update button label/state
    render();
  }

  function applySimpleSectionFilter() {
    const section = _simpleSectionFilter;
    // Section blocks are wrapper divs around each [data-section-table]
    // They have a sibling _sectionHeader above — target the wrapping div by its table's section
    document.querySelectorAll('[data-section-name]').forEach(block => {
      const name = block.getAttribute('data-section-name');
      block.style.display = (!section || name === section) ? '' : 'none';
    });
  }

  function toggleAdvancedEntry() {
    // On mobile, show a recommendation before entering monthly entry mode
    if (!_advancedMode && window.innerWidth <= 768) {
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:500;display:flex;align-items:center;justify-content:center;padding:24px 16px';
      overlay.innerHTML = `
        <div style="background:var(--bg-modal);border:1px solid var(--border-card);border-radius:16px;padding:24px 20px;max-width:360px;width:calc(100% - 32px);box-shadow:var(--shadow-modal);text-align:center">
          <div style="font-size:28px;margin-bottom:12px">🖥️</div>
          <div style="font-size:15px;font-weight:700;color:var(--text-primary);margin-bottom:8px">Desktop Recommended</div>
          <div style="font-size:13px;color:var(--text-secondary);line-height:1.5;margin-bottom:20px">Monthly data entry works best on a larger screen. You can still continue on mobile, but it may be cramped.</div>
          <div style="display:flex;gap:10px;justify-content:center">
            <button id="_adv-cancel" class="btn btn-ghost" style="flex:1">Go Back</button>
            <button id="_adv-continue" class="btn btn-primary" style="flex:1">Continue Anyway</button>
          </div>
        </div>`;
      document.body.appendChild(overlay);
      overlay.querySelector('#_adv-cancel').onclick  = () => overlay.remove();
      overlay.querySelector('#_adv-continue').onclick = () => { overlay.remove(); _advancedMode = true; render(); };
      overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
      return;
    }
    _advancedMode = !_advancedMode;
    if (!_advancedMode) _kpiSearchQuery = '';
    _kpiSectionFilter = ''; _simpleSectionFilter = '';
    render();
  }

  // ── Navigate KPI card → advanced data entry ───────────────────────────────
  function openKpiInDataEntry(id) {
    _advancedMode = true;
    _dataEntryTab = 'kpis';
    _expandedKpiId = id; // expand monthly row for this KPI on arrival
    navigate('data-entry');
    // Scroll to the exact KPI row after render — use requestAnimationFrame inside
    // the timeout to ensure layout is complete before measuring position
    setTimeout(() => {
      requestAnimationFrame(() => {
        const el = document.querySelector(`tr[data-kpi-id="${id}"]`);
        if (!el) return;
        const top = el.getBoundingClientRect().top + window.scrollY - 70;
        window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
      });
    }, 150);
  }

  // ── Manage KPIs dropdown (Overview) ───────────────────────────────────────
  function _openManageKpiDropdown() {
    const dd   = document.getElementById('manage-kpi-dropdown');
    const wrap = document.getElementById('manage-kpi-wrap');
    if (!dd || !wrap) return;
    const rect  = wrap.getBoundingClientRect();
    const menuW = Math.min(320, window.innerWidth - 16);
    let left = rect.right - menuW;
    if (left < 8) left = 8;
    dd.style.position = 'fixed';
    dd.style.left     = left + 'px';
    dd.style.right    = 'auto';
    dd.style.top      = (rect.bottom + 6) + 'px';
    dd.style.width    = menuW + 'px';
    dd.style.display  = 'block';
    _manageKpiDropdownOpen = true;
  }

  function toggleManageKpiDropdown() {
    const dd = document.getElementById('manage-kpi-dropdown');
    if (!dd) return;
    const isOpen = dd.style.display !== 'none';
    if (isOpen) {
      dd.style.display = 'none';
      _manageKpiDropdownOpen = false;
    } else {
      _openManageKpiDropdown();
      setTimeout(() => document.getElementById('manage-kpi-search')?.focus(), 50);
      // Close on outside click — always look up elements fresh so stale refs don't confuse post-render clicks
      setTimeout(() => {
        const handler = (e) => {
          if (!e.target.isConnected) return; // click on a DOM node removed by re-render — ignore
          const wrap = document.getElementById('manage-kpi-wrap');
          const ddEl = document.getElementById('manage-kpi-dropdown');
          if (wrap && !wrap.contains(e.target) && (!ddEl || !ddEl.contains(e.target))) {
            if (ddEl) ddEl.style.display = 'none';
            _manageKpiDropdownOpen = false;
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
    // setOverviewKpiIds triggers a full re-render (closes dropdown) — reopen after
    const wasOpen = _manageKpiDropdownOpen;
    DataStore.setOverviewKpiIds(ids);
    if (wasOpen) _openManageKpiDropdown();
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
    // setOverviewKpiIds triggers a full re-render (closes dropdown) — reopen after
    const wasOpen = _manageKpiDropdownOpen;
    DataStore.setOverviewKpiIds(ids);
    if (wasOpen) _openManageKpiDropdown();
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
    const items = list.querySelectorAll('div.mkpi-row');
    const query = q.toLowerCase().trim();
    let visible = 0;
    items.forEach(div => {
      const nameEl = div.querySelector('div > div:first-child');
      const sectEl = div.querySelector('div > div:last-child');
      const name = nameEl?.textContent?.toLowerCase() || '';
      const sect = sectEl?.textContent?.toLowerCase() || '';
      const show = !query || name.includes(query) || sect.includes(query);
      div.style.display = show ? 'flex' : 'none';
      if (show) {
        div.style.borderTop = visible === 0 ? 'none' : '1px solid var(--border-subtle)';
        visible++;
      }
    });
  }

  // ── Formula KPI management ────────────────────────────────────────────────
  function filterFormulaPage(q) {
    const query = q.toLowerCase().trim();
    const rows  = document.querySelectorAll('#fml-table tbody tr[data-fml-name]');
    let visible = 0;
    rows.forEach(row => {
      const name = row.dataset.fmlName || '';
      const sect = row.dataset.fmlSection || '';
      const show = !query || name.includes(query) || sect.includes(query);
      row.style.display = show ? '' : 'none';
      if (show) visible++;
    });
    const emptyRow = document.getElementById('fml-empty-row');
    if (emptyRow) emptyRow.style.display = visible === 0 && rows.length > 0 ? '' : 'none';
    const box = document.getElementById('fml-page-search');
    if (box) { box.focus(); box.setSelectionRange(q.length, q.length); }
  }

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
    get _kpiSectionFilter()    { return _kpiSectionFilter; },
    get _simpleSectionFilter() { return _simpleSectionFilter; },
    toggleSectionFilterDropdown, setSectionFilter, setSimpleSectionFilter, applySimpleSectionFilter,
    toggleManageKpiDropdown, toggleOverviewKpi, addKpiToOverview, setAllOverviewKpis, filterManageKpiList,
    updateFmtPreview,
    openKpiDetail, openEditKpi, openAddKpi,
    openAddThreshold, openEditThreshold, openAddSectionModal, saveNewSection,
    promptInsertSection, confirmInsertSection,
    updateMonthlyActual,
    autoCalcAnnual, saveKpiEdit, quickUpdate,
    confirmRemoveKpi, promptRenameSection, confirmRemoveSection,
    addThresholdLevel, saveThreshold, confirmRemoveThreshold,
    openAddFormulaKpi, openEditFormulaKpi, saveFormulaKpi, confirmRemoveFormulaKpi, filterFormulaPage,
    saveSettings, handleXlsxUpload, handleJsonRestore, exportCSV, exportData, confirmReset, showToast, showModal, closeModal,
  };
})();
