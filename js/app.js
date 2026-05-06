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
    if (!name) { showToast('Enter a section name', 'amber'); return; }
    DataStore.addSection(name);
    closeModal();
    navigate('section:'+encodeURIComponent(name));
    showToast('Section "'+name+'" created');
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
    if (!name) { showToast('Enter a section name', 'amber'); return; }
    DataStore.splitSectionAt(kpiId, name);
    closeModal();
    showToast('Section "'+name+'" created');
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
    if (!metric) { showToast('KPI Name is required', 'amber'); return; }

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

    if (isNew) { DataStore.addKpi(fields); showToast('KPI added'); }
    else       { DataStore.updateKpi(id, fields); showToast('KPI saved'); }
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
    showToast('Section renamed');
  }

  function confirmRemoveSection(encodedSection) {
    const name  = decodeURIComponent(encodedSection);
    const count = DataStore.getKpisBySection(name).length;
    if (!confirm(`Remove section "${name}" and its ${count} KPI${count!==1?'s':''}? This cannot be undone.`)) return;
    DataStore.removeSection(name);
    if (_currentPage === 'section:'+encodedSection) navigate('overview');
    showToast('Section removed');
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
    if (!name) { showToast('Threshold name is required', 'amber'); return; }
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
    if (isNew) { DataStore.addThreshold(fields); showToast('Threshold created'); }
    else       { DataStore.updateThreshold(id, fields); showToast('Threshold saved'); }
    closeModal();
  }

  function confirmRemoveThreshold(id) {
    const th = DataStore.getThresholdById(id);
    if (!th) return;
    const usageCount = DataStore.getKpis().filter(k=>k.thresholdId===id).length;
    if (!confirm(`Delete threshold "${th.name}"? ${usageCount} KPI${usageCount!==1?'s':''} will revert to Manual.`)) return;
    DataStore.removeThreshold(id);
    showToast('Threshold deleted');
    closeModal();
  }

  // ── KPI remove ────────────────────────────────────────────────────────────
  function confirmRemoveKpi(id) {
    const kpi = DataStore.getKpiById(id);
    if (!kpi) return;
    if (!confirm(`Remove KPI "${kpi.metric}"? This cannot be undone.`)) return;
    DataStore.removeKpi(id);
    closeModal();
    showToast('KPI removed');
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
    showToast('Settings saved');
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
        showToast(`${updated} updated, ${added} added from CSV`);
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
        showToast(`${updated} KPIs updated, ${added} new KPIs added`);
        render();
      } catch(err) {
        console.error('XLSX parse error:', err);
        showToast('Error reading file — check format', 'amber');
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
      DataStore.resetToDefaults(); navigate('overview'); showToast('Reset complete');
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
    if (!metric) { showToast('Name is required', 'amber'); return; }
    const section = document.getElementById('fml-section')?.value?.trim() || 'Formulas';
    const unit    = document.getElementById('fml-unit')?.value || '';
    const isKey   = document.getElementById('fml-iskey')?.checked || false;
    const thId    = document.getElementById('fml-threshold')?.value || 'th_manual_only';
    const op      = document.getElementById('fml-op')?.value || 'sum';
    const comment = document.getElementById('fml-comment')?.value?.trim() || '';

    // Collect selected operand KPI ids (checkboxes)
    const operands = Array.from(document.querySelectorAll('.fml-operand-cb:checked')).map(el => el.value);
    if (operands.length === 0 && op !== 'custom') {
      showToast('Select at least one source KPI', 'amber'); return;
    }

    let expression = '';
    if (op === 'custom') {
      expression = document.getElementById('fml-expression')?.value?.trim() || '';
      if (!expression) { showToast('Enter a custom expression', 'amber'); return; }
    }

    const kpiFields  = { metric, section, unit, isKey, thresholdId: thId, comment };
    const formulaDef = { op, operands, expression };

    if (isEdit) {
      DataStore.updateFormulaKpi(kpiId, kpiFields, formulaDef);
      showToast('Formula KPI updated');
    } else {
      DataStore.addFormulaKpi(kpiFields, formulaDef);
      showToast('Formula KPI created');
    }
    closeModal();
  }

  function confirmRemoveFormulaKpi(kpiId) {
    const kpi = DataStore.getKpiById(kpiId);
    if (!kpi) return;
    if (!confirm(`Remove formula KPI "${kpi.metric}"? This cannot be undone.`)) return;
    DataStore.removeFormulaKpi(kpiId);
    showToast('Formula KPI removed');
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
