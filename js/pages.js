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
