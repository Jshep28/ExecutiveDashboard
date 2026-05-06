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
