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
              <i class="fa-solid fa-table-cells"></i> Manage KPIs
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
