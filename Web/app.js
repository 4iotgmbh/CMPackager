'use strict';
(function () {

  // ── State ───────────────────────────────────────────────────────────────
  const state = {
    activeTab:    'recipes',
    recipes:      { enabled: [], disabled: [] },
    running:      false,
    prefsExists:  true,
    outputCount:  0,          // total lines seen (for SSE reconnect ?from=N)
    sseSource:    null,
    logLineCount: 0,
    tests:        { rows: [], file: null, available: false },
    sortCol:      'Timestamp',
    sortDir:      'desc',
    sccmLoaded:   false,
    sccm:         { available: false, apps: [] },
  };

  // ── DOM shortcuts ────────────────────────────────────────────────────────
  const $ = id => document.getElementById(id);

  // ── Helpers ─────────────────────────────────────────────────────────────
  function esc(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function showToast(msg, type = 'info') {
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.textContent = msg;
    $('toast-container').appendChild(el);
    setTimeout(() => el.remove(), 4000);
  }

  async function fetchJSON(url, opts) {
    try {
      const r = await fetch(url, opts);
      if (!r.ok) {
        let errMsg = `HTTP ${r.status}`;
        try { const j = await r.json(); errMsg = j.error || errMsg; } catch {}
        showToast(errMsg, 'error');
        return null;
      }
      return await r.json();
    } catch (e) {
      showToast(e.message, 'error');
      return null;
    }
  }

  // ── Tab switching ────────────────────────────────────────────────────────
  function switchTab(name) {
    state.activeTab = name;
    document.querySelectorAll('.tab-btn').forEach(b =>
      b.classList.toggle('active', b.dataset.tab === name));
    document.querySelectorAll('.tab-content').forEach(t =>
      t.classList.toggle('active', t.dataset.tab === name));

    if (name === 'tests') loadTests();
    if (name === 'sccm' && !state.sccmLoaded) loadSccm(false);
  }

  document.querySelectorAll('.tab-btn').forEach(b =>
    b.addEventListener('click', () => switchTab(b.dataset.tab)));

  // ── Status polling ───────────────────────────────────────────────────────
  async function pollStatus() {
    const s = await fetchJSON('/api/status');
    if (!s) return;

    // Prefs banner
    if (!s.prefsExists && state.prefsExists) {
      state.prefsExists = false;
      $('prefs-banner').classList.remove('banner-hide');
      $('btn-run-all').disabled = true;
    } else if (s.prefsExists && !state.prefsExists) {
      state.prefsExists = true;
      $('prefs-banner').classList.add('banner-hide');
      $('btn-run-all').disabled = false;
    }

    // Running state
    if (s.running !== state.running) {
      state.running = s.running;
      updateRunUI();
    }

    // Seed output if SSE hasn't fired yet (e.g., page load during a run)
    if (s.lastLines && s.lastLines.length && state.outputCount === 0 && state.running) {
      s.lastLines.forEach(l => appendOutput(l));
    }
  }

  function updateRunUI() {
    const pill  = $('status-pill');
    const label = $('status-label');
    const stop  = $('btn-stop');
    const runAll = $('btn-run-all');

    if (state.running) {
      pill.classList.add('running');
      label.textContent = 'Running';
      stop.style.display = '';
      runAll.style.display = 'none';
      $('output-meta').textContent = 'CMPackager is running…';
    } else {
      pill.classList.remove('running');
      label.textContent = 'Idle';
      stop.style.display = 'none';
      runAll.style.display = '';
      if ($('output-meta').textContent === 'CMPackager is running…') {
        $('output-meta').textContent = 'Run finished.';
      }
    }

    // Disable recipe Run buttons while running
    document.querySelectorAll('.btn-run-single').forEach(b => {
      b.disabled = state.running;
    });
  }

  // ── Recipes ──────────────────────────────────────────────────────────────
  async function loadRecipes() {
    const data = await fetchJSON('/api/recipes');
    if (!data) return;
    state.recipes = data;
    renderRecipes();
  }

  function renderRecipes() {
    const { enabled, disabled } = state.recipes;
    $('count-enabled').textContent  = enabled.length;
    $('count-disabled').textContent = disabled.length;

    renderRecipeList($('list-enabled'),  enabled,  'enabled');
    renderRecipeList($('list-disabled'), disabled, 'disabled');
  }

  function renderRecipeList(container, recipes, side) {
    container.innerHTML = '';
    if (!recipes.length) {
      const empty = document.createElement('div');
      empty.className = 'recipe-empty';
      empty.textContent = side === 'enabled' ? 'No enabled recipes.' : 'No disabled recipes.';
      container.appendChild(empty);
      return;
    }

    recipes.forEach(r => {
      const card = document.createElement('div');
      card.className = 'recipe-card';

      const info = document.createElement('div');
      info.className = 'recipe-info';
      info.innerHTML =
        `<div class="recipe-name">${esc(r.appName || r.file)}</div>` +
        `<div class="recipe-meta">${esc(r.file)}${r.publisher ? ' &bull; ' + esc(r.publisher) : ''}</div>`;

      const actions = document.createElement('div');
      actions.className = 'recipe-actions';

      if (side === 'enabled') {
        const btnRun = document.createElement('button');
        btnRun.className = 'btn btn-sm btn-primary btn-run-single';
        btnRun.textContent = 'Run';
        btnRun.disabled = state.running;
        btnRun.addEventListener('click', () => runSingle(r.file));

        const btnDis = document.createElement('button');
        btnDis.className = 'btn btn-sm btn-ghost';
        btnDis.textContent = 'Disable';
        btnDis.addEventListener('click', () => toggleRecipe('disable', r.file));

        actions.appendChild(btnRun);
        actions.appendChild(btnDis);
      } else {
        const btnEn = document.createElement('button');
        btnEn.className = 'btn btn-sm btn-success';
        btnEn.textContent = 'Enable';
        btnEn.addEventListener('click', () => toggleRecipe('enable', r.file));
        actions.appendChild(btnEn);
      }

      card.appendChild(info);
      card.appendChild(actions);
      container.appendChild(card);
    });
  }

  async function toggleRecipe(action, file) {
    const r = await fetchJSON(`/api/${action}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file }),
    });
    if (r?.ok) {
      showToast(`${action === 'enable' ? 'Enabled' : 'Disabled'}: ${file}`, 'success');
      loadRecipes();
    }
  }

  // ── Run controls ──────────────────────────────────────────────────────────
  async function runAll() {
    if (state.running) return;
    clearOutput();
    switchTab('output');
    const r = await fetchJSON('/api/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mode: 'all' }),
    });
    if (r?.ok) {
      state.running = true;
      updateRunUI();
      $('output-meta').textContent = 'Running all recipes…';
    }
  }

  async function runSingle(file) {
    if (state.running) return;
    clearOutput();
    switchTab('output');
    const r = await fetchJSON('/api/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mode: 'single', recipe: file }),
    });
    if (r?.ok) {
      state.running = true;
      updateRunUI();
      $('output-meta').textContent = `Running: ${file}`;
    }
  }

  async function stopRun() {
    await fetchJSON('/api/stop', { method: 'POST' });
    state.running = false;
    updateRunUI();
    $('output-meta').textContent = 'Stopped by user.';
    showToast('Process stopped.', 'warning');
  }

  function clearOutput() {
    $('terminal').innerHTML = '';
    $('log-lines').innerHTML = '';
    state.outputCount  = 0;
    state.logLineCount = 0;
    $('line-count').textContent = '0 lines';
    reconnectSSE();
  }

  // Expose to HTML onclick attributes
  window.runAll    = runAll;
  window.stopRun   = stopRun;
  window.runSingle = runSingle;
  window.loadTests = loadTests;
  window.loadSccm  = loadSccm;
  window.clearOutput = clearOutput;

  // ── Output / Terminal ─────────────────────────────────────────────────────
  function lineClass(line) {
    const l = line.toLowerCase();
    if (l.includes('[err]') || l.includes('error:') || l.includes('exception'))  return 'err';
    if (l.includes('warning') || l.includes('warn:'))  return 'warn';
    if (l.includes('success') || l.includes('complete') || l.includes('done'))   return 'ok';
    if (l.startsWith(':') || l.trim() === '')  return 'dim';
    return '';
  }

  function appendOutput(line) {
    const term = $('terminal');
    const autoScroll = $('chk-autoscroll').checked;
    const atBottom = term.scrollHeight - term.scrollTop <= term.clientHeight + 60;

    const div = document.createElement('div');
    div.className = 'output-line ' + lineClass(line);
    div.textContent = line;
    term.appendChild(div);

    // Cap DOM nodes to avoid memory growth on very long runs
    while (term.children.length > 2000) term.removeChild(term.firstChild);

    state.outputCount++;
    $('line-count').textContent = `${state.outputCount} lines`;

    if (autoScroll && atBottom) term.scrollTop = term.scrollHeight;
  }

  function appendLogLine(line) {
    if (!line.trim()) return;
    const container = $('log-lines');
    const div = document.createElement('div');
    div.className = 'log-line';
    div.textContent = line;
    container.appendChild(div);
    state.logLineCount++;
    // Auto-scroll log section
    const ls = $('log-section');
    ls.scrollTop = ls.scrollHeight;
    // Cap
    while (container.children.length > 500) container.removeChild(container.firstChild);
  }

  function toggleLogPane() {
    const show = $('chk-showlog').checked;
    $('log-section').style.display = show ? '' : 'none';
  }
  window.toggleLogPane = toggleLogPane;

  // ── SSE ───────────────────────────────────────────────────────────────────
  function connectSSE() {
    if (state.sseSource) {
      state.sseSource.close();
      state.sseSource = null;
    }

    const url = `/api/stream?from=${state.outputCount}`;
    const src = new EventSource(url);
    state.sseSource = src;

    // Default event = process output line
    src.onmessage = e => {
      if (e.data && e.data !== ': heartbeat') appendOutput(e.data);
    };

    // Log file tail events
    src.addEventListener('log', e => appendLogLine(e.data));

    // Server sends updated buffer index so reconnect skips duplicates
    src.addEventListener('index', e => {
      const n = parseInt(e.data, 10);
      if (!isNaN(n)) state.outputCount = n;
    });

    src.onerror = () => {
      // EventSource will automatically reconnect; nothing to do here.
      // If server is down, onmessage stops firing — status poll will catch that.
    };
  }

  function reconnectSSE() {
    connectSSE();
  }

  // ── Test Results ───────────────────────────────────────────────────────────
  async function loadTests() {
    const data = await fetchJSON('/api/tests');
    if (!data) return;
    state.tests = data;
    renderTests();
  }

  function resultBadgeClass(result) {
    if (!result) return 'badge-unknown';
    const r = result.toUpperCase();
    if (r.startsWith('PASS'))    return 'badge-pass';
    if (r === 'FAIL')            return 'badge-fail';
    if (r === 'TIMEOUT')         return 'badge-timeout';
    if (r === 'SKIPPED')         return 'badge-skipped';
    if (r === 'ERROR')           return 'badge-error';
    return 'badge-unknown';
  }

  const TEST_COLS = [
    { key: 'Timestamp',              label: 'Time',          cls: 'mono' },
    { key: 'Application',            label: 'Application',   cls: '' },
    { key: 'DeploymentType',         label: 'Deploy Type',   cls: '' },
    { key: 'Result',                 label: 'Result',        cls: 'center', badge: true },
    { key: 'DurationMinutes',        label: 'Min',           cls: 'mono center' },
    { key: 'InstallExitCode',        label: 'Install',       cls: 'mono center' },
    { key: 'DetectionAfterInstall',  label: 'Det.Install',   cls: 'center', bool: true },
    { key: 'UninstallExitCode',      label: 'Uninstall',     cls: 'mono center' },
    { key: 'DetectionAfterUninstall',label: 'Det.Uninstall', cls: 'center', bool: true },
    { key: 'Notes',                  label: 'Notes',         cls: '' },
  ];

  function renderTests() {
    const body = $('tests-body');
    const meta = $('tests-meta');

    if (!state.tests.available || !state.tests.rows.length) {
      meta.textContent = state.tests.file
        ? `File: ${state.tests.file} — no rows`
        : 'No test results found.';
      body.innerHTML = `<div class="empty-state">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="var(--border)" stroke-width="1.5">
          <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2"/>
          <rect x="9" y="3" width="6" height="4" rx="1"/>
        </svg>
        <p>Run <code>Test-RecipeBatch.ps1</code> to generate results.</p>
      </div>`;
      return;
    }

    meta.textContent = `${state.tests.file} — ${state.tests.rows.length} row(s)`;

    // Sort
    const rows = [...state.tests.rows].sort((a, b) => {
      const av = a[state.sortCol] ?? '';
      const bv = b[state.sortCol] ?? '';
      const cmp = String(av).localeCompare(String(bv), undefined, { numeric: true });
      return state.sortDir === 'asc' ? cmp : -cmp;
    });

    // Build table
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';

    const tbl = document.createElement('table');
    tbl.className = 'data-table';

    // Head
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    TEST_COLS.forEach(col => {
      const th = document.createElement('th');
      const isSorted = col.key === state.sortCol;
      th.classList.toggle('sorted', isSorted);
      th.innerHTML = `${esc(col.label)}<span class="sort-ic">${isSorted ? (state.sortDir === 'asc' ? '▲' : '▼') : '⇅'}</span>`;
      th.addEventListener('click', () => {
        if (state.sortCol === col.key) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortCol = col.key;
          state.sortDir = 'asc';
        }
        renderTests();
      });
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    tbl.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    rows.forEach(row => {
      const tr = document.createElement('tr');
      TEST_COLS.forEach(col => {
        const td = document.createElement('td');
        if (col.cls) td.className = col.cls;
        const val = row[col.key] ?? '';

        if (col.badge) {
          td.innerHTML = `<span class="badge ${resultBadgeClass(val)}">${esc(val || '—')}</span>`;
        } else if (col.bool) {
          const v = String(val).toLowerCase();
          if (v === 'true')  td.innerHTML = '<span style="color:var(--success)">&#10003;</span>';
          else if (v === 'false') td.innerHTML = '<span style="color:var(--danger)">&#10007;</span>';
          else td.textContent = val || '—';
        } else {
          td.textContent = val || '—';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    body.innerHTML = '';
    body.appendChild(wrap);
  }

  // ── SCCM Status ────────────────────────────────────────────────────────────
  async function loadSccm(force = false) {
    if (state.sccmLoaded && !force) return;
    $('sccm-meta').textContent = 'Querying SCCM…';
    $('sccm-body').innerHTML = '<div class="empty-state"><div class="spinner"></div></div>';

    const data = await fetchJSON('/api/sccm');
    if (!data) { $('sccm-meta').textContent = 'Failed to query SCCM.'; return; }

    state.sccm = data;
    state.sccmLoaded = true;
    renderSccm();
  }

  function renderSccm() {
    const { available, apps, message, error } = state.sccm;
    const body = $('sccm-body');
    const meta = $('sccm-meta');

    if (!available) {
      meta.textContent = 'SCCM unavailable';
      body.innerHTML = `<div class="banner banner-info">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
        </svg>
        <span>${esc(message || error || 'ConfigurationManager module not available.')}</span>
      </div>`;
      return;
    }

    meta.textContent = `${apps.length} recipe(s) checked`;
    body.innerHTML = '';

    if (!apps.length) {
      body.innerHTML = '<div class="empty-state"><p>No enabled recipes found.</p></div>';
      return;
    }

    const grid = document.createElement('div');
    grid.className = 'sccm-grid';

    apps.forEach(app => {
      const card = document.createElement('div');
      card.className = 'sccm-app';

      const hdr = document.createElement('div');
      hdr.className = 'sccm-app-header';
      const sccmDisplayName = app.sccmName || app.appName || app.recipe;
      const versionStr = app.version ? ` &bull; v${esc(app.version)}` : '';
      const versionsStr = (app.allVersions > 1) ? ` &bull; ${app.allVersions} versions` : '';
      hdr.innerHTML = `
        <div>
          <div class="sccm-app-name">${esc(sccmDisplayName)}</div>
          <div class="sccm-app-meta">${esc(app.recipe)}${versionStr}${versionsStr}</div>
        </div>
        <div class="flex-row">
          ${app.found
            ? `<span class="badge badge-pass">In SCCM</span>`
            : `<span class="badge badge-fail">Not Found</span>`}
          ${app.deployments && app.deployments.length
            ? `<span class="badge badge-skipped">${app.deployments.length} deployment${app.deployments.length !== 1 ? 's' : ''}</span>`
            : ''}
        </div>`;

      // Toggle expand
      const depWrap = document.createElement('div');
      depWrap.style.display = 'none';

      hdr.addEventListener('click', () => {
        depWrap.style.display = depWrap.style.display === 'none' ? '' : 'none';
      });

      if (app.deployments && app.deployments.length) {
        const dt = document.createElement('table');
        dt.className = 'sccm-deploy-table';

        const head = document.createElement('tr');
        head.innerHTML = '<td style="color:var(--muted);font-size:11px;font-weight:600;padding:6px 16px 4px;text-transform:uppercase;letter-spacing:.4px">Collection</td>' +
                         '<td style="color:var(--muted);font-size:11px;font-weight:600;padding:6px 16px 4px;text-transform:uppercase;letter-spacing:.4px">Purpose</td>' +
                         '<td style="color:var(--muted);font-size:11px;font-weight:600;padding:6px 16px 4px;text-transform:uppercase;letter-spacing:.4px;text-align:right">Stats</td>';
        dt.appendChild(head);

        app.deployments.forEach(dep => {
          const tr = document.createElement('tr');
          const purpose = dep.AssignmentType === 1 ? 'Required' : dep.DesiredConfigType === 1 ? 'Required' : 'Available';
          const statsparts = [];
          if (dep.NumberTotal != null)    statsparts.push(`Total: ${dep.NumberTotal}`);
          if (dep.NumberSuccess != null)  statsparts.push(`OK: ${dep.NumberSuccess}`);
          if (dep.NumberErrors != null)   statsparts.push(`Err: ${dep.NumberErrors}`);
          tr.innerHTML = `
            <td class="col-coll">${esc(dep.CollectionName || '—')}</td>
            <td class="col-type">${esc(purpose)}</td>
            <td class="col-num">${esc(statsparts.join(' | ') || '—')}</td>`;
          dt.appendChild(tr);
        });
        depWrap.appendChild(dt);
      } else if (app.found) {
        depWrap.innerHTML = '<div style="padding:10px 16px;font-size:12px;color:var(--muted)">No deployments configured.</div>';
      }

      card.appendChild(hdr);
      card.appendChild(depWrap);
      grid.appendChild(card);
    });

    body.appendChild(grid);
  }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    loadRecipes();
    connectSSE();
    pollStatus();
    setInterval(pollStatus, 3000);
  }

  document.addEventListener('DOMContentLoaded', init);

})();
