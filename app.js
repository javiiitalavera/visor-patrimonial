/**
 * VISOR PATRIMONIAL v3.7
 * Añade pestaña "Desglose de Cartera" con:
 *  - Mapeador de etiquetas configurable (persistido en localStorage)
 *  - Área apilada histórica: Cash / RF / RV / Fijos
 *  - Donut de composición actual
 *  - Deuda Ocean restada automáticamente al activo bruto
 */

(() => {
    'use strict';

    // ── 1. MAPA DE CATEGORÍAS (configurable, persiste en localStorage) ──────
    // Bump key cuando cambie DEFAULT_CATMAP para invalidar localStorage antiguo
    const LS_CATMAP = 'visorCatMap_v2';
    const DEFAULT_CATMAP = {
        RV:    ['fondos rv','acciones ib (rv)','pp (rv)','fondos','pp','acciones'],
        RF:    ['fondos rf','ib (rf)'],
        Cash:  ['cuentas','efectivo','caixa'],
        // "deuda ocean" es negativo en el Excel -> incluirlo aqui reduce el total de Fijos
        Fijos: ['x3','ocean','deuda ocean']
    };

    function loadCatMap() {
        try {
            const raw = localStorage.getItem(LS_CATMAP);
            if (raw) {
                const saved = JSON.parse(raw);
                // Merge: respeta customizaciones del usuario pero garantiza todas las clases
                const merged = JSON.parse(JSON.stringify(DEFAULT_CATMAP));
                for (const cls of Object.keys(merged)) {
                    if (saved[cls] && saved[cls].length) merged[cls] = saved[cls];
                }
                return merged;
            }
        } catch(e) {}
        return JSON.parse(JSON.stringify(DEFAULT_CATMAP));
    }
    function saveCatMap(map) {
        try { localStorage.setItem(LS_CATMAP, JSON.stringify(map)); } catch(e) {}
    }
    // Persiste el mapa migrado inmediatamente (invalida versiones antiguas)
    let catMap = loadCatMap();
    saveCatMap(catMap);

    // Colores por clase
    const CAT_COLORS = {
        RV:    { border: '#00ffaa', bg: 'rgba(0,255,170,0.75)' },
        RF:    { border: '#58a6ff', bg: 'rgba(88,166,255,0.75)' },
        Cash:  { border: '#ffcc00', bg: 'rgba(255,204,0,0.75)' },
        Fijos: { border: '#ffb86c', bg: 'rgba(255,184,108,0.75)' }
    };

    // ── 2. ESTADO ────────────────────────────────────────────────────────────
    let state = {
        labels: [], nw: [], inv: [], saved: [],
        inc: [], expenses: [], rinv: [],
        drawdown: [], drawdownInv: [],
        // arrays de desglose: cada elemento es {RV, RF, Cash, Fijos} para ese mes
        breakdown: [],
        through: '', vol: 0.02, peak: 0, peakInv: 0
    };
    let charts = {};

    const MONTHLY_RATE_6 = Math.pow(1.06, 1/12) - 1;
    const MONTHLY_RATE_7 = Math.pow(1.07, 1/12) - 1;
    const MONTHLY_RATE_8 = Math.pow(1.08, 1/12) - 1;
    let ui = { rangePreset: 'all', rangeFrom: null, rangeTo: null, rollingWindow: 12 };

    // ── 3. UTILIDADES ────────────────────────────────────────────────────────
    const log = (msg) => {
        const el = document.getElementById('debug-log');
        if (el) el.innerText = '[' + new Date().toLocaleTimeString() + '] ' + msg;
        console.log(msg);
    };
    const pad2 = (n) => String(n).padStart(2, '0');
    const cleanValue = (val) => {
        if (typeof val === 'number') return val;
        const s = String(val ?? '').trim();
        if (!s || s === '-' || s.startsWith('#')) return 0;
        let t = s;
        if (t.includes('.') && t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
        else t = t.replace(',', '.');
        const n = parseFloat(t.replace(/[^-0-9.]/g, ''));
        return isNaN(n) ? 0 : n;
    };
    const formatEuro = new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
    const normalRand = () => {
        let u = 0, v = 0;
        while (!u) u = Math.random();
        while (!v) v = Math.random();
        return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
    };

    // ── 4. DETECCIÓN DE FILAS POR ETIQUETA ──────────────────────────────────
    function findRow(aoa, labels, occurrence) {
        occurrence = occurrence || 1;
        const targets = labels.map(l => l.toLowerCase().trim());
        let count = 0;
        for (let r = 0; r < aoa.length; r++) {
            const cell = String(aoa[r] && aoa[r][0] != null ? aoa[r][0] : '').toLowerCase().trim();
            if (targets.indexOf(cell) !== -1) { count++; if (count === occurrence) return r; }
        }
        return -1;
    }
    function rowScore(aoa, r) {
        let s = 0;
        for (let c = 1; c <= 12; c++) { if (cleanValue(aoa[r] && aoa[r][c]) !== 0) s++; }
        return s;
    }
    function findNWRow(aoa, yearNum) {
        for (let r = 0; r < aoa.length; r++) {
            if (aoa[r] && (aoa[r][0] == yearNum)) return r;
        }
        for (let r = 0; r < Math.min(15, aoa.length); r++) {
            if (String(aoa[r] && aoa[r][0] ? aoa[r][0] : '').toLowerCase().trim() === 'patrimonio') return r;
        }
        let best = 10, bestS = 0;
        for (let r = 8; r <= 14; r++) { const s = rowScore(aoa, r); if (s > bestS) { bestS = s; best = r; } }
        return best;
    }
    function findInvRow(aoa) {
        const r = findRow(aoa, ['total invertido','invertido total','total inv.','total inv']);
        if (r !== -1) return r;
        let best = 30, bestS = 0;
        for (let rr = 20; rr < Math.min(aoa.length, 55); rr++) {
            const s = rowScore(aoa, rr); if (s > bestS) { bestS = s; best = rr; }
        }
        return best;
    }

    /**
     * Extrae el desglose por clase de activo para una columna (mes) dada.
     * Mapea etiquetas del Excel a RV/RF/Cash/Fijos según catMap.
     * Los valores negativos (ej: "Deuda Ocean" = -500 en el Excel) se suman
     * directamente — no se filtran ni se tratan de forma especial.
     * Para incluir "Deuda Ocean" en Fijos, añádela al catMap de Fijos
     * (ya está por defecto como "deuda ocean").
     */
    function extractBreakdown(aoa, col) {
        // Construir mapa etiqueta→valor para este mes (incluyendo negativos)
        const rowValues = {};
        for (let r = 0; r < aoa.length; r++) {
            const label = String(aoa[r] && aoa[r][0] != null ? aoa[r][0] : '').trim();
            if (!label || label === 'None') continue;
            const val = cleanValue(aoa[r][col]);
            if (val === 0) continue; // ignorar celdas vacías
            rowValues[label.toLowerCase()] = (rowValues[label.toLowerCase()] || 0) + val;
        }

        // Sumar por clase según catMap — sin filtrar negativos
        const result = { RV: 0, RF: 0, Cash: 0, Fijos: 0 };
        for (const [cls, labels] of Object.entries(catMap)) {
            for (const lbl of labels) {
                const v = rowValues[lbl.toLowerCase()];
                if (v !== undefined) result[cls] += v; // suma directa, negativos incluidos
            }
        }

        return result;
    }

    // ── 5. LECTURA EXCEL ────────────────────────────────────────────────────
    async function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        log('Procesando: ' + file.name + '…');
        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array' });
            state = { labels: [], nw: [], inv: [], saved: [], inc: [], expenses: [], rinv: [],
                      drawdown: [], drawdownInv: [], breakdown: [],
                      through: '', vol: 0.02, peak: 0, peakInv: 0 };
            let cumulativeSave = 0;

            const now = new Date();
            const lastCompleted = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            const lastCompletedYM = lastCompleted.getFullYear() * 100 + (lastCompleted.getMonth() + 1);
            state.through = lastCompleted.getFullYear() + '-' + pad2(lastCompleted.getMonth() + 1);

            const yearSheets = workbook.SheetNames
                .filter(n => /^\d{4}$/.test(n.trim()))
                .sort((a, b) => parseInt(a) - parseInt(b));
            if (!yearSheets.length) throw new Error('No se encontraron pestañas con formato año.');

            yearSheets.forEach(yearStr => {
                const yearNum = parseInt(yearStr);
                const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[yearStr], { header: 1 });

                const ROW_INC  = findRow(aoa, ['ingresos'], 1);
                const ROW_OUT  = findRow(aoa, ['gastos'], 1);
                const ROW_SAVE = findRow(aoa, ['ahorro'], 1);
                const ROW_RINV = findRow(aoa, ['rentab inv','rentabilidad inv','rentab. inv','rend inv','rentabilidad inversión','rinv'], 1);
                const ROW_NW   = findNWRow(aoa, yearNum);
                const ROW_INV  = findInvRow(aoa);

                if (ROW_INC === -1 || ROW_OUT === -1 || ROW_SAVE === -1) { log('Aviso: hoja ' + yearStr + ' omitida.'); return; }

                for (let col = 1; col <= 12; col++) {
                    if (yearNum * 100 + col > lastCompletedYM) continue;
                    const valNW = ROW_NW !== -1 ? cleanValue(aoa[ROW_NW] && aoa[ROW_NW][col]) : 0;
                    if (valNW === 0 && state.nw.length === 0) continue;

                    state.labels.push(yearStr + '-' + pad2(col));
                    state.nw.push(valNW);
                    const valInv = ROW_INV !== -1 ? cleanValue(aoa[ROW_INV] && aoa[ROW_INV][col]) : 0;
                    state.inv.push(valInv);
                    state.inc.push(cleanValue(aoa[ROW_INC] && aoa[ROW_INC][col]));
                    state.expenses.push(cleanValue(aoa[ROW_OUT] && aoa[ROW_OUT][col]));
                    state.rinv.push(ROW_RINV !== -1 ? cleanValue(aoa[ROW_RINV] && aoa[ROW_RINV][col]) : 0);
                    const monthlySave = cleanValue(aoa[ROW_SAVE] && aoa[ROW_SAVE][col]);
                    cumulativeSave += monthlySave;
                    state.saved.push(cumulativeSave);
                    if (valNW > state.peak) state.peak = valNW;
                    state.drawdown.push(state.peak > 0 ? ((valNW - state.peak) / state.peak) * 100 : 0);
                    if (valInv > state.peakInv) state.peakInv = valInv;
                    state.drawdownInv.push(state.peakInv > 0 ? ((valInv - state.peakInv) / state.peakInv) * 100 : 0);

                    // Desglose por clase de activo
                    state.breakdown.push(extractBreakdown(aoa, col));
                }
            });

            if (!state.nw.length) throw new Error('No se encontraron datos.');
            calcVol();
            saveLS();
            updateDashboardUI();
            log('Sincronización completa — ' + state.labels.length + ' meses.');
        } catch (err) {
            log('Error: ' + err.message);
            alert('Error al procesar el Excel:\n' + err.message);
        }
    }

    // ── 6. PERSISTENCIA ─────────────────────────────────────────────────────
    const LS_KEY = 'visorPatrimonial_v37';
    function saveLS() { try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch(e) {} }
    function loadLS() {
        try {
            const raw = localStorage.getItem(LS_KEY);
            if (!raw) return false;
            const s = JSON.parse(raw);
            if (!s || !s.labels || !s.labels.length) return false;
            state = s;
            if (!state.breakdown) state.breakdown = state.labels.map(() => ({ RV:0, RF:0, Cash:0, Fijos:0 }));
            return true;
        } catch(e) { return false; }
    }

    // ── 7. VOLATILIDAD ──────────────────────────────────────────────────────
    function calcVol() {
        const returns = [];
        for (let i = 1; i < state.nw.length; i++) {
            const prev = state.inv[i-1];
            if (prev > 0) { const r = state.rinv[i] / prev; if (isFinite(r)) returns.push(r); }
        }
        if (returns.length) {
            const mean = returns.reduce((a,b) => a+b, 0) / returns.length;
            const variance = returns.reduce((a, x) => a + (x-mean)**2, 0) / returns.length;
            state.vol = Math.sqrt(variance);
        } else { state.vol = 0.02; }
    }

    // ── 8. UI ────────────────────────────────────────────────────────────────
    function updateDashboardUI() {
        document.getElementById('status').style.display = 'none';
        document.getElementById('tabContainer').style.display = 'flex';
        document.getElementById('dashboard').style.display = 'block';
        const throughEl = document.getElementById('kThrough');
        if (throughEl) throughEl.textContent = state.through || '—';

        const lastNW = state.nw[state.nw.length - 1];
        const windowMonths = parseInt(document.getElementById('coverageWindow').value || '12', 10);
        const last12mExp = state.expenses.slice(-12).reduce((a,b) => a+b, 0) || 1;
        const last12mInc = state.inc.slice(-12).reduce((a,b) => a+b, 0) || 1;
        const expWin = state.expenses.slice(-windowMonths);
        const rinvWin = state.rinv.slice(-windowMonths);
        const expWindowSum = expWin.reduce((a,b) => a+b, 0) || 1;
        const rinvWindowSum = rinvWin.reduce((a,b) => a+b, 0);

        document.getElementById('kTotal').innerText = formatEuro.format(lastNW);
        document.getElementById('kAvgSave').innerText = formatEuro.format(
            (state.saved[state.saved.length-1] - (state.saved[state.saved.length-13] || 0)) / 12);
        document.getElementById('kSaveRate').innerText = (((last12mInc - last12mExp) / last12mInc) * 100).toFixed(1) + '%';
        document.getElementById('kFireMult').innerText = (lastNW / last12mExp).toFixed(1) + 'x';

        const monthlyNeeds = expWindowSum / Math.max(windowMonths, 1);
        const avgMonthlyRinv = rinvWindowSum / Math.max(windowMonths, 1);
        const passiveCoverage = (avgMonthlyRinv / monthlyNeeds) * 100;
        document.getElementById('kPassivePct').innerText = Math.max(0, passiveCoverage).toFixed(1) + '%';
        document.getElementById('gaugeFill').style.width = Math.min(Math.max(0, passiveCoverage), 100) + '%';
        document.getElementById('kPassiveStatus').innerText = passiveCoverage >= 100
            ? '¡LIBERTAD FINANCIERA ALCANZADA!'
            : 'Faltan ' + formatEuro.format(monthlyNeeds - avgMonthlyRinv) + ' mensuales de rentas.';

        const btnClear = document.getElementById('btnClearCache');
        if (btnClear) btnClear.style.display = 'inline-block';

        renderCharts();
        syncRangeControls();
        renderInvestmentExtras();
        renderBreakdownTab();
        renderDataTable();
    }

    // ── 9. RANGO ────────────────────────────────────────────────────────────
    function syncRangeControls() {
        const presetEl = document.getElementById('rangePreset');
        const fromEl   = document.getElementById('rangeFrom');
        const toEl     = document.getElementById('rangeTo');
        const customWrap = document.getElementById('customRange');
        if (!presetEl || !fromEl || !toEl) return;
        ui.rangePreset = presetEl.value || ui.rangePreset;
        if (customWrap) customWrap.style.display = (ui.rangePreset === 'custom') ? 'flex' : 'none';
        if (fromEl.options.length !== state.labels.length) {
            fromEl.innerHTML = ''; toEl.innerHTML = '';
            state.labels.forEach(lbl => {
                const o1 = document.createElement('option'); o1.value = lbl; o1.textContent = lbl; fromEl.appendChild(o1);
                const o2 = document.createElement('option'); o2.value = lbl; o2.textContent = lbl; toEl.appendChild(o2);
            });
        }
        const first = state.labels[0], last = state.labels[state.labels.length-1];
        if (!ui.rangeFrom || state.labels.indexOf(ui.rangeFrom) === -1) ui.rangeFrom = first;
        if (!ui.rangeTo   || state.labels.indexOf(ui.rangeTo)   === -1) ui.rangeTo   = last;
        fromEl.value = ui.rangeFrom; toEl.value = ui.rangeTo;
    }
    function getRangeIndices() {
        const n = state.labels.length, end = n - 1;
        const preset = (document.getElementById('rangePreset') || {}).value || ui.rangePreset || 'all';
        ui.rangePreset = preset;
        if (preset === 'all') return { start: 0, end };
        if (preset === 'custom') {
            const fromVal = (document.getElementById('rangeFrom') || {}).value || ui.rangeFrom || state.labels[0];
            const toVal   = (document.getElementById('rangeTo')   || {}).value || ui.rangeTo   || state.labels[end];
            ui.rangeFrom = fromVal; ui.rangeTo = toVal;
            let i0 = state.labels.indexOf(fromVal), i1 = state.labels.indexOf(toVal);
            if (i0 < 0) i0 = 0; if (i1 < 0) i1 = end;
            if (i0 > i1) { const t = i0; i0 = i1; i1 = t; }
            return { start: i0, end: i1 };
        }
        const months = parseInt(preset, 10);
        if (!isFinite(months) || months <= 0) return { start: 0, end };
        return { start: Math.max(0, end - months + 1), end };
    }
    const sliceArr = (arr, start, end) => arr.slice(start, end + 1);

    // ── 10. GRÁFICOS BASE ───────────────────────────────────────────────────
    function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

    function baseOpts(yTickFn, extraTooltipCb) {
        return {
            responsive: true, maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { labels: { color: '#c9d1d9', usePointStyle: true, pointStyleWidth: 10 } },
                tooltip: {
                    backgroundColor: '#161b22', borderColor: '#30363d', borderWidth: 1,
                    titleColor: '#f0f6fc', bodyColor: '#8b949e',
                    callbacks: extraTooltipCb || {}
                }
            },
            scales: {
                x: { ticks: { color: '#8b949e', maxTicksLimit: 24, maxRotation: 0 }, grid: { color: '#21262d' } },
                y: { ticks: { color: '#8b949e', callback: yTickFn || (v => v) }, grid: { color: '#21262d' } }
            }
        };
    }

    function renderCharts() {
        // Días de libertad
        destroyChart('freedom');
        const days = state.expenses.map((exp, i) => {
            const saved = state.saved[i] - (state.saved[i-1] || 0);
            const cpd = exp / 30.4;
            return cpd > 0 ? parseFloat((saved / cpd).toFixed(1)) : 0;
        });
        charts.freedom = new Chart(document.getElementById('freedomChart'), {
            type: 'bar',
            data: { labels: state.labels, datasets: [{ label: 'Días comprados', data: days, borderRadius: 3,
                backgroundColor: days.map(v => v >= 0 ? 'rgba(88,166,255,0.75)' : 'rgba(218,54,51,0.75)') }] },
            options: baseOpts(null, { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(1) + ' días' })
        });

        // Alpha
        destroyChart('alpha');
        let v6=100, v7=100, v8=100, vMy=100;
        const curves = state.nw.map((_, i) => {
            v6 *= (1+MONTHLY_RATE_6); v7 *= (1+MONTHLY_RATE_7); v8 *= (1+MONTHLY_RATE_8);
            if (i > 0 && state.inv[i-1] > 0) vMy *= (1 + state.rinv[i] / state.inv[i-1]);
            return { v6: +v6.toFixed(2), v7: +v7.toFixed(2), v8: +v8.toFixed(2), vMy: +vMy.toFixed(2) };
        });
        charts.alpha = new Chart(document.getElementById('alphaChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: [
                { label: 'Mi Cartera', data: curves.map(c=>c.vMy), borderColor: '#00ffaa', borderWidth: 3, pointRadius: 0, tension: 0.2 },
                { label: 'Benchmark 8%', data: curves.map(c=>c.v8), borderColor: '#ffb86c', borderDash: [8,6], pointRadius: 0, borderWidth: 1.5 },
                { label: 'Benchmark 7%', data: curves.map(c=>c.v7), borderColor: '#8b949e', borderDash: [5,5], pointRadius: 0, borderWidth: 1.5 },
                { label: 'Benchmark 6%', data: curves.map(c=>c.v6), borderColor: '#484f58', borderDash: [2,2], pointRadius: 0, borderWidth: 1.5 }
            ]},
            options: baseOpts(v => v.toFixed(1), { label: ctx => ctx.dataset.label + ': ' + Number(ctx.parsed.y).toFixed(1) })
        });

        // Creep
        destroyChart('creep');
        charts.creep = new Chart(document.getElementById('creepChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: [
                { label: 'Ingresos', data: state.inc, borderColor: '#238636', backgroundColor: 'rgba(35,134,54,0.1)', fill: true, tension: 0.3, pointRadius: 0 },
                { label: 'Gastos', data: state.expenses, borderColor: '#da3633', backgroundColor: 'rgba(218,54,51,0.08)', fill: true, tension: 0.3, pointRadius: 0 }
            ]},
            options: baseOpts(v => formatEuro.format(v), { label: ctx => ctx.dataset.label + ': ' + formatEuro.format(ctx.parsed.y) })
        });

        // Evolución
        destroyChart('main');
        charts.main = new Chart(document.getElementById('mainChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: [
                { label: 'Patrimonio Neto', data: state.nw, borderColor: '#58a6ff', fill: true, backgroundColor: 'rgba(88,166,255,0.08)', tension: 0.2, pointRadius: 0, borderWidth: 2.5 },
                { label: 'Total Invertido', data: state.inv, borderColor: '#ffb86c', borderDash: [6,4], pointRadius: 0, tension: 0.2, borderWidth: 1.5 },
                { label: 'Ahorro Acumulado', data: state.saved, borderColor: '#d278ff', borderDash: [3,3], pointRadius: 0, tension: 0.2, borderWidth: 1.5 }
            ]},
            options: baseOpts(v => formatEuro.format(v), { label: ctx => ctx.dataset.label + ': ' + formatEuro.format(ctx.parsed.y) })
        });

        // Rendimiento
        destroyChart('yield');
        charts.yield = new Chart(document.getElementById('yieldChart'), {
            type: 'bar',
            data: { labels: state.labels, datasets: [{ label: 'Rendimiento (€)', data: state.rinv, borderRadius: 3,
                backgroundColor: state.rinv.map(v => v >= 0 ? 'rgba(35,134,54,0.8)' : 'rgba(218,54,51,0.8)') }] },
            options: baseOpts(v => formatEuro.format(v), { label: ctx => 'Rendimiento: ' + formatEuro.format(ctx.parsed.y) })
        });

        // Drawdown
        destroyChart('dd');
        charts.dd = new Chart(document.getElementById('drawdownChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: [
                { label: 'Drawdown Patrimonio', data: state.drawdown, borderColor: '#ff5050', fill: true, backgroundColor: 'rgba(218,54,51,0.1)', pointRadius: 0, tension: 0.2 },
                { label: 'Drawdown Invertido', data: state.drawdownInv, borderColor: '#ffb86c', fill: false, pointRadius: 0, tension: 0.2 }
            ]},
            options: baseOpts(v => v.toFixed(1) + '%', { label: ctx => ctx.dataset.label + ': ' + Number(ctx.parsed.y).toFixed(1) + '%' })
        });

        renderMontecarlo();
    }

    function renderInvestmentExtras() {
        if (!state.labels.length) return;
        const { start, end } = getRangeIndices();
        const labels = sliceArr(state.labels, start, end);
        const invShare = state.labels.map((_, i) => {
            const nw = state.nw[i] || 0, inv = state.inv[i] || 0;
            return nw > 0 ? +((inv/nw)*100).toFixed(1) : 0;
        });
        destroyChart('invShare');
        charts.invShare = new Chart(document.getElementById('invShareChart'), {
            type: 'line',
            data: { labels, datasets: [{ label: '% Patrimonio invertido', data: sliceArr(invShare, start, end), borderColor: '#58a6ff', pointRadius: 0, tension: 0.25 }] },
            options: baseOpts(v => v + '%', { label: ctx => '% invertido: ' + Number(ctx.parsed.y).toFixed(1) + '%' })
        });
        const w = parseInt((document.getElementById('rollingWindow') || {}).value || '12', 10) || 12;
        ui.rollingWindow = w;
        const n = state.labels.length, roll = new Array(n).fill(null);
        for (let i = 0; i < n; i++) {
            if (i < w-1) continue;
            let sumR = 0, sumInv = 0;
            for (let j = i-w+1; j <= i; j++) { sumR += (state.rinv[j]||0); sumInv += (state.inv[j]||0); }
            roll[i] = sumInv/w > 0 ? +((sumR/(sumInv/w))*100).toFixed(2) : 0;
        }
        destroyChart('rolling');
        charts.rolling = new Chart(document.getElementById('rollingReturnChart'), {
            type: 'line',
            data: { labels, datasets: [{ label: 'Rentabilidad móvil (' + w + 'm)', data: sliceArr(roll, start, end), borderColor: '#00ffaa', pointRadius: 0, tension: 0.25, spanGaps: false }] },
            options: baseOpts(v => v + '%', { label: ctx => { const y = ctx.parsed.y; return y == null ? ctx.dataset.label + ': —' : ctx.dataset.label + ': ' + Number(y).toFixed(2) + '%'; } })
        });
    }

    // ── 11. PESTAÑA DESGLOSE ────────────────────────────────────────────────
    function renderBreakdownTab() {
        if (!state.breakdown || !state.breakdown.length) return;

        // ── A. Área apilada histórica ──
        destroyChart('stackArea');
        const classes = ['RV', 'RF', 'Cash', 'Fijos'];
        const stackDatasets = classes.map(cls => ({
            label: cls,
            data: state.breakdown.map(b => b[cls] || 0),
            borderColor: CAT_COLORS[cls].border,
            backgroundColor: CAT_COLORS[cls].bg,
            fill: true, pointRadius: 0, tension: 0.2, borderWidth: 1.5
        }));
        charts.stackArea = new Chart(document.getElementById('stackAreaChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: stackDatasets },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { labels: { color: '#c9d1d9', usePointStyle: true } },
                    tooltip: {
                        backgroundColor: '#161b22', borderColor: '#30363d', borderWidth: 1,
                        titleColor: '#f0f6fc', bodyColor: '#8b949e',
                        callbacks: {
                            label: ctx => ctx.dataset.label + ': ' + formatEuro.format(ctx.parsed.y),
                            footer: items => {
                                const total = items.reduce((s, i) => s + i.parsed.y, 0);
                                return 'Total: ' + formatEuro.format(total);
                            }
                        }
                    }
                },
                scales: {
                    x: { stacked: true, ticks: { color: '#8b949e', maxTicksLimit: 24, maxRotation: 0 }, grid: { color: '#21262d' } },
                    y: { stacked: true, ticks: { color: '#8b949e', callback: v => formatEuro.format(v) }, grid: { color: '#21262d' } }
                }
            }
        });

        // ── B. Donut composición actual ──
        const last = state.breakdown[state.breakdown.length - 1];
        const donutValues = classes.map(cls => Math.round(last[cls] || 0));
        const donutTotal = donutValues.reduce((a,b) => a+b, 0);
        destroyChart('donut');
        charts.donut = new Chart(document.getElementById('donutChart'), {
            type: 'doughnut',
            data: {
                labels: classes,
                datasets: [{
                    data: donutValues,
                    backgroundColor: classes.map(cls => CAT_COLORS[cls].bg),
                    borderColor: classes.map(cls => CAT_COLORS[cls].border),
                    borderWidth: 2, hoverOffset: 10
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                cutout: '58%',
                layout: { padding: { bottom: 8 } },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            color: '#c9d1d9', usePointStyle: true, pointStyle: 'circle',
                            padding: 20, font: { size: 13 },
                            generateLabels: chart => {
                                return chart.data.labels.map((lbl, i) => ({
                                    text: lbl + '   ' + formatEuro.format(donutValues[i]) + '  (' + (donutTotal > 0 ? ((donutValues[i]/donutTotal)*100).toFixed(1) : 0) + '%)',
                                    fillStyle: CAT_COLORS[lbl].bg, strokeStyle: CAT_COLORS[lbl].border,
                                    fontColor: '#c9d1d9', lineWidth: 1, hidden: false, index: i
                                }));
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: '#161b22', borderColor: '#30363d', borderWidth: 1,
                        titleColor: '#f0f6fc', bodyColor: '#8b949e',
                        callbacks: {
                            label: ctx => {
                                const v = ctx.parsed;
                                const pct = donutTotal > 0 ? ((v/donutTotal)*100).toFixed(1) : 0;
                                return '  ' + formatEuro.format(v) + ' (' + pct + '%)';
                            }
                        }
                    }
                }
            }
        });

        // ── C. % por clase histórico (líneas) ──
        destroyChart('classPct');
        const pctDatasets = classes.map(cls => {
            const data = state.breakdown.map(b => {
                const total = classes.reduce((s, c) => s + (b[c]||0), 0);
                return total > 0 ? +((b[cls]||0)/total*100).toFixed(1) : 0;
            });
            return { label: '% ' + cls, data, borderColor: CAT_COLORS[cls].border, pointRadius: 0, tension: 0.25, borderWidth: 2, fill: false };
        });
        charts.classPct = new Chart(document.getElementById('classPctChart'), {
            type: 'line',
            data: { labels: state.labels, datasets: pctDatasets },
            options: baseOpts(v => v + '%', { label: ctx => ctx.dataset.label + ': ' + Number(ctx.parsed.y).toFixed(1) + '%' })
        });

        // ── D. KPIs de desglose ──
        renderBreakdownKPIs(last, donutTotal);
    }

    function renderBreakdownKPIs(last, total) {
        const el = document.getElementById('breakdownKPIs');
        if (!el) return;
        const classes = ['RV', 'RF', 'Cash', 'Fijos'];
        el.innerHTML = classes.map(cls => {
            const val = last[cls] || 0;
            const pct = total > 0 ? ((val/total)*100).toFixed(1) : '0.0';
            return '<div class="kpi-card" style="border-left:3px solid ' + CAT_COLORS[cls].border + '">' +
                '<span class="k-label">' + cls + '</span>' +
                '<span class="k-val" style="color:' + CAT_COLORS[cls].border + '">' + formatEuro.format(val) + '</span>' +
                '<span style="font-size:0.85rem;color:#8b949e;margin-top:4px;display:block">' + pct + '% del patrimonio</span>' +
                '</div>';
        }).join('');
    }

    // ── 12. MODAL CONFIGURACIÓN CATMAP ──────────────────────────────────────
    function openCatMapModal() {
        const modal = document.getElementById('catMapModal');
        if (!modal) return;
        const classes = ['RV', 'RF', 'Cash', 'Fijos'];
        document.getElementById('catMapForm').innerHTML = classes.map(cls =>
            '<div class="catmap-row">' +
            '<label class="catmap-label" style="color:' + CAT_COLORS[cls].border + '">' + cls + '</label>' +
            '<input class="catmap-input" data-cls="' + cls + '" value="' + (catMap[cls] || []).join(', ') + '" ' +
            'placeholder="etiqueta1, etiqueta2, …" />' +
            '</div>'
        ).join('');
        modal.classList.add('open');
    }
    function saveCatMapFromModal() {
        document.querySelectorAll('.catmap-input').forEach(inp => {
            const cls = inp.dataset.cls;
            catMap[cls] = inp.value.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
        });
        saveCatMap(catMap);
        document.getElementById('catMapModal').classList.remove('open');
        // Re-renderizar si hay datos
        if (state.breakdown && state.breakdown.length) renderBreakdownTab();
    }

    // ── 13. MONTECARLO ───────────────────────────────────────────────────────
    function renderMontecarlo() {
        const ITER = 2000, PERIODS = 240;
        const startValue = state.nw[state.nw.length-1];
        if (!startValue) return;
        const annualReturn = (parseFloat((document.getElementById('fireReturn') || {}).value) || 6) / 100;
        const monthlyReturn = Math.pow(1+annualReturn, 1/12) - 1;
        const n = state.saved.length;
        const avgMonthlySave = (state.saved[n-1] - (state.saved[Math.max(0,n-13)] || 0)) / 12;
        const btn = document.getElementById('btnRecalcFire');
        if (btn) { btn.disabled = true; btn.textContent = '⏳ Calculando…'; }
        setTimeout(() => {
            const results = [];
            for (let i = 0; i < ITER; i++) {
                let cur = startValue, path = [cur];
                for (let m = 1; m <= PERIODS; m++) {
                    let sh = normalRand() * state.vol;
                    const cap = 4 * state.vol;
                    sh = Math.max(-cap, Math.min(cap, sh));
                    cur = cur * (1 + monthlyReturn + sh) + avgMonthlySave;
                    path.push(cur);
                }
                results.push(path);
            }
            const pct = (mIdx, p) => results.map(path => path[mIdx]).sort((a,b) => a-b)[Math.floor(ITER*p)];
            const lbs = Array.from({length:PERIODS+1}, (_,i) => i%12===0 ? 'Año '+(i/12) : '');
            const p5  = lbs.map((_,i) => pct(i,0.05));
            const p25 = lbs.map((_,i) => pct(i,0.25));
            const p50 = lbs.map((_,i) => pct(i,0.50));
            const p75 = lbs.map((_,i) => pct(i,0.75));
            const p95 = lbs.map((_,i) => pct(i,0.95));
            destroyChart('fire');
            charts.fire = new Chart(document.getElementById('fireChart'), {
                type: 'line',
                data: { labels: lbs, datasets: [
                    { label:'P95 (muy optimista)', data:p95, borderColor:'rgba(0,255,170,0.20)', backgroundColor:'rgba(0,255,170,0.05)', fill:true, pointRadius:0, borderWidth:1 },
                    { label:'P75', data:p75, borderColor:'rgba(0,255,170,0.45)', backgroundColor:'rgba(0,255,170,0.07)', fill:'-1', pointRadius:0, borderWidth:1.5 },
                    { label:'P50 (mediana)', data:p50, borderColor:'#00ffaa', fill:false, pointRadius:0, borderWidth:3 },
                    { label:'P25', data:p25, borderColor:'rgba(255,184,108,0.45)', backgroundColor:'rgba(255,80,80,0.05)', fill:'-1', pointRadius:0, borderWidth:1.5 },
                    { label:'P5 (muy pesimista)', data:p5, borderColor:'rgba(255,80,80,0.25)', backgroundColor:'rgba(255,80,80,0.03)', fill:'-1', pointRadius:0, borderWidth:1 }
                ]},
                options: { responsive:true, maintainAspectRatio:false, interaction:{mode:'index',intersect:false},
                    plugins:{ legend:{labels:{color:'#c9d1d9',usePointStyle:true}},
                        tooltip:{backgroundColor:'#161b22',borderColor:'#30363d',borderWidth:1,titleColor:'#f0f6fc',bodyColor:'#8b949e',
                            callbacks:{label:ctx=>ctx.dataset.label+': '+formatEuro.format(ctx.parsed.y)}} },
                    scales:{ x:{ticks:{color:'#8b949e',maxTicksLimit:21,maxRotation:0},grid:{color:'#21262d'}},
                             y:{ticks:{color:'#8b949e',callback:v=>formatEuro.format(v)},grid:{color:'#21262d'}} } }
            });
            const statsEl = document.getElementById('fireStats');
            if (statsEl) statsEl.innerHTML =
                '<div class="fire-stat"><span class="fs-label">Mediana 10 años</span><span class="fs-val">'+formatEuro.format(p50[120])+'</span></div>'+
                '<div class="fire-stat"><span class="fs-label">Mediana 20 años</span><span class="fs-val">'+formatEuro.format(p50[240])+'</span></div>'+
                '<div class="fire-stat"><span class="fs-label">Rango P5–P95 a 20a</span><span class="fs-val">'+formatEuro.format(p5[240])+' – '+formatEuro.format(p95[240])+'</span></div>'+
                '<div class="fire-stat"><span class="fs-label">Volatilidad mensual</span><span class="fs-val">'+(state.vol*100).toFixed(2)+'%</span></div>';
            if (btn) { btn.disabled=false; btn.textContent='🔄 Recalcular'; }
        }, 10);
    }

    // ── 14. TABLA ────────────────────────────────────────────────────────────
    function renderDataTable() {
        const tbody = document.querySelector('#dataTable tbody');
        if (!tbody) return;
        const n = state.labels.length, take = Math.min(36,n), start = Math.max(0,n-take);
        tbody.innerHTML = '';
        for (let i = n-1; i >= start; i--) {
            const tr = document.createElement('tr');
            const sm = state.saved[i] - (state.saved[i-1]||0);
            [state.labels[i], formatEuro.format(state.inc[i]||0), formatEuro.format(state.expenses[i]||0),
             formatEuro.format(sm||0), formatEuro.format(state.rinv[i]||0),
             formatEuro.format(state.inv[i]||0), formatEuro.format(state.nw[i]||0),
             (state.drawdown[i]||0).toFixed(1)+'%'].forEach((t, ci) => {
                const td = document.createElement('td'); td.textContent = t;
                if (ci===3) td.style.color = sm>=0 ? '#238636':'#da3633';
                if (ci===4) td.style.color = (state.rinv[i]||0)>=0 ? '#238636':'#da3633';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        }
    }

    // ── 15. TABS ─────────────────────────────────────────────────────────────
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn,.tab-content').forEach(el => el.classList.remove('active'));
            btn.classList.add('active');
            const content = document.getElementById(btn.dataset.tab);
            if (content) content.classList.add('active');
        });
    });

    // ── 16. EVENTOS ─────────────────────────────────────────────────────────
    document.getElementById('file').addEventListener('change', handleFileUpload);
    document.getElementById('fireReturn').addEventListener('input', () => { if (state.nw.length) renderMontecarlo(); });
    document.getElementById('btnRecalcFire') && document.getElementById('btnRecalcFire').addEventListener('click', () => { if (state.nw.length) renderMontecarlo(); });
    document.getElementById('coverageWindow') && document.getElementById('coverageWindow').addEventListener('change', () => { if (state.nw.length) updateDashboardUI(); });
    document.getElementById('btnConfigCatMap') && document.getElementById('btnConfigCatMap').addEventListener('click', openCatMapModal);
    document.getElementById('btnSaveCatMap') && document.getElementById('btnSaveCatMap').addEventListener('click', saveCatMapFromModal);
    document.getElementById('btnCancelCatMap') && document.getElementById('btnCancelCatMap').addEventListener('click', () => {
        document.getElementById('catMapModal').classList.remove('open');
    });

    const presetEl = document.getElementById('rangePreset');
    const fromEl   = document.getElementById('rangeFrom');
    const toEl     = document.getElementById('rangeTo');
    const rollEl   = document.getElementById('rollingWindow');
    const customWrap = document.getElementById('customRange');
    const onRangeChange = () => {
        if (!state.labels.length) return;
        ui.rangePreset = (presetEl||{}).value || ui.rangePreset;
        if (customWrap) customWrap.style.display = (ui.rangePreset==='custom')?'flex':'none';
        ui.rangeFrom = (fromEl||{}).value || ui.rangeFrom;
        ui.rangeTo   = (toEl  ||{}).value || ui.rangeTo;
        ui.rollingWindow = parseInt(((rollEl||{}).value||ui.rollingWindow||'12'),10)||12;
        renderInvestmentExtras();
    };
    presetEl && presetEl.addEventListener('change', onRangeChange);
    fromEl   && fromEl.addEventListener('change', onRangeChange);
    toEl     && toEl.addEventListener('change', onRangeChange);
    rollEl   && rollEl.addEventListener('change', onRangeChange);

    const modal = document.getElementById('dataModal');
    document.getElementById('btnData') && document.getElementById('btnData').addEventListener('click', () => {
        if (!state.labels.length) { alert('Primero carga el Excel.'); return; }
        renderDataTable(); modal && modal.classList.add('open'); modal && modal.setAttribute('aria-hidden','false');
    });
    document.getElementById('btnCloseModal') && document.getElementById('btnCloseModal').addEventListener('click', () => {
        modal && modal.classList.remove('open');
    });
    modal && modal.addEventListener('click', ev => { if (ev.target===modal) modal.classList.remove('open'); });
    document.addEventListener('keydown', ev => {
        if (ev.key==='Escape') {
            modal && modal.classList.remove('open');
            document.getElementById('catMapModal') && document.getElementById('catMapModal').classList.remove('open');
        }
    });
    document.getElementById('btnClearCache') && document.getElementById('btnClearCache').addEventListener('click', () => {
        if (confirm('¿Borrar datos guardados? Tendrás que recargar el Excel.')) { localStorage.removeItem(LS_KEY); location.reload(); }
    });
    document.getElementById('btnExportCsv') && document.getElementById('btnExportCsv').addEventListener('click', () => {
        if (!state.labels.length) return;
        const rows = [['Mes','Ingresos','Gastos','AhorroMes','RinvEUR','Invertido','Patrimonio','DrawdownPct']];
        for (let i = 0; i < state.labels.length; i++) {
            const sm = state.saved[i] - (state.saved[i-1]||0);
            rows.push([state.labels[i],state.inc[i]||0,state.expenses[i]||0,sm||0,state.rinv[i]||0,state.inv[i]||0,state.nw[i]||0,(state.drawdown[i]||0).toFixed(2)]);
        }
        const csv = rows.map(r => r.join(';')).join('\n');
        const blob = new Blob([csv],{type:'text/csv;charset=utf-8'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href=url; a.download='visor_datos.csv';
        document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    });

    // ── 17. ARRANQUE ────────────────────────────────────────────────────────
    if (loadLS()) {
        calcVol();
        updateDashboardUI();
        log('Datos restaurados desde caché (' + state.labels.length + ' meses). Hasta: ' + state.through);
    }

})();
