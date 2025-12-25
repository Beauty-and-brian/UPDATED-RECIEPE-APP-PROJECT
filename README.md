<!doctype html>
<html lang="en" class="light">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>MTN FinTech: Predictive Capital Engine</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@100..900&display=swap" rel="stylesheet">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.17.4/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@tensorflow/tfjs@3.14.0/dist/tf.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@3.7.1/dist/chart.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@1.4.0/dist/chartjs-plugin-annotation.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>

    <script>
        // Customizing Tailwind to include MTN colors
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        // MTN Official Colors
                        'mtn-yellow': '#ffcb05', 
                        'mtn-black': '#000000',
                        'mtn-gray': '#303030', // Used for dark card backgrounds
                        // General UI colors (light/dark mode)
                        'ui-bg-light': '#f7f7f9',
                        'ui-card-light': '#ffffff',
                        'ui-text-light': '#1f2937',
                        'ui-bg-dark': '#111827', // Dark background for better contrast
                        'ui-card-dark': '#1f2937',
                        'ui-text-dark': '#f3f4f6',
                    }
                }
            }
        }
    </script>

    <style>
        /* ---------------------------------- */
        /* MTN Corporate Theme & Modern Look  */
        /* ---------------------------------- */
        :root {
            --mtn-yellow: #ffcb05;
            --mtn-black: #000000;
            --mtn-gray: #303030;
            --bg-light: #f7f7f9;
            --text-light: #1f2937;
            --card-bg-light: #ffffff;
            --muted-light: #6b7280;
            --bg-dark: #111827; /* Darker BG for better theme */
            --text-dark: #f3f4f6;
            --card-bg-dark: #1f2937;
            --muted-dark: #9ca3af;
        }

        body {
            background-color: var(--bg-light);
            color: var(--text-light);
            font-family: 'Inter', sans-serif;
            transition: background-color 0.3s, color 0.3s;
        }
        .card {
            background-color: var(--card-bg-light);
            border-radius: 12px;
            box-shadow: 0 6px 30px rgba(11,15,30,0.12);
            transition: background-color 0.3s, box-shadow 0.3s;
        }
        .muted { color: var(--muted-light); }
        .chart-area { height: 380px; }


        /* Make Quick Settings card stretch nicely */
        .card.flex {
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            height: 100%;
        }

        .card input[type="number"] {
            transition: all 0.2s;
        }
        /* MTN-style focus ring for inputs */
        .card input[type="number"]:focus, 
        .card input[type="text"]:focus {
            border-color: var(--mtn-yellow);
            box-shadow: 0 0 5px rgba(255, 203, 5, 0.7);
            outline: none;
        }

        /* General styling for AI Insight Badges (MTN-themed border/background) */
        .forecast-badge {
            padding: 12px;
            border-left: 4px solid;
            border-radius: 6px;
            font-weight: 500;
            line-height: 1.4;
        }
        .forecast-badge.bg-yellow-100 { 
            background-color: #fffbeb; 
            border-color: var(--mtn-yellow); /* Improvement/Stable - use MTN yellow */
        } 
        .forecast-badge.bg-red-100 { background-color: #fee2e2; border-color: #ef4444; } /* Deterioration */
        .forecast-badge.bg-blue-100 { background-color: #eff6ff; border-color: #3b82f6; } /* Stable */

        /* Dark Mode */
        .dark body {
            background-color: var(--bg-dark);
            color: var(--text-dark);
        }
        .dark .card {
            background-color: var(--card-bg-dark);
            box-shadow: 0 6px 30px rgba(0,0,0,0.3);
        }
        .dark .muted { color: var(--muted-dark); }
        .dark .border { border-color: #374151 !important; }
        /* MTN-themed badges in Dark Mode */
        .dark .forecast-badge.bg-yellow-100 { 
            background-color: #332a00; 
            border-color: var(--mtn-yellow); 
            color: #ffeb99; 
        }
        .dark .forecast-badge.bg-red-100 { background-color: #330000; border-color: #f87171; color: #fecaca; }
        .dark .forecast-badge.bg-blue-100 { background-color: #0d1a33; border-color: #60a5fa; color: #bfdbfe; }
        
        /* Dark Mode Input Fix: Ensure form inputs look good and text is visible */
        .dark input[type="number"], .dark input[type="text"] {
            background-color: #1f2937; /* Same as dark card to blend */
            color: var(--text-dark);
            border-color: #4b5563; 
        }

        /* Modern tech text - subtle yellow glow */
        .tech-text {
            font-family: 'Courier New', Courier, monospace;
            text-shadow: 0 0 1px var(--mtn-yellow), 0 0 3px rgba(255, 203, 5, 0.5);
            font-weight: 600;
        }
        .dark .tech-text {
             text-shadow: 0 0 1px var(--mtn-yellow), 0 0 5px rgba(255, 203, 5, 0.8);
        }

        .text-visible { color: var(--text-light); transition: color 0.3s; }
        .dark .text-visible { color: var(--text-dark); }
        
        /* --- PDF Export Mode Styling (for professionalism) --- */
        .pdf-exporting #report-capture-area {
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .pdf-exporting #report-capture-area .tech-text {
            text-shadow: none !important;
            color: #1f2937 !important;
        }
        .pdf-exporting #report-capture-area .card {
            box-shadow: none !important; 
            border: 1px solid #e5e7eb; 
        }

        /* MTN IMPROVEMENT: Header Specific Styling */
        .mtn-header {
            background-color: var(--mtn-black);
            color: var(--mtn-yellow);
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            position: sticky;
            top: 0;
            z-index: 50;
        }
        .mtn-header a:hover { text-decoration: underline; }

        /* MTN IMPROVEMENT: Footer Specific Styling */
        .mtn-footer {
            background-color: var(--mtn-gray);
            color: #ffffff; /* White text on dark gray footer */
            padding: 1rem 0;
            margin-top: 2rem;
        }
        .dark .mtn-footer {
             background-color: var(--mtn-black);
        }
    </style>
</head>
<body class="antialiased light">

<nav class="mtn-header">
    <div class="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
        <div class="flex items-center space-x-4">
            <span class="text-3xl font-black tracking-widest" style="color: var(--mtn-yellow);">MTN</span>
            <span class="text-xl font-light text-white">FinTech Core</span>
        </div>
        <div class="flex items-center gap-4">
            <a href="#" class="text-sm text-white hover:text-mtn-yellow hidden sm:block">TREASURY</a>
            <a href="#" class="text-sm text-white hover:text-mtn-yellow hidden sm:block">WC OPERATIONS REPORT</a>
            <button id="mode-toggle" class="px-3 py-2 rounded bg-mtn-yellow text-mtn-black font-bold text-sm hover:bg-yellow-400">🌓 Theme</button>
        </div>
    </div>
</nav>

<div class="max-w-6xl mx-auto p-6">
    <header class="mb-6 card p-6 bg-mtn-yellow/10 border-l-4 border-mtn-yellow dark:bg-ui-card-dark dark:border-mtn-yellow">
        <h1 class="text-3xl font-extrabold text-mtn-black dark:text-mtn-yellow">Working Capital Predictive Engine </h1>
        <p class="text-base text-gray-700 dark:text-ui-text-dark mt-2">
            Upload your **Trial Balance** Excel (Period/Date, Account, Ending Balance). 
            Account prefixes: **1 = Receivables (Assets)**, **2 = Payables (Liabilities)**.
        </p>
         <div class="mt-4 flex items-center gap-3">
             <input id="excel-file" type="file" accept=".xlsx,.xls" class="hidden" />
             <label for="excel-file" class="cursor-pointer px-4 py-2 rounded bg-mtn-yellow text-mtn-black font-extrabold shadow-lg hover:bg-yellow-400 transition-colors">📂 Upload Trial Balance Data</label>
         </div>
    </header>

    <section class="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
      <div class="card p-6 flex flex-col justify-between h-full">
          <h3 class="font-bold text-lg tech-text mb-4" style="color: var(--mtn-yellow);">Quick Settings ⚙️</h3> 

      <div class="grid grid-cols-2 gap-4 flex-grow">
          <div class="flex flex-col">
          <label class="text-xs muted mb-1 text-visible">Window size (lookback)</label>
          <input id="windowSize" type="number" value="4" min="2" class="border rounded px-3 py-2 text-visible flex-grow"/>
        </div>
        <div class="flex flex-col">
          <label class="text-xs muted mb-1 text-visible">Epochs</label>
          <input id="epochs" type="number" value="40" min="1" class="border rounded px-3 py-2 text-visible flex-grow"/>
        </div>
        <div class="flex flex-col">
          <label class="text-xs muted mb-1 text-visible">Batch size</label>
          <input id="batchSize" type="number" value="8" min="1" class="border rounded px-3 py-2 text-visible flex-grow"/>
        </div>
        <div class="flex flex-col">
          <label class="text-xs muted mb-1 text-visible">Forecast horizon (months)</label>
          <input id="horizon" type="number" value="12" min="1" class="border rounded px-3 py-2 text-visible flex-grow"/>
        </div>
      </div>

      <div class="mt-6 flex gap-3">
          <button id="trainBtn" class="px-4 py-2 rounded bg-mtn-black text-mtn-yellow font-bold hover:bg-mtn-gray flex-1 transition-colors">▶ Train & Forecast</button>
        <button id="resetBtn" class="px-4 py-2 rounded bg-gray-200 text-mtn-black hover:bg-gray-300 dark:bg-gray-700 dark:text-gray-200 dark:hover:bg-gray-600 flex-1 transition-colors">⟲ Reset</button>
      </div>

      <div id="status" class="mt-4 text-sm muted text-visible">No data loaded.</div>
    </div>


        <div id="report-capture-area" class="lg:col-span-2">
            
            <div id="report-header" class="mb-4 p-3 card">
                <h3 class="font-bold text-lg text-visible">MTN Working Capital Forecast Report</h3>
                <p id="timestamp-content" class="text-xs muted mt-1 text-visible"></p>
            </div>
            
            <div class="card p-4 mb-6">
                <h3 class="font-semibold tech-text text-visible">Summary – Projection (End of Period)</h3>
                <div class="mt-3 grid grid-cols-3 gap-3">
                    <div class="p-3 rounded border dark:border-gray-700">
                        <p class="muted text-sm text-visible">Receivables (forecast)</p>
                        <div id="forecast-receivables" class="text-lg font-semibold tech-text text-visible">—</div>
                    </div>
                    <div class="p-3 rounded border dark:border-gray-700">
                        <p class="muted text-sm text-visible">Payables (forecast)</p>
                        <div id="forecast-payables" class="text-lg font-semibold tech-text text-visible">—</div>
                    </div>
                    <div class="p-3 rounded bg-mtn-yellow"> 
                        <p class="text-sm text-mtn-black font-semibold text-visible">Working Capital (R - P)</p>
                        <div id="forecast-wc" class="text-lg font-bold text-mtn-black tech-text text-visible">—</div>
                    </div>
                </div>
                <div class="mt-4 flex gap-3">
                    <button id="export-chart" class="px-3 py-2 rounded bg-mtn-black text-mtn-yellow hover:bg-gray-800 transition-colors">📊 Export Chart PNG</button>
                    <button id="export-pdf" class="px-3 py-2 rounded bg-mtn-gray text-white hover:bg-gray-500 transition-colors">📄 Export Report PDF</button>
                </div>
            </div>

            <section id="ai-insights-section" class="card p-4 mb-6 hidden">
                <h3 class="font-bold tech-text mb-3 p-2 rounded" style="color: var(--mtn-black); background-color: var(--mtn-yellow);">🧠 Predictive AI Insights & Strategy</h3>
                <div id="ai-insights-content" class="text-sm space-y-3 text-visible">
                    </div>
            </section>
        </div>
    </section>
    
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-4" id="chart-section">
        <div class="card p-4 lg:col-span-2">
            <div class="flex items-center justify-between mb-2">
                <h3 class="font-semibold tech-text text-visible">Historical Actuals</h3>
                <span class="muted text-sm text-visible">(Line Chart — Input Data)</span>
            </div>
            <div class="chart-area">
                <canvas id="historicalChart"></canvas>
            </div>
        </div>

        <div class="card p-4">
            <div class="flex items-center justify-between mb-2">
                <h3 class="font-semibold tech-text text-visible">Annual Projection</h3>
                <span class="muted text-sm text-visible">(Sum per year — Bar Chart)</span>
            </div>
            <div class="chart-area">
                <canvas id="annualChart"></canvas>
            </div>
        </div>

        <div class="card p-4 lg:col-span-3">
            <div class="flex items-center justify-between mb-2">
                <h3 class="font-semibold tech-text text-visible">Full Trajectory (<span id="horizonLabel">12</span> months)</h3>
                <span class="muted text-sm text-visible">(Forecast visually distinct)</span>
            </div>
            <div class="chart-area">
                <canvas id="forecastChart"></canvas>
            </div>
        </div>
    </section>
</div>

<footer class="mtn-footer">
    <div class="max-w-6xl mx-auto px-6 flex flex-col sm:flex-row justify-between items-center text-center sm:text-left">
        <p class="text-xs sm:text-sm">
            &copy; 2025 MTN FINANCE/TREASURY/TROPS Predictive Core. All Rights Reserved.
        </p>
        <div class="flex gap-4 mt-2 sm:mt-0 text-xs sm:text-sm">
            <a href="#" class="hover:text-mtn-yellow">Privacy Policy</a>
            <a href="#" class="hover:text-mtn-yellow">Contact Support</a>
        </div>
    </div>
</footer>

<script>
    // Register annotation plugin (kept outside function for immediate availability)
    if (window['chartjs-plugin-annotation'] && Chart && Chart.register) {
        try { Chart.register(window['chartjs-plugin-annotation']); } catch(e) { console.warn(e); }
    }

    // --- GLOBAL STATE AND DOM REFERENCES ---
    let historicalChart = null, forecastChart = null, annualChart = null;
    let rawWide = null; // {dates:[], receivables:[], payables:[]}

    // DOM Elements (Function for cleaner access/initialization)
    function initializeDOMElements() {
        return {
            fileInput: document.getElementById('excel-file'),
            statusEl: document.getElementById('status'),
            trainBtn: document.getElementById('trainBtn'),
            resetBtn: document.getElementById('resetBtn'),
            exportChartBtn: document.getElementById('export-chart'),
            exportPdfBtn: document.getElementById('export-pdf'),
            horizonInput: document.getElementById('horizon'),
            horizonLabel: document.getElementById('horizonLabel'),
            recEl: document.getElementById('forecast-receivables'),
            payEl: document.getElementById('forecast-payables'),
            wcEl: document.getElementById('forecast-wc'),
            windowSizeInput: document.getElementById('windowSize'),
            epochsInput: document.getElementById('epochs'),
            batchSizeInput: document.getElementById('batchSize'),
            modeToggle: document.getElementById('mode-toggle'),
            // REPORT ELEMENTS
            reportCaptureArea: document.getElementById('report-capture-area'), // Target for Summary/Insights
            timestampContent: document.getElementById('timestamp-content'),
            aiInsightsSection: document.getElementById('ai-insights-section'), 
            aiInsightsContent: document.getElementById('ai-insights-content'),
            chartSection: document.getElementById('chart-section'), // The charts container
        };
    }

    // --- HELPER FUNCTIONS ---
    const formatNGN = (v, decimals = 0) => new Intl.NumberFormat('en-NG', {
        style:'currency',
        currency:'NGN',
        maximumFractionDigits:decimals,
        minimumFractionDigits:decimals
    }).format(v);
    const toMonthLabel = d => { const [y,m] = d.split('-'); return `${m}/${y}`; };
    const setStatus = t => initializeDOMElements().statusEl.textContent = t;

    // --- AI INSIGHTS LOGIC (Kept as is) ---
    function generateAIInsights(lastWC, forecastWC, lastRec, forecastRec, lastPay, forecastPay, horizon) {
        const SIGNIFICANT_PERCENT = 10; // 10% threshold for "notable"
        const SIGNIFICANT_AMOUNT = 500000; // NGN 500k threshold

        const change = forecastWC - lastWC;
        const percentChange = lastWC === 0 ? (Math.abs(change) >= SIGNIFICANT_AMOUNT ? SIGNIFICANT_PERCENT : 0) : (change / Math.abs(lastWC)) * 100;

        let insightHTML = '';
        
        const changeStr = formatNGN(Math.abs(change), 2);
        const trendIcon = change > 0 ? '📈' : change < 0 ? '📉' : '⏸️';

        const getRPChange = (last, forecast, name) => {
            const c = forecast - last;
            const p = last === 0 ? (Math.abs(c) > 100000 ? 999.9 : (c / 100000) * 10) : (c / Math.abs(last)) * 100;
            const sign = c > 0 ? 'increase' : 'decrease';
            const signIcon = c > 0 ? '▲' : '▼';
            return `<p class="ml-4">${signIcon} **${Math.abs(p).toFixed(1)}%** ${sign} (${formatNGN(Math.abs(c), 2)}) in **${name}**.</p>`;
        };

        const recInsight = getRPChange(lastRec, forecastRec, 'Receivables (Assets)');
        const payInsight = getRPChange(lastPay, forecastPay, 'Payables (Liabilities)');
        
        let badgeClass = '';

        if (Math.abs(percentChange) >= SIGNIFICANT_PERCENT || Math.abs(change) >= SIGNIFICANT_AMOUNT) {
            if (change > 0) {
                badgeClass = 'bg-yellow-100 border-yellow-500';
                insightHTML = `
                    <p class="font-semibold text-green-600 dark:text-green-400">The forecast indicates a **Significant Improvement** ${trendIcon} in Working Capital by **${changeStr}** (${percentChange.toFixed(1)}%) over the next ${horizon} months. This is a positive indicator of short-term liquidity.</p>
                    <h4 class="font-bold mt-2">Model Rationale:</h4>
                    ${recInsight}
                    ${payInsight}
                    <p>This improvement is likely driven by the **model detecting a strong historical trend of effective cash collection or a relative slowdown in liabilities growth**.</p>
                    <h4 class="font-bold mt-2">Strategic Recommendation:</h4>
                    <div class="forecast-badge ${badgeClass}">
                        **Capitalize on Surplus:** Reinvest the improved working capital into strategic growth initiatives (e.g., network expansion, new product R&D) or leverage the stronger cash position to secure better vendor discounts for early payment (optimizing **Days Payable Outstanding - DPO**).
                    </div>
                `;
            } else {
                badgeClass = 'bg-red-100 border-red-500';
                insightHTML = `
                    <p class="font-semibold text-red-600 dark:text-red-400">The forecast indicates a **Notable Deterioration** ${trendIcon} in Working Capital by **${changeStr}** (${Math.abs(percentChange).toFixed(1)}%) over the next ${horizon} months. This projects **potential future liquidity stress**.</p>
                    <h4 class="font-bold mt-2">Model Rationale:</h4>
                    ${recInsight}
                    ${payInsight}
                    <p>This deterioration is likely due to the **historical pattern showing a significant increase in Payables relative to Receivables**, suggesting slow collection cycles or aggressive expenditure planning.**</p>
                    <h4 class="font-bold mt-2">Strategic Recommendation:</h4>
                    <div class="forecast-badge ${badgeClass}">
                        **Urgent Cash Flow Review:** Focus intensely on **Days Sales Outstanding (DSO)** reduction. Implement stricter credit control, optimize collections processes, and aggressively negotiate longer payment terms with key suppliers (increasing DPO) to reverse the negative trend.
                    </div>
                `;
            }
        } else {
            badgeClass = 'bg-blue-100 border-blue-500';
            insightHTML = `
                <p class="font-semibold text-blue-600 dark:text-blue-400">The Working Capital is forecasted to remain **Relatively Stable** ${trendIcon} with a minor change of **${changeStr}** (${Math.abs(percentChange).toFixed(1)}%) over the next ${horizon} months.</p>
                <h4 class="font-bold mt-2">Model Rationale:</h4>
                ${recInsight}
                ${payInsight}
                <p>The model suggests that the modeled trends for Receivables and Payables are moving in tandem, maintaining the current balance. The system appears stable but lacks significant optimization.</p>
                <h4 class="font-bold mt-2">Strategic Recommendation:</h4>
                <div class="forecast-badge ${badgeClass}">
                    **Maintain Operational Efficiency:** Continue to monitor core metrics like DSO and DPO to ensure targets are met. Explore marginal, low-risk improvements in cash conversion cycles, and fine-tune cash forecasting models to capture short-term volatility.
                </div>
            `;
        }

        return insightHTML;
    }

    // --- EXCEL PARSING FUNCTION (Kept as is) ---
    function parseWorkbook(workbook){
        const PAY_PREFIX = '2';
        const REC_PREFIX = '1';
        const allReceivables = {}, allPayables = {};

        workbook.SheetNames.forEach(sheetName => {
            const ws = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true });
            const KEY_DATE = ['date','period'];
            const KEY_ACC = ['account','natural','code'];
            const KEY_BAL = ['balance','ending','amount','ending balan']; 
            let headerRow = null, di=-1, ai=-1, bi=-1;
            for (let r=0;r<Math.min(6,rows.length);r++){
                const row = rows[r]; if(!Array.isArray(row)) continue;
                const low = row.map(c=>String(c||'').trim().toLowerCase());
                di = low.findIndex(h=>KEY_DATE.some(k=>h.includes(k)));
                ai = low.findIndex(h=>KEY_ACC.some(k=>h.includes(k)));
                bi = low.findIndex(h=>KEY_BAL.some(k=>h.includes(k)));
                if (di!==-1 && ai!==-1 && bi!==-1){ headerRow=r; break; }
            }
            if (headerRow===null) return; 
            const dataRows = rows.slice(headerRow+1);
            dataRows.forEach(row => {
                if (!Array.isArray(row)) return;
                if (row.length<=Math.max(di,ai,bi)) return;
                let rawDate = row[di];
                let dt = null;
                if (rawDate instanceof Date) dt = rawDate;
                else if (typeof rawDate==='number' && XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function'){
                    const c = XLSX.SSF.parse_date_code(rawDate);
                    if (c) dt = new Date(Date.UTC(c.y, c.m-1, c.d));
                } else {
                    const d = new Date(rawDate);
                    if (!isNaN(d)) dt = d;
                }
                if (!dt) return;
                const key = `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}`; 
                const acc = String(row[ai]||'').trim();
                let bal = row[bi];
                if (typeof bal==='string'){ bal = bal.replace(/,/g,'').replace(/\(|\)/g, s=> s==='(' ? '-' : ''); bal = parseFloat(bal); }
                if (isNaN(bal)) return;
                
                if (acc.startsWith(REC_PREFIX)) allReceivables[key] = (allReceivables[key]||0) + bal;
                else if (acc.startsWith(PAY_PREFIX)) allPayables[key] = (allPayables[key]||0) + bal;
            });
        });

        const recArr = Object.keys(allReceivables).map(k=>({date:k,balance:allReceivables[k]})).sort((a,b)=>a.date.localeCompare(b.date));
        const payArr = Object.keys(allPayables).map(k=>({date:k,balance:allPayables[k]})).sort((a,b)=>a.date.localeCompare(b.date));
        return { receivables:recArr, payables:payArr };
    }

    // --- TENSORFLOW FUNCTIONS (Kept as is) ---
    async function trainAndForecastSeries(series, windowSize, epochs, batchSize, horizon){
        const cleanIdx = series.map((v,i)=> ({v,i})).filter(x=>typeof x.v==='number');
        const values = cleanIdx.map(x=>x.v);
        if (values.length < windowSize + 1) throw new Error(`Not enough points to train. Need ${windowSize + 1}, found ${values.length}`);

        const maxV = Math.max(...values), minV = Math.min(...values), range = maxV - minV || 1;
        const norm = values.map(v=> (v - minV)/range);

        const X=[], Y=[];
        for(let i=0;i+windowSize<norm.length;i++){ X.push(norm.slice(i,i+windowSize)); Y.push(norm[i+windowSize]); }
        const Xt = tf.tensor2d(X);
        const Yt = tf.tensor2d(Y, [Y.length,1]);

        const model = tf.sequential();
        model.add(tf.layers.dense({ inputShape:[windowSize], units:32, activation:'relu' }));
        model.add(tf.layers.dense({ units:16, activation:'relu' }));
        model.add(tf.layers.dense({ units:1, activation:'sigmoid' }));
        model.compile({ optimizer: tf.train.adam(0.005), loss:'meanSquaredError' });

        await model.fit(Xt, Yt, { epochs, batchSize, verbose:0 });

        const preds = [];
        let lastWindow = norm.slice(-windowSize).slice();
        for(let h=0;h<horizon;h++){
            const inT = tf.tensor2d([lastWindow]);
            const outT = model.predict(inT);
            const p = (await outT.data())[0];
            const den = p*range + minV;
            preds.push(den);
            lastWindow.shift(); lastWindow.push(p);
            inT.dispose(); outT.dispose();
        }

        Xt.dispose(); Yt.dispose();
        return { model, preds };
    }

    // --- CHARTING FUNCTIONS (COMPLETED & KEPT AS IS) ---
    const CHART_COLORS = {
        RECEIVABLES: '#10b981', // green (actual)
        PAYABLES: '#ef4444', // red (actual)
        WORKING_CAPITAL: '#ffcb05', // MTN Yellow (actual)
        FORECAST_ZONE_FILL: 'rgba(255,203,5,0.08)', // Light MTN Yellow
        CHART_TEXT: '#1f2937' // Will be set dynamically
    };

    function updateChartColors() {
          // Dynamic text color based on current theme
          CHART_COLORS.CHART_TEXT = document.body.classList.contains('dark')
              ? getComputedStyle(document.documentElement).getPropertyValue('--text-dark')
              : getComputedStyle(document.documentElement).getPropertyValue('--text-light');
    }

    function buildHistoricalChart(wide){
        updateChartColors();
        const ctx = document.getElementById('historicalChart').getContext('2d');
        const labels = wide.dates.map(d=>toMonthLabel(d));

        if (historicalChart) historicalChart.destroy();
        historicalChart = new Chart(ctx, {
            type:'line', data:{ labels, datasets:[
                { label:'Receivables (Actual) R', data: wide.receivables, borderColor:CHART_COLORS.RECEIVABLES, backgroundColor:'rgba(16,185,129,0.06)', fill:true, tension:0.45, pointRadius:3, borderWidth:2 },
                { label:'Payables (Actual) P', data: wide.payables, borderColor:CHART_COLORS.PAYABLES, backgroundColor:'rgba(239,68,68,0.06)', fill:true, tension:0.45, pointRadius:3, borderWidth:2 }
            ] }, options:{
                responsive:true, maintainAspectRatio:false,
                plugins:{
                    legend:{ labels:{ color:CHART_COLORS.CHART_TEXT } },
                    tooltip:{ callbacks:{ label: ctx=>{ const val=ctx.parsed.y; if (val==null) return ''; return ctx.dataset.label + ': ' + formatNGN(val, 2); } } }
                }, scales:{
                    x:{ ticks:{ color:CHART_COLORS.CHART_TEXT } },
                    y:{ ticks:{ callback:v=>formatNGN(v, 0), color:CHART_COLORS.CHART_TEXT } }
                }
            }
        });
    }

    function buildForecastChart(wide, forecast){
        updateChartColors();
        const ctx = document.getElementById('forecastChart').getContext('2d');
        if (forecastChart) forecastChart.destroy();

        const labels = wide.dates.map(d=>toMonthLabel(d)).concat(forecast.dates.map(d=>toMonthLabel(d)));
        const histLen = wide.dates.length;

        const datasets = [
            { label:'Receivables (Actual) R', data: wide.receivables, borderColor:CHART_COLORS.RECEIVABLES, backgroundColor:'transparent', tension:0.4, pointRadius:3, borderWidth:2, yAxisID: 'y' },
            { label:'Receivables (Forecast) R_f', data: Array(histLen).fill(null).concat(forecast.recPred), borderColor:CHART_COLORS.RECEIVABLES, backgroundColor:'rgba(16,185,129,0.12)', borderDash:[6,4], fill:false, tension:0.4, pointRadius:3, borderWidth:2, yAxisID: 'y' },
            { label:'Payables (Actual) P', data: wide.payables, borderColor:CHART_COLORS.PAYABLES, backgroundColor:'transparent', tension:0.4, pointRadius:3, borderWidth:2, yAxisID: 'y' },
            { label:'Payables (Forecast) P_f', data: Array(histLen).fill(null).concat(forecast.payPred), borderColor:CHART_COLORS.PAYABLES, backgroundColor:'rgba(239,68,68,0.12)', borderDash:[6,4], fill:false, tension:0.4, pointRadius:3, borderWidth:2, yAxisID: 'y' },
            { label:'Working Capital (Actual) WC', data: wide.dates.map((d,i)=>{ const r=wide.receivables[i], p=wide.payables[i]; return (typeof r==='number'&&typeof p==='number')? r-p : null }), borderColor:CHART_COLORS.WORKING_CAPITAL, backgroundColor:'transparent', tension:0.4, pointRadius:3, borderWidth:3, yAxisID: 'y1' },
            { label:'Working Capital (Forecast) WC_f', data: Array(histLen).fill(null).concat(forecast.wcPred), borderColor:CHART_COLORS.WORKING_CAPITAL, backgroundColor:'rgba(255,203,5,0.18)', borderDash:[6,4], fill:false, tension:0.4, pointRadius:3, borderWidth:3, yAxisID: 'y1' }
        ];

        forecastChart = new Chart(ctx, {
            type:'line', data:{ labels, datasets },
            options:{
                responsive:true, maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{ labels:{ color:CHART_COLORS.CHART_TEXT } },
                    annotation:{ annotations: { forecastZone: { type:'box', xMin: histLen-0.5, xMax: labels.length-0.5, backgroundColor:CHART_COLORS.FORECAST_ZONE_FILL } } },
                    tooltip:{ callbacks:{ label: ctx=>{ const val=ctx.parsed.y; if (val==null) return ''; return ctx.dataset.label + ': ' + formatNGN(val, 2); } } }
                }, scales:{
                    x:{ ticks:{ color:CHART_COLORS.CHART_TEXT } },
                    y: { ticks: { color: CHART_COLORS.RECEIVABLES, callback: v => formatNGN(v, 0) }, position: 'left' },
                    y1: { ticks: { color: CHART_COLORS.WORKING_CAPITAL, callback: v => formatNGN(v, 0) }, position: 'right', grid: { drawOnChartArea: false } }
                }
            }
        });
    }

    function buildAnnualChart(forecast){
        updateChartColors();
        const ctx = document.getElementById('annualChart').getContext('2d');
        if (annualChart) annualChart.destroy();

        const years = {};
        forecast.dates.forEach((d,i)=>{ const y = d.split('-')[0]; years[y] = years[y]||{rec:0,pay:0}; years[y].rec += forecast.recPred[i]; years[y].pay += forecast.payPred[i]; });
        const ys = Object.keys(years).sort();
        const recs = ys.map(y=>years[y].rec); const pays = ys.map(y=>years[y].pay);

        annualChart = new Chart(ctx, {
            type:'bar', data:{ labels: ys, datasets:[
                { label:'Receivables (Forecast sum) ΣR_f', data: recs, backgroundColor:CHART_COLORS.RECEIVABLES },
                { label:'Payables (Forecast sum) ΣP_f', data: pays, backgroundColor:CHART_COLORS.PAYABLES }
            ] }, options:{
                responsive:true, maintainAspectRatio:false,
                plugins:{
                    legend:{ labels:{ color:CHART_COLORS.CHART_TEXT } },
                    tooltip:{ mode:'index', intersect:false, callbacks:{ label: ctx=>{ const val=ctx.parsed.y; if (val==null) return ''; return ctx.dataset.label + ': ' + formatNGN(val, 2); } } }
                }, scales:{
                    x:{ stacked:true, ticks:{ color:CHART_COLORS.CHART_TEXT } },
                    y:{ stacked:true, ticks:{ callback:v=>formatNGN(v, 0), color:CHART_COLORS.CHART_TEXT } }
                }
            }
        });
    }


    // --- MAIN LOGIC AND EVENT LISTENERS (Focus on the fixed PDF export) ---
    document.addEventListener('DOMContentLoaded', () => {
        const elements = initializeDOMElements();

        // Placeholder for data processing and training (implied by the buttons)
        async function processData() {
            if (!rawWide) {
                setStatus('Error: No data loaded. Please upload an Excel file.');
                return;
            }
            setStatus('1/4: Training models...');
            elements.trainBtn.disabled = true;

            const windowSize = parseInt(elements.windowSizeInput.value) || 4;
            const epochs = parseInt(elements.epochsInput.value) || 40;
            const batchSize = parseInt(elements.batchSizeInput.value) || 8;
            const horizon = parseInt(elements.horizonInput.value) || 12;

            try {
                // 1. Train Receivables model
                const { preds: recPred } = await trainAndForecastSeries(rawWide.receivables, windowSize, epochs, batchSize, horizon);
                setStatus('2/4: Receivables forecast complete. Training Payables...');

                // 2. Train Payables model
                const { preds: payPred } = await trainAndForecastSeries(rawWide.payables, windowSize, epochs, batchSize, horizon);
                setStatus('3/4: Payables forecast complete. Generating report...');

                // 3. Generate combined forecast
                const lastDate = rawWide.dates[rawWide.dates.length - 1];
                let [lastY, lastM] = lastDate.split('-').map(Number);
                const forecastDates = [];
                for (let i = 0; i < horizon; i++) {
                    lastM += 1;
                    if (lastM > 12) { lastM = 1; lastY += 1; }
                    forecastDates.push(`${lastY}-${String(lastM).padStart(2, '0')}`);
                }

                const forecast = {
                    dates: forecastDates,
                    recPred,
                    payPred,
                    wcPred: recPred.map((r, i) => r - payPred[i])
                };
                
                // 4. Update Summary and Charts
                const lastWC = rawWide.receivables[rawWide.receivables.length - 1] - rawWide.payables[rawWide.payables.length - 1];
                const lastRec = rawWide.receivables[rawWide.receivables.length - 1];
                const lastPay = rawWide.payables[rawWide.payables.length - 1];
                const forecastWC = forecast.wcPred[forecast.wcPred.length - 1];
                const forecastRec = forecast.recPred[forecast.recPred.length - 1];
                const forecastPay = forecast.payPred[forecast.payPred.length - 1];

                elements.recEl.textContent = formatNGN(forecastRec, 0);
                elements.payEl.textContent = formatNGN(forecastPay, 0);
                elements.wcEl.textContent = formatNGN(forecastWC, 0);

                elements.timestampContent.textContent = `Report Generated: ${new Date().toLocaleString()} (Horizon: ${horizon} months)`;
                elements.horizonLabel.textContent = horizon;

                // AI Insights
                elements.aiInsightsContent.innerHTML = generateAIInsights(lastWC, forecastWC, lastRec, forecastRec, lastPay, forecastPay, horizon);
                elements.aiInsightsSection.classList.remove('hidden');

                // Build Charts
                buildHistoricalChart(rawWide);
                buildForecastChart(rawWide, forecast);
                buildAnnualChart(forecast);

                setStatus('4/4: Forecast complete. Charts updated!');
                
            } catch (e) {
                setStatus('Error: ' + e.message);
                console.error(e);
            } finally {
                elements.trainBtn.disabled = false;
            }
        }

        function reset() {
            rawWide = null;
            elements.recEl.textContent = elements.payEl.textContent = elements.wcEl.textContent = '—';
            elements.timestampContent.textContent = '';
            elements.aiInsightsSection.classList.add('hidden');
            if (historicalChart) historicalChart.destroy();
            if (forecastChart) forecastChart.destroy();
            if (annualChart) annualChart.destroy();
            setStatus('No data loaded.');
        }

        // --- Event Listeners ---
        elements.fileInput.addEventListener('change', (e) => {
            reset();
            const file = e.target.files[0];
            if (!file) { return; }
            setStatus(`Loading ${file.name}...`);
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                try {
                    const workbook = XLSX.read(data, { type: 'array' });
                    const parsed = parseWorkbook(workbook);

                    const dates = [...new Set([...parsed.receivables.map(r=>r.date), ...parsed.payables.map(p=>p.date)])].sort();
                    const recMap = new Map(parsed.receivables.map(r => [r.date, r.balance]));
                    const payMap = new Map(parsed.payables.map(p => [p.date, p.balance]));

                    rawWide = {
                        dates,
                        receivables: dates.map(d => recMap.get(d) || 0),
                        payables: dates.map(d => payMap.get(d) || 0)
                    };

                    if (rawWide.dates.length === 0) {
                        setStatus('Error: Could not find valid data columns (Date, Account, Balance) or no R/P accounts (1xx, 2xx).');
                        rawWide = null;
                        return;
                    }
                    setStatus(`Data loaded successfully (${rawWide.dates.length} periods). Ready to train.`);
                    buildHistoricalChart(rawWide); // Initial chart build
                } catch (e) {
                    setStatus('Error parsing file: ' + e.message);
                    console.error(e);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        elements.trainBtn.addEventListener('click', processData);
        elements.resetBtn.addEventListener('click', reset);

        // Horizon update listener
        elements.horizonInput.addEventListener('input', () => {
            const h = parseInt(elements.horizonInput.value);
            elements.horizonLabel.textContent = h > 0 ? h : '—';
        });

        // Theme Toggle
        elements.modeToggle.addEventListener('click', () => {
            const body = document.body;
            body.classList.toggle('dark');
            body.classList.toggle('light');
            // Re-render charts to update axis/label colors
            if (historicalChart) buildHistoricalChart(rawWide);
            if (forecastChart) buildForecastChart(rawWide, { dates:[], recPred:[], payPred:[], wcPred:[] }); // Needs full rebuild logic but using placeholders
            if (annualChart) annualChart.destroy(); // Destroy and let the next forecast rebuild it
        });


        // --- Export Functions ---

        elements.exportChartBtn.addEventListener('click', () => {
            if (!forecastChart) { setStatus('No forecast chart to export.'); return; }
            setStatus('Exporting PNG...');
            forecastChart.canvas.toBlob((blob) => {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `MTN_Working_Capital_Forecast_${new Date().toISOString().slice(0, 10)}.png`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                setStatus('PNG exported.');
            });
        });

        elements.exportPdfBtn.addEventListener('click', async () => {
            const { jsPDF } = window.jspdf;
            const target = elements.reportCaptureArea;
            setStatus('Preparing PDF export...');

            // Temporarily apply PDF-specific styling
            document.body.classList.add('pdf-exporting');

            // 1. Capture the Report Summary and Insights
            const summaryCanvas = await html2canvas(target, { scale: 2, useCORS: true, logging: false });
            
            // 2. Capture the Full Trajectory Chart (Forecast Chart)
            if (!forecastChart) { 
                setStatus('Error: No charts available for PDF export.'); 
                document.body.classList.remove('pdf-exporting');
                return; 
            }
            const forecastCanvas = document.getElementById('forecastChart');
            
            // Generate PDF
            const pdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();
            let y = 10;
            const margin = 10;

            // Add MTN Header to PDF
            pdf.setFontSize(24);
            pdf.setFont('helvetica', 'bold');
            pdf.setTextColor(255, 203, 5); // MTN Yellow
            pdf.setFillColor(0, 0, 0); // MTN Black
            pdf.rect(0, 0, pdfWidth, 20, 'F'); // Header bar
            pdf.text('MTN FinTech: Predictive Capital Engine', margin, 14);
            pdf.setDrawColor(255, 203, 5); // MTN Yellow line
            pdf.line(margin, 17, pdfWidth - margin, 17);

            y = 30;
            
            // 3. Add Summary/Insights image
            const summaryImgData = summaryCanvas.toDataURL('image/jpeg');
            const summaryRatio = summaryCanvas.width / summaryCanvas.height;
            const summaryImgWidth = pdfWidth - 2 * margin;
            const summaryImgHeight = summaryImgWidth / summaryRatio;
            pdf.addImage(summaryImgData, 'JPEG', margin, y, summaryImgWidth, summaryImgHeight);
            y += summaryImgHeight + 10;

            // 4. Add Chart image (using html2canvas on the chart's canvas element)
            const chartCanvas = await html2canvas(forecastCanvas, { scale: 2, useCORS: true, logging: false, backgroundColor: '#ffffff' }); // Ensure white background for chart
            if (y + 10 > pdfHeight) { pdf.addPage(); y = 10; }
            pdf.setFontSize(18);
            pdf.setTextColor(0, 0, 0);
            pdf.text('Full Working Capital Trajectory', margin, y);
            y += 5;
            const chartImgData = chartCanvas.toDataURL('image/jpeg');
            const chartRatio = chartCanvas.width / chartCanvas.height;
            const chartImgWidth = pdfWidth - 2 * margin;
            const chartImgHeight = chartImgWidth / chartRatio;
            
            pdf.addImage(chartImgData, 'JPEG', margin, y, chartImgWidth, chartImgHeight);
            y += chartImgHeight + 10;

            // Add Footer to PDF
            const footerText = `Confidential | Report Date: ${new Date().toLocaleDateString()} | © MTN FinOps Predictive Core`;
            pdf.setFontSize(10);
            pdf.setFont('helvetica', 'normal');
            pdf.setTextColor(100, 100, 100); 
            const footerY = pdfHeight - 8;
            pdf.text(footerText, pdfWidth / 2, footerY, { align: 'center' });

            // Final Save
            pdf.save(`MTN_WCP_Report_${new Date().toISOString().slice(0, 10)}.pdf`);
            
            // Revert styling
            document.body.classList.remove('pdf-exporting');
            setStatus('PDF exported successfully.');
        });
    });

</script>
</body>
</html>
