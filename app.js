/**
 * SheetCalc - Application Logic
 * Implements interactive spreadsheet generation, math tracing animations, and data exports.
 */

document.addEventListener('DOMContentLoaded', () => {
    // Apps Script Setup Code
    const APPS_SCRIPT_SOURCE = `function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById("1azRoUDoaCwqpzIftBMrCWGkURmkdLmfdMVJfTkQh3hM");
    var sheet = ss.getSheetByName("ISSUE DR") || ss.getSheets()[0];
    
    // Write metadata
    sheet.getRange("E2").setValue(data.ref);
    sheet.getRange("E3").setValue(data.date);
    
    // Write shipping details
    sheet.getRange("C2").setValue(data.transfer);
    sheet.getRange("C3").setValue(data.address1);
    sheet.getRange("C4").setValue(data.address2);
    
    // Write line item details
    sheet.getRange("A7").setValue(data.qty);
    sheet.getRange("C7").setValue(data.desc);
    
    // Write logistics summary
    sheet.getRange("C25").setValue(data.items);
    sheet.getRange("C26").setValue(data.cbm);
    
    // Write base rate
    sheet.getRange("C29").setValue(data.cnyRate);
    
    // Force spreadsheet to flush and recalculate formulas
    SpreadsheetApp.flush();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("J2N API Sync Active!");
}`;
    // DOM Elements
    const form = document.getElementById('generator-form');
    const inputRef = document.getElementById('input-ref');
    const inputDate = document.getElementById('input-date');
    const inputItems = document.getElementById('input-items');
    const inputCbm = document.getElementById('input-cbm');
    const inputDesc = document.getElementById('input-desc');
    const inputQty = document.getElementById('input-qty');
    const inputCnyRate = document.getElementById('input-cny-rate');

    const btnGenerate = document.getElementById('btn-generate');
    const btnToggleGrid = document.getElementById('btn-toggle-grid');
    const btnToggleRows = document.getElementById('btn-toggle-rows');
    const btnTheme = document.getElementById('btn-theme');

    const inputGsheetUrl = document.getElementById('input-gsheet-url');
    const btnShowInstructions = document.getElementById('btn-show-instructions');
    const btnCloseModal = document.getElementById('btn-close-modal');
    const btnCopyCode = document.getElementById('btn-copy-code');
    const instructionsModal = document.getElementById('instructions-modal');
    const appsScriptCode = document.getElementById('apps-script-code');

    const sheet = document.getElementById('invoice-sheet');
    const collapsedIndicator = document.getElementById('collapsed-indicator');
    const collapsibleRowsContainer = document.getElementById('collapsible-rows-container');

    const formulaPopup = document.getElementById('formula-popup');

    // Canvas overlay for animations
    let canvas, ctx;
    let particles = [];
    let activePaths = [];
    let animationFrameId = null;

    // State Variables
    let isRowsExpanded = false;
    let isDarkTheme = false;
    let currentCalculationData = null;

    // Initialize Page
    initRefInput();
    initDateInput();
    initCollapsedRows();
    setupCanvas();
    setupCellInteractions();
    runInitialCalculations();
    initSyncSettings();

    // Event Listeners
    const btnFetchRow = document.getElementById('btn-fetch-row');
    const inputRowNum = document.getElementById('input-row-num');
    const rowInputWrapper = document.getElementById('row-input-wrapper');

    btnGenerate.addEventListener('click', generateInvoice);
    btnToggleGrid.addEventListener('click', toggleGrid);
    btnToggleRows.addEventListener('click', toggleRows);
    collapsedIndicator.addEventListener('click', toggleRows);
    btnTheme.addEventListener('click', toggleTheme);

    if (btnFetchRow && inputRowNum) {
        btnFetchRow.addEventListener('click', () => {
            fetchRowData(parseInt(inputRowNum.value));
        });

        inputRowNum.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                fetchRowData(parseInt(inputRowNum.value));
            }
        });
    }

    // Modal Event Listeners
    btnShowInstructions.addEventListener('click', () => {
        appsScriptCode.textContent = APPS_SCRIPT_SOURCE;
        instructionsModal.classList.add('show');
    });
    btnCloseModal.addEventListener('click', () => {
        instructionsModal.classList.remove('show');
    });
    instructionsModal.addEventListener('click', (e) => {
        if (e.target === instructionsModal) {
            instructionsModal.classList.remove('show');
        }
    });
    btnCopyCode.addEventListener('click', () => {
        navigator.clipboard.writeText(APPS_SCRIPT_SOURCE).then(() => {
            btnCopyCode.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
            setTimeout(() => {
                btnCopyCode.innerHTML = '<i class="fa-solid fa-copy"></i> Copy';
            }, 2000);
        });
    });

    inputGsheetUrl.addEventListener('change', () => {
        localStorage.setItem('j2n_gsheet_url', inputGsheetUrl.value.trim());
    });

    function initSyncSettings() {
        const savedUrl = localStorage.getItem('j2n_gsheet_url');
        if (savedUrl) {
            inputGsheetUrl.value = savedUrl;
        }
    }

    // Sync input handlers for real-time visual feedback on non-computed cells
    inputRef.addEventListener('input', (e) => {
        document.getElementById('cell-e2').textContent = e.target.value;
    });
    inputDate.addEventListener('input', (e) => {
        document.getElementById('cell-e3').textContent = formatDateToDdMmmYyyy(e.target.value);
    });
    inputItems.addEventListener('input', (e) => {
        document.getElementById('cell-items').innerHTML = `<span class="print-only-label">ITEMS</span><span class="value-span">${e.target.value}</span>`;
    });
    inputCbm.addEventListener('input', (e) => {
        document.getElementById('cell-cbm').innerHTML = `<span class="print-only-label">CBM</span><span class="value-span">${e.target.value}</span>`;
    });
    inputDesc.addEventListener('input', (e) => {
        document.getElementById('cell-c7').textContent = e.target.value;
    });
    inputQty.addEventListener('input', (e) => {
        document.getElementById('cell-a7').textContent = formatQty(e.target.value);
    });
    inputCnyRate.addEventListener('input', (e) => {
        document.getElementById('cell-c29').innerHTML = `<span class="print-only-label">CNY</span><span class="value-span">${formatRate(e.target.value)}</span>`;
    });

    /**
     * Set reference input default based on current year and month (YY-MM)
     */
    function initRefInput() {
        const today = new Date();
        const yy = String(today.getFullYear()).slice(-2);
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        inputRef.value = `${yy}-${mm}`;
    }

    /**
     * Set date input to today
     */
    function initDateInput() {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        inputDate.value = `${yyyy}-${mm}-${dd}`;
    }

    /**
     * Format date string to dd-Mmm-yyyy (e.g. 11-Jun-2026) in a timezone-safe manner
     */
    function formatDateToDdMmmYyyy(dateStr) {
        if (!dateStr) return '';
        const parts = dateStr.split('-');
        if (parts.length !== 3) return dateStr;
        const year = parts[0];
        const monthIndex = parseInt(parts[1], 10) - 1;
        const day = parts[2].padStart(2, '0');
        
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const monthName = months[monthIndex] || 'Jan';
        
        return `${day}-${monthName}-${year}`;
    }

    /**
     * Format decimal rate values to 4 decimal places
     */
    function formatRate(num) {
        const val = parseFloat(num);
        return isNaN(val) ? '' : val.toFixed(4);
    }

    /**
     * Format quantity values with thousands separators
     */
    function formatQty(num) {
        const val = parseFloat(num);
        return isNaN(val) ? '' : val.toLocaleString('en-US');
    }

    /**
     * Format currency totals to 2 decimal places with thousands separator
     */
    function formatMoney(num) {
        const val = parseFloat(num);
        return isNaN(val) ? '' : val.toLocaleString('en-US', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    }

    /**
     * Dynamically insert empty rows 11 to 24
     */
    function initCollapsedRows() {
        collapsibleRowsContainer.innerHTML = '';
        for (let row = 11; row <= 24; row++) {
            const rowEl = document.createElement('div');
            rowEl.className = 'sheet-row';
            rowEl.setAttribute('data-row', row);
            
            rowEl.innerHTML = `
                <div class="row-num-col">${row}</div>
                <div class="sheet-cell col-a"></div>
                <div class="sheet-cell col-b"></div>
                <div class="sheet-cell col-c"></div>
                <div class="sheet-cell col-d"></div>
                <div class="sheet-cell col-e"></div>
            `;
            collapsibleRowsContainer.appendChild(rowEl);
        }
    }

    /**
     * Initialize canvas element and size it to match the spreadsheet
     */
    function setupCanvas() {
        canvas = document.createElement('canvas');
        canvas.id = 'animation-canvas';
        canvas.style.position = 'absolute';
        canvas.style.top = '0';
        canvas.style.left = '0';
        canvas.style.width = '100%';
        canvas.style.height = '100%';
        canvas.style.pointerEvents = 'none';
        canvas.style.zIndex = '99';
        sheet.appendChild(canvas);
        ctx = canvas.getContext('2d');
        
        resizeCanvas();
        window.addEventListener('resize', resizeCanvas);
    }

    function resizeCanvas() {
        if (canvas) {
            canvas.width = sheet.offsetWidth;
            canvas.height = sheet.offsetHeight;
        }
    }

    /**
     * Set up interactions when clicking cells (selecting matching form field)
     */
    function setupCellInteractions() {
        const cellMap = {
            'cell-e2': inputRef,
            'cell-e3': inputDate,
            'cell-items': inputItems,
            'cell-cbm': inputCbm,
            'cell-a7': inputQty,
            'cell-c7': inputDesc,
            'cell-c29': inputCnyRate
        };

        // Click handler to select and highlight cells
        document.querySelectorAll('.cell-highlightable').forEach(cell => {
            cell.addEventListener('click', (e) => {
                document.querySelectorAll('.cell-highlightable').forEach(c => c.classList.remove('cell-selected'));
                cell.classList.add('cell-selected');

                const matchingInput = cellMap[cell.id];
                if (matchingInput) {
                    matchingInput.focus();
                    matchingInput.select();
                } else {
                    // For computed cells, highlight the source cells instead
                    flashSourceCellsForComputed(cell.id);
                }
            });
        });
    }

    /**
     * Highlights spreadsheet sources when clicking formula-derived cells
     */
    function flashSourceCellsForComputed(cellId) {
        let sources = [];
        if (cellId === 'cell-c31') sources = ['cell-c29'];
        else if (cellId === 'cell-d7') sources = ['cell-c31'];
        else if (cellId === 'cell-e7') sources = ['cell-a7', 'cell-d7'];
        else if (cellId === 'cell-e38') sources = ['cell-e7'];

        sources.forEach(srcId => {
            const el = document.getElementById(srcId);
            if (el) {
                el.classList.add('cell-selected');
                setTimeout(() => el.classList.remove('cell-selected'), 1000);
            }
        });
    }

    /**
     * Perform initial fast calculations without step animations on load
     */
    function runInitialCalculations() {
        const refVal = inputRef.value;
        const dateVal = inputDate.value;
        const transferVal = 'DMC - Marlon';
        const address1Val = '22 Ford Ave., Doña Manuela Subd.,';
        const address2Val = 'Pamplona Tres, Las Piñas';
        const itemsVal = inputItems.value;
        const cbmVal = inputCbm.value;
        const descVal = inputDesc.value;
        const qtyVal = parseFloat(inputQty.value) || 0;
        const cnyRateVal = parseFloat(inputCnyRate.value) || 0;

        const markupRate = cnyRateVal * 1.05;
        const totalCny = qtyVal * markupRate;

        // Populate values
        document.getElementById('cell-e2').textContent = refVal;
        document.getElementById('cell-e3').textContent = formatDateToDdMmmYyyy(dateVal);
        document.getElementById('cell-transfer').textContent = transferVal;
        document.getElementById('cell-address-1').textContent = address1Val;
        document.getElementById('cell-address-2').textContent = address2Val;
        document.getElementById('cell-items').innerHTML = `<span class="print-only-label">ITEMS</span><span class="value-span">${itemsVal}</span>`;
        document.getElementById('cell-cbm').innerHTML = `<span class="print-only-label">CBM</span><span class="value-span">${cbmVal}</span>`;
        document.getElementById('cell-a7').textContent = formatQty(qtyVal);
        document.getElementById('cell-c7').textContent = descVal;
        document.getElementById('cell-c29').innerHTML = `<span class="print-only-label">CNY</span><span class="value-span">${formatRate(cnyRateVal)}</span>`;

        document.getElementById('cell-c31').innerHTML = `<span class="print-only-label">RATE</span><span class="value-span">${formatRate(markupRate)}</span>`;
        document.getElementById('cell-d7').textContent = formatRate(markupRate);
        document.getElementById('cell-e7').textContent = formatMoney(totalCny);
        document.getElementById('cell-e38').textContent = formatMoney(totalCny);

        currentCalculationData = {
            ref: refVal,
            date: dateVal,
            transfer: transferVal,
            address1: address1Val,
            address2: address2Val,
            items: itemsVal,
            cbm: cbmVal,
            desc: descVal,
            qty: qtyVal,
            cnyRate: cnyRateVal,
            markupRate: markupRate,
            totalCny: totalCny
        };
    }

    /**
     * Main calculation function invoked by the GENERATE button
     * Triggers Google Sheet background syncing and visual animation sequences
     */
    async function generateInvoice() {
        // Validate form
        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }

        // Open Google Sheet in a new tab immediately to bypass browser popup blockers
        window.open('https://docs.google.com/spreadsheets/d/1azRoUDoaCwqpzIftBMrCWGkURmkdLmfdMVJfTkQh3hM/edit?gid=165820978#gid=165820978', '_blank');

        // Get Input Data
        const refVal = inputRef.value;
        const dateVal = inputDate.value;
        const transferVal = 'DMC - Marlon';
        const address1Val = '22 Ford Ave., Doña Manuela Subd.,';
        const address2Val = 'Pamplona Tres, Las Piñas';
        const itemsVal = inputItems.value;
        const cbmVal = inputCbm.value;
        const descVal = inputDesc.value;
        const qtyVal = parseFloat(inputQty.value) || 0;
        const cnyRateVal = parseFloat(inputCnyRate.value) || 0;

        // Perform calculations
        const markupRate = cnyRateVal * 1.05;
        const totalCny = qtyVal * markupRate;

        currentCalculationData = {
            ref: refVal,
            date: dateVal,
            transfer: transferVal,
            address1: address1Val,
            address2: address2Val,
            items: itemsVal,
            cbm: cbmVal,
            desc: descVal,
            qty: qtyVal,
            cnyRate: cnyRateVal,
            markupRate: markupRate,
            totalCny: totalCny
        };

        const syncUrl = inputGsheetUrl.value.trim();
        if (syncUrl) {
            // Set Loading state on Generate button
            const origContent = btnGenerate.innerHTML;
            btnGenerate.disabled = true;
            btnGenerate.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Syncing to Sheets...';

            try {
                // Send plain text POST request to avoid CORS issues with Apps Script redirect redirects
                await fetch(syncUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: {
                        'Content-Type': 'text/plain'
                    },
                    body: JSON.stringify(currentCalculationData)
                });
                
                // Show success feedback
                btnGenerate.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Synced to Google Sheets!';
                btnGenerate.style.background = 'linear-gradient(135deg, #10b981, #059669)';
                
                setTimeout(() => {
                    btnGenerate.disabled = false;
                    btnGenerate.innerHTML = origContent;
                    btnGenerate.style.background = '';
                }, 2500);

            } catch (error) {
                console.error("GSheet sync failed:", error);
                alert("Google Sheets sync failed. Please check your Web App URL. Generating local invoice instead.");
                btnGenerate.disabled = false;
                btnGenerate.innerHTML = origContent;
            }
        }

        // Clear all derived cells to show recalculation
        const computedCells = ['cell-c31', 'cell-d7', 'cell-e7', 'cell-e38'];
        computedCells.forEach(cellId => {
            const el = document.getElementById(cellId);
            el.textContent = '---';
            el.classList.remove('pulse-animate');
        });

        // Sync raw inputs
        document.getElementById('cell-e2').textContent = refVal;
        document.getElementById('cell-e3').textContent = formatDateToDdMmmYyyy(dateVal);
        document.getElementById('cell-transfer').textContent = transferVal;
        document.getElementById('cell-address-1').textContent = address1Val;
        document.getElementById('cell-address-2').textContent = address2Val;
        document.getElementById('cell-items').innerHTML = `<span class="print-only-label">ITEMS</span><span class="value-span">${itemsVal}</span>`;
        document.getElementById('cell-cbm').innerHTML = `<span class="print-only-label">CBM</span><span class="value-span">${cbmVal}</span>`;
        document.getElementById('cell-a7').textContent = formatQty(qtyVal);
        document.getElementById('cell-c7').textContent = descVal;
        document.getElementById('cell-c29').innerHTML = `<span class="print-only-label">CNY</span><span class="value-span">${formatRate(cnyRateVal)}</span>`;

        // Cancel any active animation loops
        cancelAnimationFrame(animationFrameId);
        activePaths = [];
        particles = [];
        ctx.clearRect(0, 0, canvas.width, canvas.height);

        // Define Sequential Formula Animation Steps
        const steps = [
            {
                // Step 1: Base Rate (C29) markup * 1.05 -> Markup Rate (C31)
                startId: 'cell-c29',
                endId: 'cell-c31',
                color: '#8b5cf6', // purple
                duration: 1000,
                formulaText: `${formatRate(cnyRateVal)} × 1.05 = ${formatRate(markupRate)}`,
                onComplete: () => {
                    const cell = document.getElementById('cell-c31');
                    cell.innerHTML = `<span class="print-only-label">RATE</span><span class="value-span">${formatRate(markupRate)}</span>`;
                    cell.classList.add('pulse-animate');
                }
            },
            {
                // Step 2: Markup Rate (C31) copies to Unit Price (D7)
                startId: 'cell-c31',
                endId: 'cell-d7',
                color: '#06b6d4', // teal
                duration: 1200,
                formulaText: `D7 = C31 (${formatRate(markupRate)})`,
                onComplete: () => {
                    const cell = document.getElementById('cell-d7');
                    cell.textContent = formatRate(markupRate);
                    cell.classList.add('pulse-animate');
                }
            },
            {
                // Step 3: Quantity (A7) & Price (D7) multiply into Total (E7)
                // Dual path merge visual
                dualStart: true,
                startIds: ['cell-a7', 'cell-d7'],
                endId: 'cell-e7',
                color: '#10b981', // green
                duration: 1300,
                formulaText: `${qtyVal} × ${formatRate(markupRate)} = ${formatMoney(totalCny)}`,
                onComplete: () => {
                    const cell = document.getElementById('cell-e7');
                    cell.textContent = formatMoney(totalCny);
                    cell.classList.add('pulse-animate');
                }
            },
            {
                // Step 4: Total (E7) cascades to Grand Total (E38)
                // Note: If rows are collapsed, this line will span down into the summary area beautifully
                startId: 'cell-e7',
                endId: 'cell-e38',
                color: '#3b82f6', // blue
                duration: 1000,
                formulaText: `E38 = E7 (${formatMoney(totalCny)})`,
                onComplete: () => {
                    const cell = document.getElementById('cell-e38');
                    cell.textContent = formatMoney(totalCny);
                    cell.classList.add('pulse-animate');
                    triggerConfetti('cell-e38');
                }
            }
        ];

        // Begin Step Sequential Execution
        let currentStepIndex = 0;

        function runStep() {
            if (currentStepIndex >= steps.length) {
                // Done! Clean canvas after short delay
                setTimeout(() => {
                    activePaths = [];
                    fadeOutCanvas();
                }, 2000);
                return;
            }

            const step = steps[currentStepIndex];
            
            if (step.dualStart) {
                // Setup two parallel moving paths
                const path1 = createPathConfig(step.startIds[0], step.endId, step.color, step.duration);
                const path2 = createPathConfig(step.startIds[1], step.endId, step.color, step.duration);
                activePaths = [path1, path2];
                showFormulaPopup(step.endId, step.formulaText);
            } else {
                const path = createPathConfig(step.startId, step.endId, step.color, step.duration);
                activePaths = [path];
                showFormulaPopup(step.endId, step.formulaText);
            }

            const startTime = performance.now();

            function animateStep(timestamp) {
                const elapsed = timestamp - startTime;
                const progress = Math.min(elapsed / step.duration, 1);

                // Update path progresses
                activePaths.forEach(p => p.progress = progress);

                // Spawn flow particles
                if (progress < 0.95 && Math.random() < 0.3) {
                    activePaths.forEach(p => {
                        particles.push({
                            x: p.startX,
                            y: p.startY,
                            startX: p.startX,
                            startY: p.startY,
                            ctrlX: p.ctrlX,
                            ctrlY: p.ctrlY,
                            endX: p.endX,
                            endY: p.endY,
                            t: 0,
                            speed: 0.02 + Math.random() * 0.015,
                            size: 2 + Math.random() * 3,
                            color: step.color
                        });
                    });
                }

                // Render frame
                renderFrame();

                if (progress < 1) {
                    animationFrameId = requestAnimationFrame(animateStep);
                } else {
                    // Step complete
                    hideFormulaPopup();
                    step.onComplete();
                    currentStepIndex++;
                    setTimeout(runStep, 400); // Small interval between formulas
                }
            }

            animationFrameId = requestAnimationFrame(animateStep);
        }

        runStep();
    }

    /**
     * Create path coordinates and Control points for Bezier curves
     */
    function createPathConfig(startId, endId, color, duration) {
        const start = getCellCenter(startId);
        const end = getCellCenter(endId);

        // Bezier curvature: offset control point slightly from midpoint
        const midX = (start.x + end.x) / 2;
        const midY = (start.y + end.y) / 2;
        const dx = end.x - start.x;
        const dy = end.y - start.y;
        
        // Perpendicular offset for curved look
        const ctrlX = midX - dy * 0.15;
        const ctrlY = midY + dx * 0.15;

        return {
            startX: start.x,
            startY: start.y,
            endX: end.x,
            endY: end.y,
            ctrlX,
            ctrlY,
            color,
            progress: 0
        };
    }

    /**
     * Calculate position along quadratic Bezier curve
     */
    function getBezierPoint(t, p0, p1, p2) {
        const x = (1 - t) * (1 - t) * p0.x + 2 * (1 - t) * t * p1.x + t * t * p2.x;
        const y = (1 - t) * (1 - t) * p0.y + 2 * (1 - t) * t * p1.y + t * t * p2.y;
        return { x, y };
    }

    /**
     * Main rendering step for canvas drawings
     */
    function renderFrame() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);

        // 1. Draw connecting guide lines
        activePaths.forEach(path => {
            ctx.beginPath();
            ctx.moveTo(path.startX, path.startY);
            ctx.quadraticCurveTo(path.ctrlX, path.ctrlY, path.endX, path.endY);
            
            // Background guide line (translucent)
            ctx.strokeStyle = path.color;
            ctx.globalAlpha = 0.15;
            ctx.lineWidth = 4;
            ctx.stroke();

            // Animated progress line overlay
            ctx.globalAlpha = 0.8;
            ctx.lineWidth = 3;
            ctx.strokeStyle = path.color;
            ctx.shadowColor = path.color;
            ctx.shadowBlur = 10;

            // Draw line up to current progress
            ctx.beginPath();
            ctx.moveTo(path.startX, path.startY);
            const steps = Math.floor(path.progress * 40);
            for (let i = 0; i <= steps; i++) {
                const t = i / 40;
                const pt = getBezierPoint(t, 
                    {x: path.startX, y: path.startY}, 
                    {x: path.ctrlX, y: path.ctrlY}, 
                    {x: path.endX, y: path.endY}
                );
                ctx.lineTo(pt.x, pt.y);
            }
            ctx.stroke();
            ctx.shadowBlur = 0;
            ctx.globalAlpha = 1.0;

            // Draw glowing bead at path progress header
            const beadPt = getBezierPoint(path.progress, 
                {x: path.startX, y: path.startY}, 
                {x: path.ctrlX, y: path.ctrlY}, 
                {x: path.endX, y: path.endY}
            );
            ctx.beginPath();
            ctx.arc(beadPt.x, beadPt.y, 6, 0, Math.PI * 2);
            ctx.fillStyle = '#ffffff';
            ctx.shadowColor = path.color;
            ctx.shadowBlur = 12;
            ctx.fill();
            ctx.shadowBlur = 0;
        });

        // 2. Animate and draw flowing micro-particles
        for (let i = particles.length - 1; i >= 0; i--) {
            const p = particles[i];
            p.t += p.speed;

            if (p.t >= 1) {
                particles.splice(i, 1);
                continue;
            }

            const pos = getBezierPoint(p.t, 
                {x: p.startX, y: p.startY}, 
                {x: p.ctrlX, y: p.ctrlY}, 
                {x: p.endX, y: p.endY}
            );

            ctx.beginPath();
            ctx.arc(pos.x, pos.y, p.size, 0, Math.PI * 2);
            ctx.fillStyle = p.color;
            ctx.globalAlpha = 1 - p.t;
            ctx.fill();
            ctx.globalAlpha = 1.0;
        }
    }

    /**
     * Show floating formula popup above targeted sheet cells
     */
    function showFormulaPopup(targetId, text) {
        const target = document.getElementById(targetId);
        if (!target) return;

        formulaPopup.textContent = text;
        formulaPopup.classList.add('show');

        // Position popup centered above target cell
        const sheetRect = sheet.getBoundingClientRect();
        const targetRect = target.getBoundingClientRect();

        const x = targetRect.left - sheetRect.left + (targetRect.width / 2) - (formulaPopup.offsetWidth / 2);
        const y = targetRect.top - sheetRect.top - formulaPopup.offsetHeight - 8;

        formulaPopup.style.left = `${x}px`;
        formulaPopup.style.top = `${y}px`;
    }

    function hideFormulaPopup() {
        formulaPopup.classList.remove('show');
    }

    /**
     * Slowly fades out the canvas overlay once animation sequence finishes
     */
    function fadeOutCanvas() {
        let opacity = 1;
        const fadeInterval = setInterval(() => {
            opacity -= 0.05;
            if (opacity <= 0) {
                clearInterval(fadeInterval);
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                canvas.style.opacity = 1;
            } else {
                canvas.style.opacity = opacity;
            }
        }, 30);
    }

    /**
     * Trigger a colorful sparkle explosion (confetti) at specific cell centers
     */
    function triggerConfetti(cellId) {
        const center = getCellCenter(cellId);
        if (!center) return;

        const particleCount = 40;
        const colors = ['#06b6d4', '#8b5cf6', '#10b981', '#3b82f6', '#f59e0b', '#ec4899'];
        
        let explosionParticles = [];

        for (let i = 0; i < particleCount; i++) {
            const angle = Math.random() * Math.PI * 2;
            const velocity = 2 + Math.random() * 6;
            explosionParticles.push({
                x: center.x,
                y: center.y,
                vx: Math.cos(angle) * velocity,
                vy: Math.sin(angle) * velocity - 1, // slight upward bias
                size: 3 + Math.random() * 4,
                color: colors[Math.floor(Math.random() * colors.length)],
                alpha: 1.0,
                decay: 0.015 + Math.random() * 0.02
            });
        }

        function animateConfetti() {
            if (explosionParticles.length === 0) return;

            // Clear previous guide lines, only draw confetti
            ctx.clearRect(0, 0, canvas.width, canvas.height);

            for (let i = explosionParticles.length - 1; i >= 0; i--) {
                const p = explosionParticles[i];
                p.x += p.vx;
                p.y += p.vy;
                p.vy += 0.1; // gravity
                p.alpha -= p.decay;

                if (p.alpha <= 0) {
                    explosionParticles.splice(i, 1);
                    continue;
                }

                ctx.beginPath();
                ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
                ctx.fillStyle = p.color;
                ctx.globalAlpha = p.alpha;
                ctx.fill();
                ctx.globalAlpha = 1.0;
            }

            if (explosionParticles.length > 0) {
                requestAnimationFrame(animateConfetti);
            }
        }

        animateConfetti();
    }

    /**
     * Helper to retrieve coordinates of the center of a cell relative to the sheet
     */
    function getCellCenter(cellId) {
        const cell = document.getElementById(cellId);
        if (!cell) return { x: 0, y: 0 };
        return {
            x: cell.offsetLeft + (cell.offsetWidth / 2),
            y: cell.offsetTop + (cell.offsetHeight / 2)
        };
    }

    /**
     * Toggle grid lines on sheet simulation
     */
    function toggleGrid() {
        sheet.classList.toggle('hide-grid');
        btnToggleGrid.classList.toggle('active');
    }

    /**
     * Toggle expansion of empty spreadsheet rows (11 to 27)
     */
    function toggleRows() {
        isRowsExpanded = !isRowsExpanded;
        
        if (isRowsExpanded) {
            collapsibleRowsContainer.classList.add('expanded');
            collapsibleRowsContainer.classList.remove('collapsed');
            collapsedIndicator.querySelector('.indicator-text').innerHTML = '<i class="fa-solid fa-arrows-up-down"></i> Click to collapse empty rows (11 - 24)';
            btnToggleRows.classList.add('active');
        } else {
            collapsibleRowsContainer.classList.remove('expanded');
            collapsibleRowsContainer.classList.add('collapsed');
            collapsedIndicator.querySelector('.indicator-text').innerHTML = '<i class="fa-solid fa-arrows-up-down"></i> Rows 11 - 24 collapsed (Empty)';
            btnToggleRows.classList.remove('active');
        }

        // Must resize canvas and recalculate heights since elements shifted layout
        setTimeout(() => {
            resizeCanvas();
        }, 400);
    }

    /**
     * Toggle Sheet preview dark/light mode
     */
    function toggleTheme() {
        isDarkTheme = !isDarkTheme;
        if (isDarkTheme) {
            sheet.setAttribute('data-sheet-theme', 'dark');
            btnTheme.querySelector('i').className = 'fa-solid fa-sun';
            btnTheme.classList.add('active');
        } else {
            sheet.removeAttribute('data-sheet-theme');
            btnTheme.querySelector('i').className = 'fa-solid fa-moon';
            btnTheme.classList.remove('active');
        }
    }

    /**
     * Triggers Print Stylesheets to export clean PDF invoices
     */
    function exportToPDF() {
        // Expand rows first to ensure print fits if they were editing, or print what is visible.
        // Clean CSS media prints hide the empty lines regardless, so just call print.
        window.print();
    }

    /**
     * Generates a CSV layout of calculated values and downloads it
     */
    function exportToCSV() {
        if (!currentCalculationData) {
            alert("Please click Generate to create the sheet data first!");
            return;
        }

        const data = currentCalculationData;
        
        let csvContent = "data:text/csv;charset=utf-8,";
        
        // CSV Lines
        const lines = [
            ["J2N"],
            [],
            ["Transfer to:", data.transfer],
            ["Address Line 1:", data.address1],
            ["Address Line 2:", data.address2],
            ["Reference #", data.ref],
            ["Date", formatDateToDdMmmYyyy(data.date)],
            ["ITEMS", data.items],
            ["CBM", data.cbm],
            [],
            ["QTY", "UNIT", "DESCRIPTION", "UNIT PRICE (CNY)", "TOTAL AMOUNT (CNY)"],
            [data.qty, "PCS", data.desc, formatRate(data.markupRate), formatMoney(data.totalCny)],
            [],
            ["CNY", formatRate(data.cnyRate)],
            ["factor", "1.05"],
            ["RATE", formatRate(data.markupRate)],
            [],
            ["GRAND TOTAL", formatMoney(data.totalCny)]
        ];

        // Format CSV matrix
        lines.forEach(row => {
            const escapedRow = row.map(val => {
                if (typeof val === 'string') {
                    // Escape quotes and wrap in quotes if commas exist
                    const escapedStr = val.replace(/"/g, '""');
                    return escapedStr.includes(',') ? `"${escapedStr}"` : escapedStr;
                }
                return val;
            });
            csvContent += escapedRow.join(",") + "\n";
        });

        // Download link trigger
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `Invoice_Sheet_${data.ref || 'export'}.csv`);
        document.body.appendChild(link);
        
        link.click();
        document.body.removeChild(link);
    }

    async function fetchRowData(rowNum) {
        if (!rowNum || rowNum < 1) {
            alert("Please enter a valid row number greater than 0.");
            return;
        }

        const icon = btnFetchRow.querySelector('i');
        const originalIconClass = icon.className;
        
        // Show loading state
        rowInputWrapper.classList.add('loading');
        icon.className = 'fa-solid fa-circle-notch fa-spin';
        btnFetchRow.disabled = true;

        try {
            const url = `https://docs.google.com/spreadsheets/d/1azRoUDoaCwqpzIftBMrCWGkURmkdLmfdMVJfTkQh3hM/export?format=csv&gid=311571294&t=${Date.now()}`;
            const response = await fetch(url);
            if (!response.ok) throw new Error('Network response was not ok');
            const csvText = await response.text();

            if (typeof Papa === 'undefined') {
                throw new Error('CSV parser is not loaded. Please check your internet connection.');
            }

            Papa.parse(csvText, {
                header: false,
                skipEmptyLines: false,
                complete: function(results) {
                    const data = results.data;
                    const rowIndex = rowNum - 1;

                    if (rowIndex >= data.length || rowIndex < 0) {
                        alert(`Row ${rowNum} is out of bounds (max rows: ${data.length}).`);
                        resetLoadingState();
                        return;
                    }

                    const row = data[rowIndex];
                    if (!row || row.length === 0 || row.every(cell => !cell)) {
                        alert(`Row ${rowNum} is empty.`);
                        resetLoadingState();
                        return;
                    }

                    const colY = (row[24] || '').trim();
                    const refVal = colY.substring(0, 7);
                    
                    // Reference # - first 7 characters of col Y
                    inputRef.value = refVal;

                    // Date - col V (index 21)
                    const dateVal = cleanDate(row[21]);
                    if (dateVal) {
                        inputDate.value = dateVal;
                    }

                    // ITEMS - same as Reference #
                    inputItems.value = refVal;

                    // CBM - replace last 3 chars of Reference # with last 3 chars of col Y
                    const last3 = colY.length >= 3 ? colY.slice(-3) : colY;
                    const cbmVal = refVal.slice(0, 4) + last3;
                    inputCbm.value = cbmVal;

                    // Description - Col C (index 2)
                    inputDesc.value = (row[2] || '').trim();

                    // Quantity - Col E (index 4)
                    const qtyVal = parseFloat((row[4] || '').replace(/,/g, '')) || 0;
                    inputQty.value = qtyVal;

                    // CNY Rate - Col O (index 14)
                    const cnyRateVal = parseFloat((row[14] || '').replace(/,/g, '')) || 0;
                    inputCnyRate.value = cnyRateVal;

                    // Trigger the calculations
                    runInitialCalculations();

                    // Success animation
                    flashInputs();
                    resetLoadingState();
                },
                error: function(err) {
                    console.error("CSV parse error:", err);
                    alert("Failed to parse sheet data.");
                    resetLoadingState();
                }
            });
        } catch (error) {
            console.error("Fetch error:", error);
            alert("Failed to fetch Google Sheet data. Error: " + error.message);
            resetLoadingState();
        }

        function resetLoadingState() {
            rowInputWrapper.classList.remove('loading');
            icon.className = originalIconClass;
            btnFetchRow.disabled = false;
        }
    }

    function cleanDate(dateStr) {
        if (!dateStr) return '';
        dateStr = dateStr.trim();
        
        // Handle "Jun-21- Sun" or "Mar-6- Fri"
        const monthMap = {
            'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
            'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
            'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
        };
        
        const match = dateStr.match(/^([A-Za-z]{3})-(\d{1,2})-\s*[A-Za-z]{3}$/);
        if (match) {
            const monthName = match[1].toLowerCase();
            const month = monthMap[monthName] || '01';
            const day = match[2].padStart(2, '0');
            let year = '2026';
            const refInputVal = document.getElementById('input-ref').value;
            if (refInputVal && /^\d{2}-/.test(refInputVal)) {
                year = '20' + refInputVal.substring(0, 2);
            } else {
                year = String(new Date().getFullYear());
            }
            return `${year}-${month}-${day}`;
        }
        
        // Handle short "Jun-21" format
        const matchShort = dateStr.match(/^([A-Za-z]{3})-(\d{1,2})/);
        if (matchShort) {
            const monthName = matchShort[1].toLowerCase();
            const month = monthMap[monthName] || '01';
            const day = matchShort[2].padStart(2, '0');
            let year = '2026';
            const refInputVal = document.getElementById('input-ref').value;
            if (refInputVal && /^\d{2}-/.test(refInputVal)) {
                year = '20' + refInputVal.substring(0, 2);
            } else {
                year = String(new Date().getFullYear());
            }
            return `${year}-${month}-${day}`;
        }

        const parts = dateStr.split('/');
        if (parts.length === 3) {
            let month = parts[0].padStart(2, '0');
            let day = parts[1].padStart(2, '0');
            let year = parts[2];
            if (year.length === 2) {
                year = '20' + year;
            }
            return `${year}-${month}-${day}`;
        }
        
        const isoParts = dateStr.split('-');
        if (isoParts.length === 3 && isoParts[0].length === 4) {
            return dateStr;
        }

        try {
            const d = new Date(dateStr);
            if (!isNaN(d.getTime())) {
                const yyyy = d.getFullYear();
                const mm = String(d.getMonth() + 1).padStart(2, '0');
                const dd = String(d.getDate()).padStart(2, '0');
                return `${yyyy}-${mm}-${dd}`;
            }
        } catch (e) {}

        return '';
    }

    function flashInputs() {
        const inputs = [inputRef, inputDate, inputItems, inputCbm, inputDesc, inputQty, inputCnyRate];
        inputs.forEach(input => {
            input.style.transition = 'background-color 0.3s ease';
            input.style.backgroundColor = 'rgba(16, 185, 129, 0.2)';
            setTimeout(() => {
                input.style.backgroundColor = '';
            }, 800);
        });
    }
});
