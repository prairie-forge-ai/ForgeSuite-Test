(() => {
    const HEADCOUNT_SHEET = 'SS_Employee_Roster';
    const HEADER_ROW = 2;
    const DATA_START_ROW = HEADER_ROW + 1;
    const HEADERS = ['Employee', 'Department Description', 'Start Date', 'Termination Date'];
    const CONFIG_FIELD = 'Headcount Last Updated';
    const CONFIG_CATEGORY = 'Run Settings';
    const CONFIG_DESCRIPTION = 'Most recent roster update logged from Headcount Review.';
    const DEPARTMENT_SHEET = 'SS_Department_Review';

    let departmentOptions = [];
    let rosterOptions = [];
    let activePanel = null;

    const statusEl = document.getElementById('hc-status');
    const panelEl = document.getElementById('hc-panel');
    const panelBodyEl = document.getElementById('hc-panel-body');
    const panelTitleEl = document.getElementById('hc-panel-title');
    const panelEyebrowEl = document.getElementById('hc-panel-eyebrow');
    const panelSaveBtn = document.getElementById('hc-panel-save');

    Office.onReady(() => init()).catch(() => init());

    async function init() {
        wireBaseEvents();
        setStatus('Loading headcount context…');
        await refreshReferenceData();
        setStatus('Ready.');
    }

    function wireBaseEvents() {
        document.querySelectorAll('.pf-action-tile[data-action]')?.forEach((btn) => {
            btn.addEventListener('click', () => openPanel(btn.dataset.action));
        });
        document.getElementById('hc-open-headcount')?.addEventListener('click', () => openWorkbookSheet(HEADCOUNT_SHEET, true));
        document.getElementById('hc-open-selector')?.addEventListener('click', () => {
            window.location.href = '../module-selector/index.html';
        });
        document.getElementById('hc-panel-close')?.addEventListener('click', closePanel);
        document.getElementById('hc-panel-cancel')?.addEventListener('click', closePanel);
        panelSaveBtn?.addEventListener('click', handlePanelSave);
        panelBodyEl?.addEventListener('click', handleNewRowRemoval);
    }

    async function refreshReferenceData() {
        const [dept, roster] = await Promise.all([loadDepartmentOptions(), loadRosterOptions()]);
        departmentOptions = dept;
        rosterOptions = roster;
        syncDatalists();
        await refreshLastUpdatedLabel();
    }

    async function refreshLastUpdatedLabel() {
        const lastUpdateEl = document.getElementById('hc-last-update');
        if (!lastUpdateEl) return;
        lastUpdateEl.textContent = 'Loading…';
        const value = await getLastUpdatedValue();
        lastUpdateEl.textContent = value ? formatTimestamp(value) : 'Not recorded yet';
    }

    function syncDatalists() {
        const deptList = document.getElementById('hc-dept-datalist');
        const locationList = document.getElementById('hc-location-datalist');
        if (deptList) {
            deptList.innerHTML = departmentOptions
                .filter((opt) => opt.label)
                .map((opt) => `<option value="${escapeHtml(opt.label)}"></option>`)
                .join('');
        }
        if (locationList) {
            const uniqueLocations = [...new Set(departmentOptions.map((opt) => opt.location).filter(Boolean))];
            locationList.innerHTML = uniqueLocations.map((loc) => `<option value="${escapeHtml(loc)}"></option>`).join('');
        }
    }

    function openPanel(type) {
        activePanel = type;
        panelEl?.classList.add('visible');
        panelEl?.classList.remove('hidden');
        panelEl?.setAttribute('aria-hidden', 'false');
        switch (type) {
            case 'new':
                panelEyebrowEl.textContent = 'Additions';
                panelTitleEl.textContent = 'New employees';
                renderNewEmployeesForm();
                break;
            case 'term':
                panelEyebrowEl.textContent = 'Separations';
                panelTitleEl.textContent = 'Terminated employees';
                renderTerminationForm();
                break;
            case 'changes':
                panelEyebrowEl.textContent = 'Roster adjustments';
                panelTitleEl.textContent = 'Department / location changes';
                renderChangesForm();
                break;
            default:
                panelEyebrowEl.textContent = '';
                panelTitleEl.textContent = '';
                panelBodyEl.innerHTML = '';
        }
    }

    function closePanel() {
        activePanel = null;
        panelEl?.classList.remove('visible');
        panelEl?.classList.add('hidden');
        panelEl?.setAttribute('aria-hidden', 'true');
        panelBodyEl.innerHTML = '';
    }

    function renderNewEmployeesForm() {
        panelBodyEl.innerHTML = `
            <div id="hc-new-rows" class="hc-form-stack"></div>
            <button type="button" id="hc-add-row" class="secondary">Add another employee</button>
            <small class="hc-form-hint">Rows are written to the Headcount Review sheet starting on row ${DATA_START_ROW}.</small>
        `;
        addNewEmployeeRow();
        document.getElementById('hc-add-row')?.addEventListener('click', () => addNewEmployeeRow());
    }

    function addNewEmployeeRow() {
        const rowsContainer = document.getElementById('hc-new-rows');
        if (!rowsContainer) return;
        const row = document.createElement('div');
        row.className = 'hc-form-row hc-new-row';
        row.innerHTML = `
            <div class="hc-form-field">
                <label>Employee Name<span class="hc-form-hint"> *</span></label>
                <input type="text" data-field="name" placeholder="Jane Doe" required>
            </div>
            <div class="hc-form-inline">
                <div class="hc-form-field">
                    <label>Employee ID</label>
                    <input type="text" data-field="id" placeholder="E12345">
                </div>
                <div class="hc-form-field">
                    <label>Start Date<span class="hc-form-hint"> *</span></label>
                    <input type="date" data-field="startDate" required>
                </div>
            </div>
            <div class="hc-form-inline">
                <div class="hc-form-field">
                    <label>Department Description</label>
                    <input type="text" data-field="department" list="hc-dept-datalist" placeholder="Ops — Chicago">
                </div>
                <div class="hc-form-field">
                    <label>Location</label>
                    <input type="text" data-field="location" list="hc-location-datalist" placeholder="Chicago HQ">
                </div>
            </div>
            <button type="button" class="link" data-remove-row>Remove</button>
        `;
        rowsContainer.appendChild(row);
    }

    function handleNewRowRemoval(event) {
        const target = event.target;
        if (!target.matches('[data-remove-row]')) return;
        const rows = panelBodyEl.querySelectorAll('.hc-new-row');
        if (rows.length <= 1) {
            target.closest('.hc-new-row')?.querySelectorAll('input').forEach((input) => (input.value = ''));
            return;
        }
        target.closest('.hc-new-row')?.remove();
    }

    function renderTerminationForm() {
        panelBodyEl.innerHTML = `
            <div class="hc-form-row">
                <div class="hc-form-field">
                    <label>Employee<span class="hc-form-hint"> *</span></label>
                    <select id="hc-term-employee">
                        <option value="">Select an employee…</option>
                        ${rosterOptions.map((emp) => `<option value="${emp.rowIndex}">${escapeHtml(emp.employee)}</option>`).join('')}
                    </select>
                </div>
                <div class="hc-form-field">
                    <label>Termination Date<span class="hc-form-hint"> *</span></label>
                    <input type="date" id="hc-term-date">
                </div>
            </div>
        `;
    }

    function renderChangesForm() {
        panelBodyEl.innerHTML = `
            <div class="hc-form-row">
                <div class="hc-form-field">
                    <label>Employee<span class="hc-form-hint"> *</span></label>
                    <select id="hc-change-employee">
                        <option value="">Select an employee…</option>
                        ${rosterOptions.map((emp) => `<option value="${emp.rowIndex}">${escapeHtml(emp.employee)}</option>`).join('')}
                    </select>
                </div>
                <div class="hc-form-field">
                    <label>New Department Description<span class="hc-form-hint"> *</span></label>
                    <input type="text" id="hc-change-dept" list="hc-dept-datalist" placeholder="Ops — Chicago">
                </div>
                <div class="hc-form-field">
                    <label>New Location</label>
                    <input type="text" id="hc-change-location" list="hc-location-datalist" placeholder="Chicago HQ">
                </div>
            </div>
        `;
    }

    async function handlePanelSave() {
        if (!activePanel) return;
        try {
            setStatus('Saving changes…');
            if (activePanel === 'new') {
                const entries = collectNewEmployeeEntries();
                if (!entries.length) throw new Error('Enter at least one new employee.');
                await addNewEmployees(entries);
            } else if (activePanel === 'term') {
                const payload = collectTerminationEntry();
                await recordTermination(payload);
            } else if (activePanel === 'changes') {
                const payload = collectChangeEntry();
                await recordChange(payload);
            }
            await setLastUpdatedStamp();
            await refreshReferenceData();
            closePanel();
            setStatus('Changes saved.');
        } catch (error) {
            console.error('Headcount Review:', error);
            setStatus(error.message || 'Unable to complete the requested action.', 'error');
        }
    }

    function collectNewEmployeeEntries() {
        const rows = panelBodyEl.querySelectorAll('.hc-new-row');
        const entries = [];
        rows.forEach((row) => {
            const name = row.querySelector('[data-field="name"]')?.value.trim() || '';
            const id = row.querySelector('[data-field="id"]')?.value.trim() || '';
            const department = row.querySelector('[data-field="department"]')?.value.trim() || '';
            const location = row.querySelector('[data-field="location"]')?.value.trim() || '';
            const startDate = row.querySelector('[data-field="startDate"]')?.value || '';
            if (!name && !id) return;
            if (!startDate) throw new Error('Start date is required for each new employee.');
            entries.push({
                employee: buildEmployeeLabel(name, id),
                department: combineDepartment(department, location),
                startDate
            });
        });
        return entries;
    }

    function collectTerminationEntry() {
        const employeeSelect = document.getElementById('hc-term-employee');
        const dateInput = document.getElementById('hc-term-date');
        if (!employeeSelect || !dateInput) throw new Error('Termination form unavailable.');
        const rowIndex = employeeSelect.value;
        const dateValue = dateInput.value;
        if (!rowIndex) throw new Error('Select an employee to mark as terminated.');
        if (!dateValue) throw new Error('Enter a termination date.');
        return { rowIndex: Number(rowIndex), terminationDate: dateValue };
    }

    function collectChangeEntry() {
        const employeeSelect = document.getElementById('hc-change-employee');
        const deptInput = document.getElementById('hc-change-dept');
        const locationInput = document.getElementById('hc-change-location');
        if (!employeeSelect || !deptInput) throw new Error('Change form unavailable.');
        const rowIndex = employeeSelect.value;
        if (!rowIndex) throw new Error('Select an employee to update.');
        const deptValue = deptInput.value.trim();
        if (!deptValue) throw new Error('Enter the new department description.');
        return {
            rowIndex: Number(rowIndex),
            department: combineDepartment(deptValue, locationInput?.value.trim() || '')
        };
    }

    async function addNewEmployees(entries) {
        await Excel.run(async (context) => {
            const sheet = await ensureHeadcountSheet(context);
            const columnRange = sheet.getRange(`A${DATA_START_ROW}:A1048576`).getUsedRangeOrNullObject();
            if (columnRange) columnRange.load('rowIndex,rowCount');
            await context.sync();
            let nextRow = DATA_START_ROW;
            if (columnRange && !columnRange.isNullObject) {
                nextRow = columnRange.rowIndex + columnRange.rowCount + 1;
            }
            const range = sheet.getRangeByIndexes(nextRow - 1, 0, entries.length, HEADERS.length);
            range.values = entries.map((entry) => [
                entry.employee,
                entry.department,
                entry.startDate ? new Date(entry.startDate) : '',
                ''
            ]);
        });
    }

    async function recordTermination({ rowIndex, terminationDate }) {
        await Excel.run(async (context) => {
            const sheet = await ensureHeadcountSheet(context);
            const cell = sheet.getCell(rowIndex, 3);
            cell.values = [[terminationDate ? new Date(terminationDate) : '']];
        });
    }

    async function recordChange({ rowIndex, department }) {
        await Excel.run(async (context) => {
            const sheet = await ensureHeadcountSheet(context);
            const cell = sheet.getCell(rowIndex, 1);
            cell.values = [[department]];
        });
    }

    async function openWorkbookSheet(sheetName, ensureStructure = false) {
        setStatus(`Opening ${sheetName}…`);
        try {
            let sheetMissing = false;
            await Excel.run(async (context) => {
                const worksheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                worksheet.load('name');
                await context.sync();
                if (worksheet.isNullObject) {
                    sheetMissing = true;
                    return;
                }
                if (ensureStructure && sheetName === HEADCOUNT_SHEET) {
                    await ensureHeadcountHeaderRow(context, worksheet);
                }
                worksheet.activate();
            });
            if (sheetMissing) {
                setStatus(
                    `${sheetName} is missing. Add the worksheet to the workbook and try again.`,
                    'error'
                );
                return;
            }
            setStatus(`${sheetName} is ready.`);
        } catch (error) {
            console.error('Unable to open worksheet:', error);
            setStatus(`Unable to open ${sheetName}.`, 'error');
        }
    }

    async function ensureHeadcountSheet(context) {
        let sheet = context.workbook.worksheets.getItemOrNullObject(HEADCOUNT_SHEET);
        sheet.load('name');
        await context.sync();
        if (sheet.isNullObject) {
            throw new Error(`${HEADCOUNT_SHEET} is missing. Add the worksheet to continue.`);
        }
        await ensureHeadcountHeaderRow(context, sheet);
        return sheet;
    }

    async function ensureHeadcountHeaderRow(context, sheet) {
        const headerRange = sheet.getRange(`A${HEADER_ROW}:D${HEADER_ROW}`);
        headerRange.load('values');
        await context.sync();
        const existing = (headerRange.values?.[0] || []).map((value) => (value ?? '').toString().trim().toLowerCase());
        const expected = HEADERS.map((value) => value.toLowerCase());
        const needsUpdate = expected.some((value, index) => existing[index] !== value);
        if (needsUpdate) {
            headerRange.values = [HEADERS];
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = '#1f2937';
            headerRange.format.font.color = '#ffffff';
        }
    }

    async function loadRosterOptions() {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(HEADCOUNT_SHEET);
            sheet.load('name');
            await context.sync();
            if (sheet.isNullObject) return [];
            const range = sheet.getRange(`A${DATA_START_ROW}:A1048576`).getUsedRangeOrNullObject();
            if (range) range.load('values,rowIndex,rowCount');
            await context.sync();
            if (!range || range.isNullObject) return [];
            const employees = [];
            for (let i = 0; i < range.values.length; i += 1) {
                const value = range.values[i][0];
                if (!value) continue;
                employees.push({
                    employee: value,
                    rowIndex: range.rowIndex + i
                });
            }
            return employees;
        });
    }

    async function loadDepartmentOptions() {
        return Excel.run(async (context) => {
            try {
                const sheet = context.workbook.worksheets.getItemOrNullObject(DEPARTMENT_SHEET);
                sheet.load('name');
                await context.sync();
                if (sheet.isNullObject) return [];
                const used = sheet.getUsedRangeOrNullObject();
                if (used) used.load('values');
                await context.sync();
                if (!used || used.isNullObject) return [];
                const rows = used.values || [];
                if (rows.length < 2) return [];
                const headers = rows[0].map((cell) => (cell ?? '').toString().trim().toLowerCase());
                const deptIdx = headers.findIndex((header) => header.includes('department'));
                const locIdx = headers.findIndex((header) => header.includes('location'));
                if (deptIdx === -1) return [];
                const options = [];
                for (let i = 1; i < rows.length; i += 1) {
                    const dept = rows[i][deptIdx];
                    if (!dept) continue;
                    const location = locIdx > -1 ? rows[i][locIdx] : '';
                    const label = combineDepartment(dept, location);
                    options.push({ label, department: dept, location: location || '' });
                }
                return options;
            } catch (error) {
                console.warn('SS_Department_Review sheet unavailable:', error);
                return [];
            }
        });
    }

    async function getLastUpdatedValue() {
        return Excel.run(async (context) => {
            try {
                const sheet = context.workbook.worksheets.getItemOrNullObject('SS_PF_Config');
                sheet.load('name');
                await context.sync();
                if (sheet.isNullObject) return null;
                const used = sheet.getUsedRangeOrNullObject();
                used.load('values');
                await context.sync();
                if (used.isNullObject) return null;
                const rows = used.values || [];
                if (rows.length < 2) return null;
                // Find Field and Value columns
                const headers = rows[0].map(h => (h ?? '').toString().trim().toLowerCase());
                const fieldIdx = headers.findIndex(h => h === 'field');
                const valueIdx = headers.findIndex(h => h === 'value');
                if (fieldIdx === -1 || valueIdx === -1) return null;
                const match = rows.slice(1).find((row) => (row[fieldIdx] ?? '').toString().trim() === CONFIG_FIELD);
                return match ? match[valueIdx] : null;
            } catch (error) {
                console.warn('Unable to read SS_PF_Config:', error);
                return null;
            }
        });
    }

    async function setLastUpdatedStamp() {
        const iso = new Date().toISOString();
        await Excel.run(async (context) => {
            try {
                const sheet = context.workbook.worksheets.getItemOrNullObject('SS_PF_Config');
                sheet.load('name');
                await context.sync();
                if (sheet.isNullObject) {
                    console.warn('SS_PF_Config sheet not found');
                    return;
                }
                const used = sheet.getUsedRangeOrNullObject();
                used.load('values,rowCount');
                await context.sync();
                if (used.isNullObject) return;
                const rows = used.values || [];
                if (rows.length < 1) return;
                // Find Field and Value columns
                const headers = rows[0].map(h => (h ?? '').toString().trim().toLowerCase());
                const fieldIdx = headers.findIndex(h => h === 'field');
                const valueIdx = headers.findIndex(h => h === 'value');
                const categoryIdx = headers.findIndex(h => h === 'category');
                if (fieldIdx === -1 || valueIdx === -1) return;
                const matchIndex = rows.slice(1).findIndex((row) => (row[fieldIdx] ?? '').toString().trim() === CONFIG_FIELD);
                if (matchIndex === -1) {
                    // Add new row at the end
                    const newRow = used.rowCount;
                    const newRowRange = sheet.getRangeByIndexes(newRow, 0, 1, headers.length);
                    const newValues = new Array(headers.length).fill('');
                    if (categoryIdx >= 0) newValues[categoryIdx] = CONFIG_CATEGORY;
                    newValues[fieldIdx] = CONFIG_FIELD;
                    newValues[valueIdx] = iso;
                    newRowRange.values = [newValues];
                } else {
                    const target = sheet.getCell(matchIndex + 1, valueIdx);
                    target.values = [[iso]];
                }
                await context.sync();
            } catch (error) {
                console.warn('Unable to write SS_PF_Config:', error);
            }
        });
        await refreshLastUpdatedLabel();
    }

    function buildEmployeeLabel(name, id) {
        if (id && name) return `${id} — ${name}`;
        return name || id || '';
    }

    function combineDepartment(department, location) {
        const dept = department || '';
        const loc = location || '';
        if (dept && loc) return `${dept} — ${loc}`;
        return dept || loc || '';
    }

    function formatTimestamp(value) {
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) return value;
        return date.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    }

    function setStatus(message, variant = 'info') {
        if (!statusEl) return;
        statusEl.textContent = message;
        statusEl.dataset.variant = variant;
    }

    function escapeHtml(value) {
        return (value ?? '').toString().replace(/[&<>"']/g, (char) => (
            { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]
        ));
    }
})();
