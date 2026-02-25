
// ==UserScript==
// @name Florence Basic Automator
// @namespace vinh.activity.plan.state
// @version 1.1.0
// @description
// @match https://us.v2.researchbinders.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/Florence%20Basic%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/Florence%20Basic%20Automator.js
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// ==/UserScript==

(function () {

    //==========================
    // TRAINING ELOG FUNCTIONS
    //==========================
    // This section contains functions to handle Training ELogs feature.
    // This includes a UI user input, parsing a list of names, and adding entries on webpages.
    //==========================

    const ELOG_SELECTORS = {
        mainTable: '.document-log-entries.document-log-entries__table',
        gridTable: '.document-log-entries__grid-table[role="table"]',
        row: 'log-entry-row[role="row"], .document-log-entries__grid-table__row[role="row"]',
        cell: '[role="cell"]',
        nameCellIndex: 3,
        namePrimary: '.u-text-overflow-ellipsis',
        nameFallback: '.test-logEntrySignature span'
    };

    const ELOG_TIMEOUTS = {
        waitTableMs: 10000,
        waitGridMs: 10000
    };

    const ELOG_CSS_CLASSNAMES = {
        panelOverlay: 'elog-panel-overlay',
        inputPanel: 'elog-input-panel',
        progressPanel: 'elog-progress-panel',
        warningPanel: 'elog-warning-panel',
        subpanelLeft: 'elog-subpanel-left',
        subpanelRight: 'elog-subpanel-right',
        searchInput: 'elog-search-input',
        listItem: 'elog-list-item',
        statusPending: 'elog-status-pending',
        statusFound: 'elog-status-found',
        statusNotFound: 'elog-status-notfound',
        statusDuplicate: 'elog-status-duplicate'
    };

    let elogState = {
        isRunning: false,
        observers: [],
        timeouts: [],
        intervals: [],
        eventListeners: [],
        parsedNames: [],
        normalizedNames: new Map(),
        scannedNames: [],
        focusReturnElement: null,
        abortController: null
    };

    function addELogStaffEntriesInit() {
        addLogMessage('addELogStaffEntriesInit: starting feature', 'log');
        elogState.focusReturnElement = document.getElementById('elog-staff-entries-btn');
        elogState.abortController = new AbortController();
        resetELogState();
        const mainTable = document.querySelector(ELOG_SELECTORS.mainTable);
        addLogMessage('addELogStaffEntriesInit: checking for main table selector', 'log');
        if (!mainTable) {
            addLogMessage('addELogStaffEntriesInit: main table not found, showing warning', 'warn');
            showELogWarning();
            return;
        }
        addLogMessage('addELogStaffEntriesInit: main table found, showing input panel', 'log');
        showELogInputPanel();
    }

    function resetELogState() {
        addLogMessage('resetELogState: resetting state', 'log');
        elogState.isRunning = false;
        elogState.parsedNames = [];
        elogState.normalizedNames = new Map();
        elogState.scannedNames = [];
    }

    function showELogWarning() {
        addLogMessage('showELogWarning: creating warning popup', 'log');
        const modal = document.createElement('div');
        modal.className = ELOG_CSS_CLASSNAMES.warningPanel;
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 30000; display: flex; align-items: center; justify-content: center;';
        const container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); border-radius: 12px; padding: 24px; width: 450px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'alertdialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'elog-warning-title');
        container.setAttribute('aria-describedby', 'elog-warning-message');
        const header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;';
        const title = document.createElement('h3');
        title.id = 'elog-warning-title';
        title.textContent = 'Document Log Not Found';
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.setAttribute('aria-label', 'Close warning');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.3)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        const closeWarning = function() {
            addLogMessage('showELogWarning: closing warning', 'log');
            if (modal.parentNode) { document.body.removeChild(modal); }
            stopELog();
            if (elogState.focusReturnElement) { elogState.focusReturnElement.focus(); }
        };
        closeButton.onclick = closeWarning;
        header.appendChild(title);
        header.appendChild(closeButton);
        const messageDiv = document.createElement('p');
        messageDiv.id = 'elog-warning-message';
        messageDiv.textContent = 'The current page does not contain the Document Log Entries table. Please navigate to a page with the Document Log Entries grid before using this feature.';
        messageDiv.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0; font-size: 14px; line-height: 1.5;';
        const okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: 2px solid rgba(255, 255, 255, 0.3); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.3s ease; margin-top: 20px; width: 100%;';
        okButton.onmouseover = function() { okButton.style.background = 'rgba(255, 255, 255, 0.3)'; };
        okButton.onmouseout = function() { okButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        okButton.onclick = closeWarning;
        const keyHandler = function(e) { if (e.key === 'Escape') { closeWarning(); } };
        document.addEventListener('keydown', keyHandler);
        elogState.eventListeners.push({ element: document, type: 'keydown', handler: keyHandler });
        container.appendChild(header);
        container.appendChild(messageDiv);
        container.appendChild(okButton);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);
        okButton.focus();
        addLogMessage('showELogWarning: warning displayed', 'log');
    }

    function showELogInputPanel() {
        addLogMessage('showELogInputPanel: creating input panel', 'log');
        const modal = document.createElement('div');
        modal.className = ELOG_CSS_CLASSNAMES.inputPanel;
        modal.id = 'elog-input-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        const container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 500px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'elog-input-title');
        const header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;';
        const title = document.createElement('h3');
        title.id = 'elog-input-title';
        title.textContent = 'Add ELog Staff Entries';
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600; letter-spacing: 0.2px;';
        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.setAttribute('aria-label', 'Close panel');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 67, 54, 0.8)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        closeButton.onclick = function() {
            addLogMessage('showELogInputPanel: closed by user', 'warn');
            if (modal.parentNode) { document.body.removeChild(modal); }
            stopELog();
            if (elogState.focusReturnElement) { elogState.focusReturnElement.focus(); }
        };
        header.appendChild(title);
        header.appendChild(closeButton);
        const description = document.createElement('p');
        description.textContent = 'Enter staff names to check against the Document Log. Separate names with commas or place each name on a new line.';
        description.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0 0 12px 0; font-size: 14px; line-height: 1.4;';
        const textarea = document.createElement('textarea');
        textarea.id = 'elog-names-input';
        textarea.placeholder = 'Name1, Name2, Name3\nor\nName1\nName2\nName3';
        textarea.setAttribute('aria-label', 'Staff names input');
        textarea.style.cssText = 'width: 100%; height: 160px; padding: 12px 14px; border: 2px solid rgba(255, 255, 255, 0.35); border-radius: 10px; background: rgba(255, 255, 255, 0.95); color: #1e293b; font-size: 14px; font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; resize: vertical; outline: none; transition: all 0.25s ease; box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset; box-sizing: border-box;';
        textarea.onfocus = function() { textarea.style.borderColor = '#8ea0ff'; textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)'; };
        textarea.onblur = function() { textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)'; textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset'; };
        const confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.disabled = true;
        confirmButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; opacity: 0.5;';
        const updateConfirmState = function() {
            const parsed = parseNamesInput(textarea.value);
            if (parsed.length > 0) { confirmButton.disabled = false; confirmButton.style.opacity = '1'; confirmButton.style.cursor = 'pointer'; }
            else { confirmButton.disabled = true; confirmButton.style.opacity = '0.5'; confirmButton.style.cursor = 'not-allowed'; }
        };
        textarea.oninput = updateConfirmState;
        confirmButton.onmouseover = function() { if (!confirmButton.disabled) { confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)'; } };
        confirmButton.onmouseout = function() { confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)'; };
        confirmButton.onclick = function() {
            addLogMessage('showELogInputPanel: Confirm clicked', 'log');
            const parsed = parseNamesInput(textarea.value);
            if (parsed.length === 0) { addLogMessage('showELogInputPanel: no valid names parsed', 'warn'); return; }
            elogState.parsedNames = parsed;
            addLogMessage('showELogInputPanel: parsed ' + parsed.length + ' unique names', 'log');
            if (modal.parentNode) { document.body.removeChild(modal); }
            showELogProgressPanel();
        };
        const clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease; backdrop-filter: blur(2px);';
        clearButton.onmouseover = function() { clearButton.style.background = 'rgba(255, 255, 255, 0.28)'; };
        clearButton.onmouseout = function() { clearButton.style.background = 'rgba(255, 255, 255, 0.18)'; };
        clearButton.onclick = function() {
            addLogMessage('showELogInputPanel: Clear All clicked', 'log');
            textarea.value = '';
            elogState.parsedNames = [];
            elogState.normalizedNames = new Map();
            confirmButton.disabled = true;
            confirmButton.style.opacity = '0.5';
            confirmButton.style.cursor = 'not-allowed';
        };
        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = 'display: flex; gap: 12px; margin-top: 20px; justify-content: flex-end;';
        buttonContainer.appendChild(clearButton);
        buttonContainer.appendChild(confirmButton);
        container.appendChild(header);
        container.appendChild(description);
        container.appendChild(textarea);
        container.appendChild(buttonContainer);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);
        textarea.focus();
        addLogMessage('showELogInputPanel: input panel displayed', 'log');
    }

    function parseNamesInput(input) {
        addLogMessage('parseNamesInput: parsing input', 'log');
        if (!input || !input.trim()) { addLogMessage('parseNamesInput: empty input', 'warn'); return []; }
        const results = [];
        const seenNormalized = new Set();
        elogState.normalizedNames = new Map();
        const lines = input.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const parts = line.split(',');
            for (let j = 0; j < parts.length; j++) {
                let name = parts[j].trim();
                name = name.replace(/,+$/, '').trim();
                if (!name) { continue; }
                const normalized = elogNormalizeName(name);
                if (!normalized) { continue; }
                if (seenNormalized.has(normalized)) { addLogMessage('parseNamesInput: duplicate detected (ignored): ' + name, 'log'); continue; }
                seenNormalized.add(normalized);
                elogState.normalizedNames.set(normalized, name);
                results.push({ display: name, normalized: normalized, status: 'Pending' });
            }
        }
        addLogMessage('parseNamesInput: parsed ' + results.length + ' unique names', 'log');
        return results;
    }

    function elogNormalizeName(name) {
        if (!name) { return ''; }
        let normalized = name.trim();
        normalized = normalized.replace(/\s+/g, ' ');
        normalized = normalized.toLowerCase();
        return normalized;
    }

    function showELogProgressPanel() {
        addLogMessage('showELogProgressPanel: creating progress panel', 'log');
        elogState.isRunning = true;
        const modal = document.createElement('div');
        modal.className = ELOG_CSS_CLASSNAMES.progressPanel;
        modal.id = 'elog-progress-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        const container = document.createElement('div');
        container.id = 'elog-progress-container';
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 900px; max-width: 95%; max-height: 80vh; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative; display: flex; flex-direction: column;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'elog-progress-title');
        const header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0;';
        const titleContainer = document.createElement('div');
        titleContainer.style.cssText = 'display: flex; align-items: center; gap: 12px;';
        const title = document.createElement('h3');
        title.id = 'elog-progress-title';
        title.textContent = 'ELog Staff Entries - Scanning';
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
        const statusBadge = document.createElement('span');
        statusBadge.id = 'elog-status-badge';
        statusBadge.textContent = 'In Progress';
        statusBadge.style.cssText = 'background: rgba(255, 255, 255, 0.3); color: #ffd93d; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
        titleContainer.appendChild(title);
        titleContainer.appendChild(statusBadge);
        const headerButtons = document.createElement('div');
        headerButtons.style.cssText = 'display: flex; gap: 8px; align-items: center;';
        const rescanButton = document.createElement('button');
        rescanButton.textContent = 'Re-scan';
        rescanButton.id = 'elog-rescan-btn';
        rescanButton.setAttribute('aria-label', 'Re-scan document log');
        rescanButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: 2px solid rgba(255, 255, 255, 0.3); color: white; padding: 6px 14px; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.25s ease;';
        rescanButton.onmouseover = function() { rescanButton.style.background = 'rgba(255, 255, 255, 0.3)'; };
        rescanButton.onmouseout = function() { rescanButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        rescanButton.onclick = function() { addLogMessage('showELogProgressPanel: Re-scan clicked', 'log'); performRescan(); };
        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.setAttribute('aria-label', 'Close and stop scanning');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 67, 54, 0.8)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        closeButton.onclick = function() {
            addLogMessage('showELogProgressPanel: closed by user', 'warn');
            if (modal.parentNode) { document.body.removeChild(modal); }
            stopELog();
            if (elogState.focusReturnElement) { elogState.focusReturnElement.focus(); }
        };
        headerButtons.appendChild(rescanButton);
        headerButtons.appendChild(closeButton);
        header.appendChild(titleContainer);
        header.appendChild(headerButtons);
        const panelsContainer = document.createElement('div');
        panelsContainer.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 16px; flex: 1; min-height: 0; overflow: hidden;';
        const leftPanel = createSubpanel('Scanned Log Entries', 'elog-left-panel', 'elog-left-search');
        const rightPanel = createSubpanel('User Names Status', 'elog-right-panel', 'elog-right-search');
        panelsContainer.appendChild(leftPanel);
        panelsContainer.appendChild(rightPanel);
        container.appendChild(header);
        container.appendChild(panelsContainer);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);
        initializeRightPanel();
        rescanButton.focus();
        addLogMessage('showELogProgressPanel: progress panel displayed, starting scan', 'log');
        startELogScan();
    }

    function createSubpanel(titleText, listId, searchId) {
        addLogMessage('createSubpanel: creating ' + titleText, 'log');
        const panel = document.createElement('div');
        panel.style.cssText = 'background: rgba(0, 0, 0, 0.2); border-radius: 8px; padding: 12px; display: flex; flex-direction: column; min-height: 0; overflow: hidden;';
        const panelHeader = document.createElement('div');
        panelHeader.style.cssText = 'margin-bottom: 10px; flex-shrink: 0;';
        const panelTitle = document.createElement('h4');
        panelTitle.textContent = titleText;
        panelTitle.style.cssText = 'margin: 0 0 8px 0; color: white; font-size: 14px; font-weight: 600;';
        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.id = searchId;
        searchInput.placeholder = 'Search...';
        searchInput.setAttribute('aria-label', 'Search ' + titleText);
        searchInput.style.cssText = 'width: 100%; padding: 8px 12px; border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 6px; background: rgba(255, 255, 255, 0.1); color: white; font-size: 13px; outline: none; transition: all 0.2s ease; box-sizing: border-box;';
        searchInput.onfocus = function() { searchInput.style.borderColor = 'rgba(255, 255, 255, 0.4)'; searchInput.style.background = 'rgba(255, 255, 255, 0.15)'; };
        searchInput.onblur = function() { searchInput.style.borderColor = 'rgba(255, 255, 255, 0.2)'; searchInput.style.background = 'rgba(255, 255, 255, 0.1)'; };
        searchInput.oninput = function() { filterSubpanelList(listId, searchInput.value); };
        panelHeader.appendChild(panelTitle);
        panelHeader.appendChild(searchInput);
        const listContainer = document.createElement('div');
        listContainer.id = listId;
        listContainer.style.cssText = 'flex: 1; overflow-y: auto; min-height: 200px; max-height: 400px;';
        panel.appendChild(panelHeader);
        panel.appendChild(listContainer);
        return panel;
    }

    function filterSubpanelList(listId, searchTerm) {
        addLogMessage('filterSubpanelList: filtering ' + listId + ' with term: ' + searchTerm, 'log');
        const list = document.getElementById(listId);
        if (!list) { return; }
        const items = list.querySelectorAll('.' + ELOG_CSS_CLASSNAMES.listItem);
        const searchLower = searchTerm.toLowerCase().trim();
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const text = item.textContent.toLowerCase();
            if (!searchLower || text.indexOf(searchLower) !== -1) { item.style.display = 'flex'; }
            else { item.style.display = 'none'; }
        }
    }

    function initializeRightPanel() {
        addLogMessage('initializeRightPanel: initializing with parsed names', 'log');
        const rightPanel = document.getElementById('elog-right-panel');
        if (!rightPanel) { addLogMessage('initializeRightPanel: right panel not found', 'error'); return; }
        rightPanel.innerHTML = '';
        for (let i = 0; i < elogState.parsedNames.length; i++) {
            const nameObj = elogState.parsedNames[i];
            const item = createListItem(nameObj.display, 'Pending', 'pending', i + 1);
            item.setAttribute('data-normalized', nameObj.normalized);
            rightPanel.appendChild(item);
        }
        addLogMessage('initializeRightPanel: added ' + elogState.parsedNames.length + ' items', 'log');
    }

    function createListItem(text, statusText, statusType, index) {
        const item = document.createElement('div');
        item.className = ELOG_CSS_CLASSNAMES.listItem;
        item.style.cssText = 'display: flex; justify-content: space-between; align-items: center; padding: 8px 10px; margin: 4px 0; background: rgba(255, 255, 255, 0.08); border-radius: 6px; transition: background 0.2s ease;';
        item.onmouseover = function() { item.style.background = 'rgba(255, 255, 255, 0.12)'; };
        item.onmouseout = function() { item.style.background = 'rgba(255, 255, 255, 0.08)'; };
        const leftSection = document.createElement('div');
        leftSection.style.cssText = 'display: flex; align-items: center; gap: 8px; flex: 1; min-width: 0;';
        if (index !== null && index !== undefined) {
            const indexBadge = document.createElement('span');
            indexBadge.textContent = index;
            indexBadge.style.cssText = 'background: rgba(255, 255, 255, 0.15); color: rgba(255, 255, 255, 0.7); font-size: 11px; font-weight: 600; min-width: 24px; height: 24px; border-radius: 12px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;';
            leftSection.appendChild(indexBadge);
        }
        const nameText = document.createElement('span');
        nameText.textContent = text;
        nameText.style.cssText = 'color: white; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;';
        leftSection.appendChild(nameText);
        item.appendChild(leftSection);
        if (statusText) {
            const statusBadge = document.createElement('span');
            statusBadge.className = 'elog-status-badge';
            statusBadge.textContent = statusText;
            let badgeColor = 'rgba(255, 255, 255, 0.7)';
            let badgeBg = 'rgba(255, 255, 255, 0.1)';
            if (statusType === 'pending') { badgeColor = '#ffd93d'; badgeBg = 'rgba(255, 217, 61, 0.2)'; }
            else if (statusType === 'found') { badgeColor = '#6bcf7f'; badgeBg = 'rgba(107, 207, 127, 0.2)'; }
            else if (statusType === 'notfound') { badgeColor = '#ff6b6b'; badgeBg = 'rgba(255, 107, 107, 0.2)'; }
            else if (statusType === 'duplicate') { badgeColor = '#ffa500'; badgeBg = 'rgba(255, 165, 0, 0.2)'; }
            statusBadge.style.cssText = 'color: ' + badgeColor + '; background: ' + badgeBg + '; font-size: 11px; font-weight: 600; padding: 3px 8px; border-radius: 10px; white-space: nowrap; flex-shrink: 0;';
            item.appendChild(statusBadge);
        }
        return item;
    }

    function startELogScan() {
        addLogMessage('startELogScan: beginning scan', 'log');
        elogState.scannedNames = [];
        waitForElement(ELOG_SELECTORS.mainTable, ELOG_TIMEOUTS.waitTableMs)
            .then(function(mainTable) {
                addLogMessage('startELogScan: main table found', 'log');
                return waitForElement(ELOG_SELECTORS.gridTable, ELOG_TIMEOUTS.waitGridMs);
            })
            .then(function(gridTable) {
                addLogMessage('startELogScan: grid table found, scanning rows', 'log');
                scanExistingStaffNames();
            })
            .catch(function(error) {
                addLogMessage('startELogScan: error during scan: ' + error, 'error');
                updateScanStatus('Error', 'error');
                showInlineNotice('An error occurred during scanning. The table may not be fully loaded.');
            });
    }

    function waitForElement(selector, timeout) {
        addLogMessage('waitForElement: waiting for ' + selector, 'log');
        return new Promise(function(resolve, reject) {
            const element = document.querySelector(selector);
            if (element) { addLogMessage('waitForElement: element found immediately', 'log'); resolve(element); return; }
            const observer = new MutationObserver(function(mutations, obs) {
                const el = document.querySelector(selector);
                if (el) {
                    addLogMessage('waitForElement: element found via observer', 'log');
                    obs.disconnect();
                    const idx = elogState.observers.indexOf(obs);
                    if (idx > -1) { elogState.observers.splice(idx, 1); }
                    resolve(el);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            elogState.observers.push(observer);
            const timeoutId = setTimeout(function() {
                observer.disconnect();
                const idx = elogState.observers.indexOf(observer);
                if (idx > -1) { elogState.observers.splice(idx, 1); }
                addLogMessage('waitForElement: timeout waiting for ' + selector, 'warn');
                reject(new Error('Timeout waiting for ' + selector));
            }, timeout);
            elogState.timeouts.push(timeoutId);
        });
    }

    function scanExistingStaffNames() {
        addLogMessage('scanExistingStaffNames: starting row iteration', 'log');
        if (!elogState.isRunning) { addLogMessage('scanExistingStaffNames: aborted, not running', 'warn'); return; }
        try {
            const rows = document.querySelectorAll(ELOG_SELECTORS.row);
            addLogMessage('scanExistingStaffNames: found ' + rows.length + ' rows', 'log');
            let rowIndex = 0;
            const processedNames = new Set();
            const processNextRow = function() {
                if (!elogState.isRunning) { addLogMessage('scanExistingStaffNames: stopped during processing', 'warn'); return; }
                if (rowIndex >= rows.length) { addLogMessage('scanExistingStaffNames: completed scanning all rows', 'log'); finalizeScan(); return; }
                const row = rows[rowIndex];
                const extractedName = extractNameFromRow(row, rowIndex + 1);
                if (extractedName && !processedNames.has(extractedName)) {
                    processedNames.add(extractedName);
                    elogState.scannedNames.push(extractedName);
                    updateLeftPanelList(extractedName, rowIndex + 1);
                    checkAndUpdateRightPanel(extractedName);
                }
                rowIndex++;
                const timeoutId = setTimeout(processNextRow, 10);
                elogState.timeouts.push(timeoutId);
            };
            processNextRow();
        } catch (error) {
            addLogMessage('scanExistingStaffNames: error: ' + error, 'error');
            showInlineNotice('Error scanning rows: ' + error.message);
        }
    }

    function extractNameFromRow(row, rowNumber) {
        addLogMessage('extractNameFromRow: processing row ' + rowNumber, 'log');
        try {
            const cells = row.querySelectorAll(ELOG_SELECTORS.cell);
            if (cells.length <= ELOG_SELECTORS.nameCellIndex) { addLogMessage('extractNameFromRow: not enough cells in row ' + rowNumber, 'warn'); return null; }
            const targetCell = cells[ELOG_SELECTORS.nameCellIndex];
            const primaryElement = targetCell.querySelector(ELOG_SELECTORS.namePrimary);
            if (primaryElement) {
                const brElement = primaryElement.querySelector('br');
                if (brElement) {
                    let nameText = '';
                    for (let i = 0; i < primaryElement.childNodes.length; i++) {
                        const node = primaryElement.childNodes[i];
                        if (node.nodeName === 'BR') { break; }
                        if (node.nodeType === Node.TEXT_NODE) { nameText += node.textContent; }
                    }
                    nameText = nameText.trim().replace(/\s+/g, ' ');
                    if (nameText) { addLogMessage('extractNameFromRow: extracted (primary): ' + nameText, 'log'); return nameText; }
                } else {
                    let nameText = primaryElement.textContent.trim().replace(/\s+/g, ' ');
                    if (nameText) { addLogMessage('extractNameFromRow: extracted (primary no br): ' + nameText, 'log'); return nameText; }
                }
            }
            const fallbackElement = targetCell.querySelector(ELOG_SELECTORS.nameFallback);
            if (fallbackElement) {
                let nameText = fallbackElement.textContent.trim().replace(/\s+/g, ' ');
                if (nameText) { addLogMessage('extractNameFromRow: extracted (fallback): ' + nameText, 'log'); return nameText; }
            }
            addLogMessage('extractNameFromRow: no name found in row ' + rowNumber, 'warn');
            return null;
        } catch (error) {
            addLogMessage('extractNameFromRow: error in row ' + rowNumber + ': ' + error, 'error');
            return null;
        }
    }

    function updateLeftPanelList(name, rowNumber) {
        addLogMessage('updateLeftPanelList: adding ' + name, 'log');
        const leftPanel = document.getElementById('elog-left-panel');
        if (!leftPanel) { addLogMessage('updateLeftPanelList: left panel not found', 'error'); return; }
        const item = createListItem(name, null, null, rowNumber);
        leftPanel.appendChild(item);
        const searchInput = document.getElementById('elog-left-search');
        if (searchInput && searchInput.value.trim()) { filterSubpanelList('elog-left-panel', searchInput.value); }
    }

    function checkAndUpdateRightPanel(scannedName) {
        addLogMessage('checkAndUpdateRightPanel: checking ' + scannedName, 'log');
        const normalizedScanned = elogNormalizeName(scannedName);
        for (let i = 0; i < elogState.parsedNames.length; i++) {
            const nameObj = elogState.parsedNames[i];
            if (nameObj.normalized === normalizedScanned && nameObj.status === 'Pending') {
                nameObj.status = 'Found';
                updateRightPanelItemStatus(nameObj.normalized, 'Found', 'found');
                addLogMessage('checkAndUpdateRightPanel: match found for ' + scannedName, 'log');
            }
        }
    }

    function updateRightPanelItemStatus(normalized, statusText, statusType) {
        addLogMessage('updateRightPanelItemStatus: updating ' + normalized + ' to ' + statusText, 'log');
        const rightPanel = document.getElementById('elog-right-panel');
        if (!rightPanel) { return; }
        const items = rightPanel.querySelectorAll('.' + ELOG_CSS_CLASSNAMES.listItem);
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            if (item.getAttribute('data-normalized') === normalized) {
                const badge = item.querySelector('.elog-status-badge');
                if (badge) {
                    badge.textContent = statusText;
                    let badgeColor = 'rgba(255, 255, 255, 0.7)';
                    let badgeBg = 'rgba(255, 255, 255, 0.1)';
                    if (statusType === 'found') { badgeColor = '#6bcf7f'; badgeBg = 'rgba(107, 207, 127, 0.2)'; }
                    else if (statusType === 'notfound') { badgeColor = '#ff6b6b'; badgeBg = 'rgba(255, 107, 107, 0.2)'; }
                    else if (statusType === 'pending') { badgeColor = '#ffd93d'; badgeBg = 'rgba(255, 217, 61, 0.2)'; }
                    badge.style.color = badgeColor;
                    badge.style.background = badgeBg;
                }
                break;
            }
        }
    }

    function finalizeScan() {
        addLogMessage('finalizeScan: marking remaining as Not Found', 'log');
        for (let i = 0; i < elogState.parsedNames.length; i++) {
            const nameObj = elogState.parsedNames[i];
            if (nameObj.status === 'Pending') {
                nameObj.status = 'Not Found';
                updateRightPanelItemStatus(nameObj.normalized, 'Not Found', 'notfound');
            }
        }
        updateScanStatus('Complete', 'complete');
        addLogMessage('finalizeScan: scan completed', 'log');
    }

    function updateScanStatus(statusText, statusType) {
        addLogMessage('updateScanStatus: ' + statusText, 'log');
        const badge = document.getElementById('elog-status-badge');
        const title = document.getElementById('elog-progress-title');
        if (badge) {
            badge.textContent = statusText;
            if (statusType === 'complete') { badge.style.background = 'rgba(107, 207, 127, 0.3)'; badge.style.color = '#6bcf7f'; }
            else if (statusType === 'error') { badge.style.background = 'rgba(255, 107, 107, 0.3)'; badge.style.color = '#ff6b6b'; }
            else { badge.style.background = 'rgba(255, 217, 61, 0.3)'; badge.style.color = '#ffd93d'; }
        }
        if (title && statusType === 'complete') { title.textContent = 'ELog Staff Entries - Complete'; }
    }

    function performRescan() {
        addLogMessage('performRescan: restarting scan', 'log');
        for (let i = 0; i < elogState.parsedNames.length; i++) {
            elogState.parsedNames[i].status = 'Pending';
            updateRightPanelItemStatus(elogState.parsedNames[i].normalized, 'Pending', 'pending');
        }
        const leftPanel = document.getElementById('elog-left-panel');
        if (leftPanel) { leftPanel.innerHTML = ''; }
        elogState.scannedNames = [];
        updateScanStatus('In Progress', 'progress');
        const title = document.getElementById('elog-progress-title');
        if (title) { title.textContent = 'ELog Staff Entries - Scanning'; }
        startELogScan();
    }

    function showInlineNotice(message) {
        addLogMessage('showInlineNotice: ' + message, 'warn');
        const container = document.getElementById('elog-progress-container');
        if (!container) { return; }
        const existingNotice = container.querySelector('.elog-inline-notice');
        if (existingNotice) { existingNotice.remove(); }
        const notice = document.createElement('div');
        notice.className = 'elog-inline-notice';
        notice.style.cssText = 'background: rgba(255, 193, 7, 0.2); border-left: 4px solid #ffc107; border-radius: 6px; padding: 10px 14px; margin-top: 12px; color: white; font-size: 13px; line-height: 1.4;';
        notice.textContent = message;
        container.appendChild(notice);
    }

    function stopELog() {
        addLogMessage('stopELog: stopping all ELog processes', 'log');
        elogState.isRunning = false;
        for (let i = 0; i < elogState.observers.length; i++) { try { elogState.observers[i].disconnect(); } catch (e) { addLogMessage('stopELog: error disconnecting observer: ' + e, 'error'); } }
        elogState.observers = [];
        for (let i = 0; i < elogState.timeouts.length; i++) { try { clearTimeout(elogState.timeouts[i]); } catch (e) { addLogMessage('stopELog: error clearing timeout: ' + e, 'error'); } }
        elogState.timeouts = [];
        for (let i = 0; i < elogState.intervals.length; i++) { try { clearInterval(elogState.intervals[i]); } catch (e) { addLogMessage('stopELog: error clearing interval: ' + e, 'error'); } }
        elogState.intervals = [];
        for (let i = 0; i < elogState.eventListeners.length; i++) { try { const listener = elogState.eventListeners[i]; listener.element.removeEventListener(listener.type, listener.handler); } catch (e) { addLogMessage('stopELog: error removing event listener: ' + e, 'error'); } }
        elogState.eventListeners = [];
        if (elogState.abortController) { elogState.abortController.abort(); elogState.abortController = null; }
        const inputModal = document.getElementById('elog-input-modal');
        if (inputModal && inputModal.parentNode) { inputModal.parentNode.removeChild(inputModal); }
        const progressModal = document.getElementById('elog-progress-modal');
        if (progressModal && progressModal.parentNode) { progressModal.parentNode.removeChild(progressModal); }
        resetELogState();
        addLogMessage('stopELog: cleanup complete', 'log');
    }

    //==========================
    // SHARED GUI AND PANEL FUNCTIONS
    //==========================
    // This section contains functions used by multiple features for panel management,
    // visibility control, hotkey handling, and UI interactions. These functions are
    // shared across all automation features and provide the common user interface.
    //==========================

    let guiVisible = localStorage.getItem('florence-gui-visible') === 'true';
    let guiScale = parseFloat(localStorage.getItem('florence-gui-scale')) || 1;
    let isDragging = false;
    let dragOffsetX = 0;
    let dragOffsetY = 0;
    let logMessages = [];
    let lastScrollPosition = 0;

    const originalLog = console.log;
    const originalError = console.error;
    const originalWarn = console.warn;

    function addLogMessage(message, type = 'log') {
        const timestamp = new Date().toLocaleTimeString();
        logMessages.push({ timestamp, message, type });
        updateLogBox();
    }

    console.log = function(...args) {
        originalLog.apply(console, args);
        addLogMessage(args.join(' '), 'log');
    };

    console.error = function(...args) {
        originalError.apply(console, args);
        addLogMessage(args.join(' '), 'error');
    };

    console.warn = function(...args) {
        originalWarn.apply(console, args);
        addLogMessage(args.join(' '), 'warn');
    };


    function showWarning(message) {
        const modal = document.createElement('div');
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            z-index: 30000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            border-radius: 12px;
            padding: 24px;
            width: 400px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Warning';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 255, 255, 0.3)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => document.body.removeChild(modal);

        header.appendChild(title);
        header.appendChild(closeButton);

        const messageDiv = document.createElement('p');
        messageDiv.textContent = message;
        messageDiv.style.cssText = `
            color: rgba(255, 255, 255, 0.9);
            margin: 0;
            font-size: 14px;
            line-height: 1.5;
        `;

        const okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            margin-top: 20px;
            width: 100%;
        `;
        okButton.onmouseover = () => okButton.style.background = 'rgba(255, 255, 255, 0.3)';
        okButton.onmouseout = () => okButton.style.background = 'rgba(255, 255, 255, 0.2)';
        okButton.onclick = () => document.body.removeChild(modal);

        container.appendChild(header);
        container.appendChild(messageDiv);
        container.appendChild(okButton);
        modal.appendChild(container);

        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);

        document.body.appendChild(modal);
    }
        function makeDraggable(container, handle) {
        let isDraggingModal = false;
        let offsetX = 0;
        let offsetY = 0;
        let currentScale = 1;

        handle.style.cursor = 'move';

        handle.addEventListener('mousedown', function(e) {
            if (e.target.tagName === 'BUTTON' || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
            isDraggingModal = true;

            const transform = container.style.transform;
            if (transform && transform.includes('scale')) {
                const match = transform.match(/scale\(([\d.]+)\)/);
                if (match) {
                    currentScale = parseFloat(match[1]);
                }
            } else {
                currentScale = 1;
            }

            const rect = container.getBoundingClientRect();
            // Store the offset in screen coordinates (already accounts for scale)
            offsetX = e.clientX - rect.left;
            offsetY = e.clientY - rect.top;

            e.preventDefault();
        });

        document.addEventListener('mousemove', function(e) {
            if (!isDraggingModal) return;

            // Calculate new position - offset is already in screen space
            let newX = e.clientX - offsetX;
            let newY = e.clientY - offsetY;

            // Use getBoundingClientRect for actual visual dimensions (accounts for scale)
            const rect = container.getBoundingClientRect();
            const visualWidth = rect.width;
            const visualHeight = rect.height;

            const viewportWidth = window.innerWidth;
            const viewportHeight = window.innerHeight;

            // Clamp to viewport using visual dimensions
            if (newX < 0) {
                newX = 0;
            } else if (newX + visualWidth > viewportWidth) {
                newX = viewportWidth - visualWidth;
            }

            if (newY < 0) {
                newY = 0;
            } else if (newY + visualHeight > viewportHeight) {
                newY = viewportHeight - visualHeight;
            }

            container.style.left = newX + 'px';
            container.style.top = newY + 'px';
            container.style.right = 'auto';

            // Preserve scale transform
            if (currentScale !== 1) {
                container.style.transform = 'scale(' + currentScale + ')';
                container.style.transformOrigin = 'top left';
            } else {
                container.style.transform = 'none';
            }
        });

        document.addEventListener('mouseup', function() {
            isDraggingModal = false;
        });
    }
    function createGUI() {
        const guiContainer = document.createElement('div');
        guiContainer.id = 'florence-gui';
        guiContainer.style.cssText = `
        position: fixed;
        top: 100px;
        right: 100px;
        width: 350px;
        min-height: 400px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        z-index: 10000;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        display: none;
        flex-direction: column;
        overflow: hidden;
        transform-origin: top left;
    `;

        const header = document.createElement('div');
        header.style.cssText = `
        background: rgba(255, 255, 255, 0.1);
        padding: 12px 16px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.2);
        cursor: move;
        display: flex;
        justify-content: space-between;
        align-items: center;
    `;

        const title = document.createElement('h3');
        title.textContent = 'Florence Automator';
        title.style.cssText = `
        margin: 0;
        color: white;
        font-size: 16px;
        font-weight: 600;
    `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
        background: rgba(255, 255, 255, 0.2);
        border: none;
        color: white;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.3s ease;
    `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => toggleGUI();

        header.appendChild(title);
        header.appendChild(closeButton);

        const buttonsContainer = document.createElement('div');
        buttonsContainer.style.cssText = `
        padding: 16px;
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        background: rgba(255, 255, 255, 0.05);
    `;

        for (let i = 1; i <= 4; i++) {
            const button = document.createElement('button');
            if (i === 1) {
                button.textContent = 'Add Signatures';
            } else if (i === 2) {
                button.textContent = 'Add ELog Staff Entries';
                button.id = 'elog-staff-entries-btn';
            } else {
                button.textContent = `Placeholder ${i}`;
            }
            button.style.cssText = `
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 12px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        `;
            button.onmouseover = () => {
                button.style.transform = 'translateY(-2px)';
                button.style.boxShadow = '0 6px 20px rgba(0, 0, 0, 0.3)';
            };
            button.onmouseout = () => {
                button.style.transform = 'translateY(0)';
                button.style.boxShadow = '0 4px 15px rgba(0, 0, 0, 0.2)';
            };

            if (i === 1) {
                button.onclick = () => {
                    console.log('Add Signatures button clicked');
                    startAddSignaturesFlow();
                };
            } else if (i === 2) {
                button.onclick = () => {
                    console.log('Add ELog Staff Entries button clicked');
                    addELogStaffEntriesInit();
                };
            } else {
                button.onclick = () => console.log(`Placeholder ${i} clicked`);
            }

            buttonsContainer.appendChild(button);
        }

        const scaleContainer = document.createElement('div');
        scaleContainer.style.cssText = `
        padding: 12px 16px;
        background: rgba(255, 255, 255, 0.05);
        border-top: 1px solid rgba(255, 255, 255, 0.1);
    `;

        const scaleLabel = document.createElement('div');
        scaleLabel.textContent = `Scale: ${guiScale.toFixed(2)}x`;
        scaleLabel.style.cssText = `
        color: white;
        font-size: 12px;
        margin-bottom: 8px;
        font-weight: 500;
    `;

        const scaleSlider = document.createElement('input');
        scaleSlider.type = 'range';
        scaleSlider.min = '0.75';
        scaleSlider.max = '1';
        scaleSlider.step = '0.05';
        scaleSlider.value = guiScale;
        scaleSlider.style.cssText = `
        width: 100%;
        height: 6px;
        border-radius: 3px;
        background: rgba(255, 255, 255, 0.3);
        outline: none;
        -webkit-appearance: none;
    `;

        scaleSlider.oninput = (e) => {
            const newScale = parseFloat(e.target.value);
            localStorage.setItem('florence-gui-scale', newScale);
            scaleLabel.textContent = `Scale: ${newScale.toFixed(2)}x (refresh to apply)`;
        };

        scaleContainer.appendChild(scaleLabel);
        scaleContainer.appendChild(scaleSlider);

        const logBox = document.createElement('div');
        logBox.id = 'florence-log-box';
        logBox.style.cssText = `
        flex: 1;
        padding: 12px;
        background: rgba(0, 0, 0, 0.3);
        overflow-y: auto;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 11px;
        line-height: 1.4;
        max-height: 200px;
    `;

        guiContainer.appendChild(header);
        guiContainer.appendChild(buttonsContainer);
        guiContainer.appendChild(scaleContainer);
        guiContainer.appendChild(logBox);

        document.body.appendChild(guiContainer);
        makeDraggable(guiContainer, header);
        addLogMessage('Florence Automator GUI initialized', 'log');
        updateGUIScale();
    }

    function updateLogBox() {
        const logBox = document.getElementById('florence-log-box');
        if (!logBox) return;

        logBox.innerHTML = logMessages.slice(-50).map(msg => {
            const color = msg.type === 'error' ? '#ff6b6b' :
            msg.type === 'warn' ? '#ffd93d' : '#6bcf7f';
            return `<div style="color: ${color}; margin-bottom: 4px;">
                <span style="opacity: 0.7;">[${msg.timestamp}]</span> ${msg.message}
            </div>`;
        }).join('');

        logBox.scrollTop = logBox.scrollHeight;
    }

    function updateGUIScale() {
        const gui = document.getElementById('florence-gui');
        if (!gui) return;

        gui.style.transform = `scale(${guiScale})`;
        gui.style.transformOrigin = 'top left';
    }

    function toggleGUI() {
        const gui = document.getElementById('florence-gui');
        if (!gui) {
            createGUI();
            setTimeout(() => {
                const guiElement = document.getElementById('florence-gui');
                guiElement.style.display = 'flex';
                guiVisible = true;
                localStorage.setItem('florence-gui-visible', 'true');
            }, 100);
        } else {
            guiVisible = !guiVisible;
            gui.style.display = guiVisible ? 'flex' : 'none';
            localStorage.setItem('florence-gui-visible', guiVisible.toString());
        }
    }

    //==========================
    // ADD SIGNATURES FUNCTIONS
    //==========================
    // This section contains functions used by Add Signatures feature
    //==========================


    function openSignaturesInputGUI() {
        addLogMessage('openSignaturesInputGUI: opening input modal', 'log');
        const modal = document.createElement('div');
        modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.7);
        z-index: 20000;
        display: flex;
        align-items: center;
        justify-content: center;
    `;

        const container = document.createElement('div');
        container.style.cssText = `
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 24px;
        width: 450px;
        max-width: 90%;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
        position: relative;
    `;

        const header = document.createElement('div');
        header.style.cssText = `
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
    `;

        const title = document.createElement('h3');
        title.textContent = 'Add Signatures';
        title.style.cssText = `
        margin: 0;
        color: white;
        font-size: 18px;
        font-weight: 600;
        letter-spacing: 0.2px;
    `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
        background: rgba(255, 255, 255, 0.2);
        border: none;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.3s ease;
    `;
        closeButton.onmouseover = () => {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = () => {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: modal closed by user', 'warn');
            document.body.removeChild(modal);
        };

        header.appendChild(title);
        header.appendChild(closeButton);

        const description = document.createElement('p');
        description.textContent = 'Paste or type the full list of signers. Separate names with commas or place each name on a new line.';
        description.style.cssText = `
        color: rgba(255, 255, 255, 0.9);
        margin: 0 0 12px 0;
        font-size: 14px;
        line-height: 1.4;
    `;

        const textarea = document.createElement('textarea');
        textarea.placeholder = 'Name1, Name2, Name3\nor\nName1\nName2\nName3';
        textarea.style.cssText = `
        width: 100%;
        height: 140px;
        padding: 12px 14px;
        border: 2px solid rgba(255, 255, 255, 0.35);
        border-radius: 10px;
        background: rgba(255, 255, 255, 0.95);
        color: #1e293b;
        font-size: 14px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        resize: vertical;
        outline: none;
        transition: all 0.25s ease;
        box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset;
    `;
        textarea.onfocus = () => {
            textarea.style.borderColor = '#8ea0ff';
            textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)';
        };
        textarea.onblur = () => {
            textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)';
            textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset';
        };

        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = `
        display: flex;
        gap: 12px;
        margin-top: 20px;
        justify-content: flex-end;
    `;

        const clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = `
        background: rgba(255, 255, 255, 0.18);
        border: 2px solid rgba(255, 255, 255, 0.35);
        color: white;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
        transition: all 0.25s ease;
        backdrop-filter: blur(2px);
    `;
        clearButton.onmouseover = () => {
            clearButton.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        clearButton.onmouseout = () => {
            clearButton.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        clearButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: Clear All clicked', 'log');
            textarea.value = '';
        };

        const confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.style.cssText = `
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        border: 2px solid rgba(255, 255, 255, 0.35);
        color: white;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 0.2px;
        transition: all 0.25s ease;
    `;
        confirmButton.onmouseover = () => {
            confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
        };
        confirmButton.onmouseout = () => {
            confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        confirmButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: Confirm clicked', 'log');
            const names = parseNames(textarea.value);
            addLogMessage('openSignaturesInputGUI: parsed ' + names.length + ' name(s)', 'log');
            if (names.length === 0) {
                addLogMessage('openSignaturesInputGUI: no names entered, showing warning', 'warn');
                showWarning('Please enter at least one name.');
                return;
            }
            document.body.removeChild(modal);
            processSignatures(names);
        };

        buttonContainer.appendChild(clearButton);
        buttonContainer.appendChild(confirmButton);

        container.appendChild(header);
        container.appendChild(description);
        container.appendChild(textarea);
        container.appendChild(buttonContainer);

        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);

        document.body.appendChild(modal);

        textarea.focus();
    }

    function startAddSignaturesFlow() {
        addLogMessage('Start Add Signatures flow clicked', 'log');
        addLogMessage('Validating if user is on the Request Signatures page', 'log');
        const isOnSignaturePage = validateSignaturePage();
        if (isOnSignaturePage) {
            addLogMessage('User is on the Request Signatures page, opening input GUI', 'log');
            openSignaturesInputGUI();
        } else {
            addLogMessage('User is not on the Request Signatures page, showing warning first', 'warn');
            showWarning('Please navigate to the "Request Signatures" page first.');
        }
    }
    function parseNames(input) {
        addLogMessage('parseNames: start', 'log');
        if (!input || !input.trim()) {
            addLogMessage('parseNames: empty input', 'warn');
            return [];
        }

        const names = [];
        const lines = input.split('\n');
        addLogMessage('parseNames: lines detected = ' + lines.length, 'log');

        for (const line of lines) {
            const lineNames = line.split(',').map(name => name.trim()).filter(name => name);
            addLogMessage('parseNames: line parsed -> ' + JSON.stringify(lineNames), 'log');
            names.push(...lineNames);
        }

        addLogMessage('parseNames: total names parsed = ' + names.length, 'log');
        return names;
    }


    function validateSignaturePage() {
        addLogMessage('validateSignaturePage: checking DOM for Request Signatures UI', 'log');
        const modalContainer = document.querySelector('modal-container');
        const signatureRequests = document.querySelector('documents-signature-requests');
        const headerTitle = document.querySelector('.c-modal-with-tabs_header_title');

        let conditionA = false;
        if (modalContainer && signatureRequests) {
            conditionA = true;
        }

        let conditionB = false;
        if (headerTitle && headerTitle.textContent && headerTitle.textContent.trim() === 'Request Signatures') {
            conditionB = true;
        }

        addLogMessage('validateSignaturePage: conditionA(modal + signatureRequests)=' + conditionA + ', conditionB(headerTitle)=' + conditionB, 'log');
        const result = conditionA || conditionB;
        addLogMessage('validateSignaturePage: result=' + result, 'log');
        return result;
    }

    function processSignatures(names) {
        addLogMessage('processSignatures: start', 'log');
        if (!validateSignaturePage()) {
            addLogMessage('processSignatures: not on Request Signatures page, showing warning', 'warn');
            showWarning('Please navigate to the "Request Signatures" page first.');
            return;
        }

        addLogMessage('processSignatures: proceeding with ' + names.length + ' name(s)', 'log');
        showLoadingGUI(names);
        setTimeout(() => {
            addLogMessage('processSignatures: calling executeSignatureSelection after delay', 'log');
            executeSignatureSelection(names);
        }, 100);
    }

    function showLoadingGUI(names) {
        const modal = document.createElement('div');
        modal.id = 'signatures-loading-modal';
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.8);
            z-index: 40000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            padding: 24px;
            width: 500px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Processing Signatures';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => {
            window.signatureProcessStopped = true;
            document.body.removeChild(modal);
        };

        header.appendChild(title);
        header.appendChild(closeButton);

        const spinner = document.createElement('div');
        spinner.style.cssText = `
            width: 40px;
            height: 40px;
            border: 4px solid rgba(255, 255, 255, 0.3);
            border-top: 4px solid white;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        `;

        const style = document.createElement('style');
        style.textContent = `
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);

        const statusContainer = document.createElement('div');
        statusContainer.id = 'signature-status-container';
        statusContainer.style.cssText = `
            max-height: 300px;
            overflow-y: auto;
            background: rgba(0, 0, 0, 0.2);
            border-radius: 8px;
            padding: 12px;
        `;

        container.appendChild(header);
        container.appendChild(spinner);

        const warningMessage = document.createElement('div');
        warningMessage.style.cssText = `
            background: rgba(255, 193, 7, 0.2);
            border-left: 4px solid #ffc107;
            border-radius: 6px;
            padding: 12px 16px;
            margin-bottom: 16px;
            color: white;
            font-size: 13px;
            line-height: 1.5;
        `;
        warningMessage.innerHTML = `
            <strong style="display: block; margin-bottom: 4px; font-size: 14px;">⚠️ Important</strong>
            Please do not click anywhere on the page while the automation is running. Clicking may close the dropdown menu and interrupt the process.
        `;
        container.appendChild(warningMessage);
        container.appendChild(statusContainer);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);

        names.forEach(name => {
            const statusDiv = document.createElement('div');
            statusDiv.id = `status-${name.replace(/\s+/g, '-')}`;
            statusDiv.style.cssText = `
                color: white;
                padding: 8px;
                margin: 4px 0;
                background: rgba(255, 255, 255, 0.1);
                border-radius: 4px;
                font-size: 14px;
            `;
            statusDiv.innerHTML = `<strong>${name}:</strong> <span style="color: #ffd93d;">Processing...</span>`;
            statusContainer.appendChild(statusDiv);
        });
    }

    function updateSignatureStatus(name, status, color = '#6bcf7f') {
        addLogMessage('updateSignatureStatus: name="' + name + '" status="' + status + '" color="' + color + '"', 'log');
        const statusDiv = document.getElementById('status-' + name.replace(/\s+/g, '-'));
        if (statusDiv) {
            statusDiv.innerHTML = '<strong>' + name + ':</strong> <span style="color: ' + color + ';">' + status + '</span>';
        } else {
            addLogMessage('updateSignatureStatus: statusDiv not found for "' + name + '"', 'warn');
        }
    }

    function executeSignatureSelection(names) {
        addLogMessage('executeSignatureSelection: start', 'log');
        window.signatureProcessStopped = false;

        try {
            const signersTab = document.querySelector('li.nav-item.active a[role="tab"]');
            if (signersTab && signersTab.textContent && signersTab.textContent.includes('Signers')) {
                addLogMessage('executeSignatureSelection: active Signers tab found, clicking', 'log');
                signersTab.click();
                setTimeout(() => {
                    addLogMessage('executeSignatureSelection: proceeding to selectSigners', 'log');
                    selectSigners(names);
                }, 500);
            } else {
                addLogMessage('executeSignatureSelection: active Signers tab not found, scanning all tabs', 'warn');
                const allTabs = document.querySelectorAll('li.nav-item a[role="tab"]');
                addLogMessage('executeSignatureSelection: total tabs found = ' + allTabs.length, 'log');
                let clicked = false;
                for (const tab of allTabs) {
                    const label = (tab.textContent || '').trim();
                    addLogMessage('executeSignatureSelection: inspecting tab label="' + label + '"', 'log');
                    if (label.includes('Signers')) {
                        addLogMessage('executeSignatureSelection: Signers tab found, clicking', 'log');
                        tab.click();
                        clicked = true;
                        setTimeout(() => {
                            addLogMessage('executeSignatureSelection: proceeding to selectSigners after tab click', 'log');
                            selectSigners(names);
                        }, 500);
                        return;
                    }
                }
                if (!clicked) {
                    addLogMessage('executeSignatureSelection: Could not find Signers tab', 'error');
                    updateSignatureStatus('System', 'Could not find Signers tab', '#ff6b6b');
                }
            }
        } catch (error) {
            addLogMessage('executeSignatureSelection: Error navigating to Signers tab: ' + error, 'error');
            updateSignatureStatus('System', 'Navigation failed', '#ff6b6b');
        }
    }

    function selectSigners(names) {
        addLogMessage('selectSigners: start', 'log');
        if (window.signatureProcessStopped) {
            addLogMessage('selectSigners: process stopped flag detected, aborting', 'warn');
            return;
        }

        try {
            const searchInput = document.getElementById('filtered-select-input');
            if (searchInput) {
                addLogMessage('selectSigners: search input found, clicking to open list', 'log');
                searchInput.click();
                searchInput.focus();
                setTimeout(() => {
                    addLogMessage('selectSigners: calling processSignerList', 'log');
                    processSignerList(names);
                }, 400);
            } else {
                addLogMessage('selectSigners: Search input not found with id="filtered-select-input"', 'error');
                updateSignatureStatus('System', 'Search input not found', '#ff6b6b');
            }
        } catch (error) {
            addLogMessage('selectSigners: Error clicking search input: ' + error, 'error');
            updateSignatureStatus('System', 'Failed to open signer list', '#ff6b6b');
        }
    }
    function closeLoadingGUI() {
        try {
            const modal = document.getElementById('signatures-loading-modal');
            if (modal) {
                document.body.removeChild(modal);
                addLogMessage('closeLoadingGUI: modal closed successfully', 'log');
            } else {
                addLogMessage('closeLoadingGUI: modal not found', 'warn');
            }
        } catch (e) {
            addLogMessage('closeLoadingGUI: error closing modal: ' + e, 'error');
        }
    }

    function getExistingSignersFromTable(callback) {
        const existingNames = new Set();
        addLogMessage('getExistingSignersFromTable: starting to collect all existing signers', 'log');

        try {
            const table = document.querySelector('.documents-signers-tab__table');
            if (!table) {
                addLogMessage('getExistingSignersFromTable: table not found', 'warn');
                callback([]);
                return;
            }

            const viewport = table.querySelector('cdk-virtual-scroll-viewport');
            if (!viewport) {
                addLogMessage('getExistingSignersFromTable: viewport not found, collecting visible rows only', 'warn');
                const rows = table.querySelectorAll('[role="row"].documents-signers-tab__table__row');
                rows.forEach(row => {
                    const ariaLabel = row.getAttribute('aria-label');
                    if (ariaLabel) {
                        existingNames.add(ariaLabel.trim().toLowerCase());
                    }
                });
                callback(Array.from(existingNames));
                return;
            }

            let scrollPosition = 0;
            const SCROLL_STEP = 200;

            function scrollAndCollect() {
                viewport.scrollTop = scrollPosition;

                setTimeout(() => {
                    const rows = table.querySelectorAll('[role="row"].documents-signers-tab__table__row');
                    rows.forEach(row => {
                        const ariaLabel = row.getAttribute('aria-label');
                        if (ariaLabel) {
                            existingNames.add(ariaLabel.trim().toLowerCase());
                        }
                    });

                    addLogMessage('getExistingSignersFromTable: collected ' + existingNames.size + ' unique names at scroll position ' + scrollPosition, 'log');

                    if (scrollPosition >= viewport.scrollHeight - viewport.clientHeight) {
                        const finalNames = Array.from(existingNames);
                        addLogMessage('getExistingSignersFromTable: finished collecting ' + finalNames.length + ' total existing signers', 'log');
                        viewport.scrollTop = 0;
                        callback(finalNames);
                        return;
                    }

                    scrollPosition += SCROLL_STEP;
                    setTimeout(scrollAndCollect, 150);
                }, 150);
            }

            scrollAndCollect();

        } catch (e) {
            addLogMessage('getExistingSignersFromTable: error: ' + e, 'error');
            callback(Array.from(existingNames));
        }
    }

    function normalizeName(name) {
        if (!name) return '';

        let normalized = name
        .replace(/[`'"''""_-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

        addLogMessage('normalizeName: "' + name + '" -> "' + normalized + '"', 'log');
        return normalized;
    }

    function filterOutExistingSigners(names, callback) {
        getExistingSignersFromTable(function(existingSigners) {
            const filteredNames = [];
            const duplicates = [];
            const normalizedExistingSigners = existingSigners.map(s => normalizeName(s).toLowerCase());

            names.forEach(name => {
                const normalizedName = normalizeName(name).toLowerCase();

                if (normalizedExistingSigners.includes(normalizedName)) {
                    addLogMessage('filterOutExistingSigners: "' + name + '" already exists, skipping', 'log');
                    updateSignatureStatus(name, 'Already Exist', '#ffd93d');
                    duplicates.push({ name: name, status: 'Already Exist', statusType: 'duplicate' });
                } else {
                    filteredNames.push(name);
                }
            });

            addLogMessage('filterOutExistingSigners: ' + filteredNames.length + ' names to process after filtering', 'log');
            callback(filteredNames, duplicates);
        });
    }
    function processSignerList(names) {
        lastScrollPosition = 0;
        addLogMessage('processSignerList: start with ' + names.length + ' names to process', 'log');

        // Filter out existing signers first (async)
        filterOutExistingSigners(names, function(filteredNames, duplicates) {
            if (filteredNames.length === 0) {
                addLogMessage('processSignerList: no names to process after filtering', 'log');
                setTimeout(() => {
                    showCompletionSummary(names, duplicates || []);
                }, 500);
                return;
            }
            const processResults = duplicates || [];
            let index = 0;

            function processNext() {
                if (window.signatureProcessStopped) {
                    addLogMessage('processSignerList: stop flag detected, ending sequence', 'warn');
                    return;
                }

                if (index >= filteredNames.length) {
                    addLogMessage('processSignerList: all names processed, showing completion summary', 'log');
                    setTimeout(() => {
                        showCompletionSummary(names, processResults);
                    }, 500);
                    return;
                }

                const originalName = filteredNames[index];
                addLogMessage('processSignerList: processing "' + originalName + '" (index ' + (index + 1) + ' of ' + filteredNames.length + ')', 'log');

                const parts = splitNameParts(originalName);

                try {
                    attemptSelectByScrolling(parts, function(success, matchType) {
                        // Clear the search input before processing the next name
                        const searchInput = document.getElementById('filtered-select-input');
                        if (searchInput) {
                            clearSearchInput(searchInput);
                            // Re-open the dropdown for the next search
                            setTimeout(() => {
                                searchInput.click();
                                searchInput.focus();
                                // Restore scroll position after reopening dropdown
                                setTimeout(() => {
                                    const viewport = document.querySelector('cdk-virtual-scroll-viewport');
                                    if (viewport) {
                                        viewport.scrollTop = lastScrollPosition;
                                        addLogMessage('processNext: restored scroll position to ' + lastScrollPosition, 'log');
                                    }
                                }, 50);
                            }, 100);
                        }

                        if (success) {
                            updateSignatureStatus(originalName, matchType, '#6bcf7f');
                            processResults.push({ name: originalName, status: matchType, statusType: 'success' });
                            index = index + 1;
                            setTimeout(processNext, 500);
                            return;
                        } else {
                            addLogMessage('processSignerList: not found after scrolling search -> "' + originalName + '"', 'warn');
                            updateSignatureStatus(originalName, 'Not found', '#ff6b6b');
                            processResults.push({ name: originalName, status: 'Not found', statusType: 'failure' });
                            index = index + 1;
                            setTimeout(processNext, 500);
                            return;
                        }
                    });
                } catch (err) {
                    addLogMessage('processSignerList: error while processing "' + originalName + '": ' + err, 'error');
                    updateSignatureStatus(originalName, 'Processing failed', '#ff6b6b');
                    index = index + 1;
                    setTimeout(processNext, 400);
                    return;
                }
            }
            processNext();
        });
    }
    function attemptSelectByQuery(searchInput, query, parts, callback) {
        addLogMessage('attemptSelectByQuery: searching for "' + query + '"', 'log');

        // Clear the input first
        clearSearchInput(searchInput);

        // Type the search query
        typeIntoSearchInput(searchInput, query, function() {
            // Wait for the list to update with search results
            waitForListUpdate(query, 3000, 200, function(items) {
                addLogMessage('attemptSelectByQuery: list updated, searching for candidate', 'log');
                const candidate = findCandidateItem(items, parts);

                if (candidate) {
                    addLogMessage('attemptSelectByQuery: found candidate for "' + query + '"', 'log');
                    const clicked = clickCheckboxForItem(candidate, parts.full);
                    if (clicked) {
                        setTimeout(function() {
                            clickAddButtonIfEnabled();
                            callback(true, 'Selected by search');
                        }, 80);
                        return;
                    }
                } else {
                    addLogMessage('attemptSelectByQuery: no candidate found for "' + query + '"', 'warn');
                }
                callback(false, '');
            }, function() {
                addLogMessage('attemptSelectByQuery: timeout waiting for list update', 'warn');
                callback(false, '');
            });
        });
    }

    function attemptSelectByScrolling(parts, callback, retryCount = 0) {
        addLogMessage('attemptSelectByScrolling: searching for "' + parts.full + '" (attempt ' + (retryCount + 1) + ')', 'log');

        const viewport = document.querySelector('cdk-virtual-scroll-viewport');
        if (!viewport) {
            addLogMessage('attemptSelectByScrolling: viewport not found', 'error');
            callback(false, '');
            return;
        }

        // Start from the last scroll position, or reset to 0 if this is a retry
        let scrollPosition = retryCount > 0 ? 0 : lastScrollPosition;
        const SCROLL_STEP = 200;
        const MAX_SCROLLS = 50;
        let scrollCount = 0;
        let hasScrolledFullCycle = false;

        function scrollAndCheck() {
            if (window.signatureProcessStopped) {
                addLogMessage('attemptSelectByScrolling: process stopped', 'warn');
                callback(false, '');
                return;
            }

            viewport.scrollTop = scrollPosition;
            lastScrollPosition = scrollPosition;

            setTimeout(function() {
                try {
                    const items = document.querySelectorAll('.filtered-select__list__item');
                    addLogMessage('attemptSelectByScrolling: found ' + items.length + ' items at scroll position ' + scrollPosition, 'log');

                    if (items.length > 0) {
                        const candidate = findCandidateItem(items, parts);
                        if (candidate) {
                            addLogMessage('attemptSelectByScrolling: found candidate for "' + parts.full + '"', 'log');
                            const clicked = clickCheckboxForItem(candidate, parts.full);
                            if (clicked) {
                                setTimeout(function() {
                                    clickAddButtonIfEnabled();
                                    // Update the last scroll position to where we found the item
                                    lastScrollPosition = scrollPosition;
                                    addLogMessage('attemptSelectByScrolling: updated lastScrollPosition to ' + lastScrollPosition, 'log');
                                    callback(true, 'Selected');
                                }, 80);
                                return;
                            }
                        }
                    }

                    // Continue scrolling
                    scrollPosition += SCROLL_STEP;
                    scrollCount++;

                    // Check if we've scrolled past the bottom
                    if (scrollPosition >= viewport.scrollHeight - viewport.clientHeight) {
                        if (!hasScrolledFullCycle) {
                            addLogMessage('attemptSelectByScrolling: reached bottom, starting from top', 'log');
                            scrollPosition = 0;
                            hasScrolledFullCycle = true;
                        } else {
                            // Completed full cycle without finding the item
                            if (retryCount === 0) {
                                // First attempt failed, retry with reset scroll
                                addLogMessage('attemptSelectByScrolling: first attempt failed, retrying with reset scroll', 'warn');
                                setTimeout(() => {
                                    attemptSelectByScrolling(parts, callback, 1);
                                }, 500);
                                return;
                            } else {
                                // Second attempt also failed
                                addLogMessage('attemptSelectByScrolling: second attempt failed, moving to next name', 'warn');
                                callback(false, '');
                                return;
                            }
                        }
                    }

                    if (scrollCount >= MAX_SCROLLS) {
                        if (retryCount === 0) {
                            // First attempt failed, retry
                            addLogMessage('attemptSelectByScrolling: max scrolls reached on first attempt, retrying', 'warn');
                            setTimeout(() => {
                                attemptSelectByScrolling(parts, callback, 1);
                            }, 500);
                            return;
                        } else {
                            // Second attempt failed
                            addLogMessage('attemptSelectByScrolling: max scrolls reached on second attempt, moving to next name', 'warn');
                            callback(false, '');
                            return;
                        }
                    }

                    // Continue scrolling
                    setTimeout(scrollAndCheck, 50);

                } catch (err) {
                    addLogMessage('attemptSelectByScrolling: error during scroll: ' + err, 'error');
                    if (retryCount === 0) {
                        setTimeout(() => {
                            attemptSelectByScrolling(parts, callback, 1);
                        }, 500);
                    } else {
                        callback(false, '');
                    }
                }
            }, 100);
        }

        scrollAndCheck();
    }

    function splitNameParts(name) {
        const result = { full: name, first: '', last: '' };

        if (!name) {
            addLogMessage('splitNameParts: empty name received', 'warn');
            return result;
        }

        const trimmed = normalizeName(name);

        if (trimmed.indexOf(',') !== -1) {
            const parts = trimmed.split(',').map(x => x.trim()).filter(x => x);
            if (parts.length >= 2) {
                result.last = parts[0];
                result.first = parts[1].split(' ').filter(x => x)[0] || '';
            }
        } else {
            const tokens = trimmed.split(' ').filter(x => x);
            if (tokens.length >= 2) {
                result.first = tokens[0];
                result.last = tokens[tokens.length - 1];
            } else if (tokens.length === 1) {
                result.first = tokens[0];
                result.last = '';
            }
        }

        addLogMessage('splitNameParts: name="' + name + '" -> first="' + result.first + '" last="' + result.last + '"', 'log');
        return result;
    }
    function clearSearchInput(input) {
        try {
            input.focus();
            input.value = '';
            const ev1 = new InputEvent('input', { bubbles: true });
            input.dispatchEvent(ev1);
            const ev2 = new KeyboardEvent('keyup', { bubbles: true, key: 'Backspace' });
            input.dispatchEvent(ev2);
            const ev3 = new Event('change', { bubbles: true });
            input.dispatchEvent(ev3);
            addLogMessage('clearSearchInput: input cleared', 'log');
        } catch (e) {
            addLogMessage('clearSearchInput: error clearing input: ' + e, 'error');
        }
    }

    function waitForListUpdate(expectedFragment, timeoutMs, intervalMs, onSuccess, onTimeout) {
        const start = Date.now();

        function poll() {
            const snapshot = getListItemsSnapshot();
            const names = snapshot.itemsText;

            if (snapshot.items.length > 0) {
                let hasRelated = false;
                for (let i = 0; i < names.length; i++) {
                    const txt = names[i].toLowerCase();
                    if (txt.indexOf(expectedFragment.toLowerCase()) !== -1) {
                        hasRelated = true;
                        break;
                    }
                }

                if (hasRelated) {
                    addLogMessage('waitForListUpdate: related items detected for "' + expectedFragment + '"', 'log');
                    onSuccess(snapshot.items);
                    return;
                }
            }

            if ((Date.now() - start) >= timeoutMs) {
                onTimeout();
                return;
            }

            setTimeout(poll, intervalMs);
        }

        poll();
    }

    function typeIntoSearchInput(input, text, done) {
        try {
            input.focus();
            input.value = '';
            const evClear = new InputEvent('input', { bubbles: true });
            input.dispatchEvent(evClear);
            addLogMessage('typeIntoSearchInput: typing "' + text + '"', 'log');

            let i = 0;

            function typeNext() {
                if (i >= text.length) {
                    const evChange = new Event('change', { bubbles: true });
                    input.dispatchEvent(evChange);
                    addLogMessage('typeIntoSearchInput: finished typing "' + text + '"', 'log');
                    if (typeof done === 'function') {
                        done();
                    }
                    return;
                }

                const ch = text.charAt(i);
                const evKeyDown = new KeyboardEvent('keydown', { bubbles: true, key: ch });
                input.dispatchEvent(evKeyDown);

                input.value = input.value + ch;

                const evInput = new InputEvent('input', { bubbles: true, data: ch });
                input.dispatchEvent(evInput);

                const evKeyUp = new KeyboardEvent('keyup', { bubbles: true, key: ch });
                input.dispatchEvent(evKeyUp);

                i = i + 1;
                setTimeout(typeNext, 30);
            }

            typeNext();
        } catch (e) {
            addLogMessage('typeIntoSearchInput: error typing "' + text + '": ' + e, 'error');
            if (typeof done === 'function') {
                done();
            }
        }
    }
    function getListItemsSnapshot() {
        const container = document.querySelector('.filtered-select__list');
        let items = [];
        if (container) {
            items = Array.from(container.querySelectorAll('.filtered-select__list__item'));
        } else {
            items = Array.from(document.querySelectorAll('.filtered-select__list__item'));
        }
        const itemsText = items.map(getItemDisplayName);
        addLogMessage('getListItemsSnapshot: items now=' + items.length, 'log');
        return { items: items, itemsText: itemsText };
    }

    function getItemDisplayName(li) {
        if (!li) {
            return '';
        }

        const aria = li.getAttribute('aria-label') || '';
        if (aria && aria.trim().length > 0) {
            return aria.trim().replace(/\s+/g, ' ');
        }

        const span = li.querySelector('.filtered-select__list__item__text');
        if (span && span.textContent) {
            return span.textContent.trim().replace(/\s+/g, ' ');
        }

        return (li.textContent || '').trim().replace(/\s+/g, ' ');
    }
    function findCandidateItem(items, parts) {
        const firstLower = normalizeName(parts.first || '').toLowerCase();
        const lastLower = normalizeName(parts.last || '').toLowerCase();

        const candidates = [];
        for (let i = 0; i < items.length; i++) {
            const nameText = normalizeName(getItemDisplayName(items[i]));
            const lc = nameText.toLowerCase();

            if (firstLower && lc.indexOf(firstLower) !== -1) {
                candidates.push({ item: items[i], name: nameText });
            }
        }

        if (candidates.length === 1) {
            addLogMessage('findCandidateItem: single candidate on first name -> "' + candidates[0].name + '"', 'log');
            return candidates[0].item;
        }

        if (candidates.length > 1 && lastLower) {
            for (let j = 0; j < candidates.length; j++) {
                const n = candidates[j].name.toLowerCase();
                if (n.indexOf(lastLower) !== -1) {
                    addLogMessage('findCandidateItem: disambiguated by last name -> "' + candidates[j].name + '"', 'log');
                    return candidates[j].item;
                }
            }
        }

        if (!parts.first && lastLower) {
            for (let k = 0; k < items.length; k++) {
                const nt = getItemDisplayName(items[k]).toLowerCase();
                if (nt.indexOf(lastLower) !== -1) {
                    addLogMessage('findCandidateItem: fallback last name matched -> "' + getItemDisplayName(items[k]) + '"', 'log');
                    return items[k];
                }
            }
        }

        addLogMessage('findCandidateItem: no candidate matched with provided parts (first="' + parts.first + '" last="' + parts.last + '")', 'warn');
        return null;
    }

    function clickCheckboxForItem(li, displayTarget) {
        if (!li) {
            addLogMessage('clickCheckboxForItem: no list item provided', 'error');
            return false;
        }

        try {
            const checkbox = li.querySelector('[role="checkbox"]');
            if (!checkbox) {
                addLogMessage('clickCheckboxForItem: checkbox not found for "' + displayTarget + '"', 'error');
                return false;
            }

            const state = checkbox.getAttribute('aria-checked');
            addLogMessage('clickCheckboxForItem: current aria-checked="' + state + '" for "' + displayTarget + '"', 'log');

            if (state === 'false') {
                checkbox.click();
                addLogMessage('clickCheckboxForItem: checkbox clicked for "' + displayTarget + '"', 'log');
                return true;
            } else {
                addLogMessage('clickCheckboxForItem: already selected for "' + displayTarget + '"', 'warn');
                return true;
            }
        } catch (e) {
            addLogMessage('clickCheckboxForItem: error clicking checkbox for "' + displayTarget + '": ' + e, 'error');
            return false;
        }
    }

    function clickAddButtonIfEnabled() {
        try {
            const addBtn = document.querySelector('[data-test="add-button"]');
            if (!addBtn) {
                addLogMessage('clickAddButtonIfEnabled: Add button not found', 'warn');
                return false;
            }

            const disabledAttr = addBtn.getAttribute('disabled');
            if (disabledAttr === null) {
                addBtn.click();
                addLogMessage('clickAddButtonIfEnabled: Add button clicked', 'log');
                return true;
            } else {
                addLogMessage('clickAddButtonIfEnabled: Add button disabled', 'warn');
                return false;
            }
        } catch (e) {
            addLogMessage('clickAddButtonIfEnabled: error clicking Add button: ' + e, 'error');
            return false;
        }
    }
    function fuzzyMatch(name1, name2) {
        const n1 = name1.toLowerCase().replace(/\s+/g, '');
        const n2 = name2.toLowerCase().replace(/\s+/g, '');
        if (n1.includes(n2) || n2.includes(n1)) {
            addLogMessage('fuzzyMatch: substring match true ("' + n1 + '" vs "' + n2 + '")', 'log');
            return true;
        }
        const similarity = calculateSimilarity(n1, n2);
        const result = similarity > 0.7;
        addLogMessage('fuzzyMatch: similarity=' + similarity.toFixed(3) + ' threshold=0.7 result=' + result + ' ("' + n1 + '" vs "' + n2 + '")', 'log');
        return result;
    }

    function calculateSimilarity(str1, str2) {
        const longer = str1.length > str2.length ? str1 : str2;
        const shorter = str1.length > str2.length ? str2 : str1;

        if (longer.length === 0) return 1.0;

        const editDistance = levenshteinDistance(longer, shorter);
        return (longer.length - editDistance) / longer.length;
    }

    function levenshteinDistance(str1, str2) {
        const matrix = [];

        for (let i = 0; i <= str2.length; i++) {
            matrix[i] = [i];
        }

        for (let j = 0; j <= str1.length; j++) {
            matrix[0][j] = j;
        }

        for (let i = 1; i <= str2.length; i++) {
            for (let j = 1; j <= str1.length; j++) {
                if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }

        return matrix[str2.length][str1.length];
    }

    function swapNameFormat(name) {
        if (name.includes(',')) {
            const parts = name.split(',').map(part => part.trim());
            const swapped = parts[1] + ' ' + parts[0];
            addLogMessage('swapNameFormat: input="' + name + '" swapped="' + swapped + '"', 'log');
            return swapped;
        } else {
            const parts = name.split(' ');
            if (parts.length >= 2) {
                const swapped = parts[parts.length - 1] + ', ' + parts.slice(0, -1).join(' ');
                addLogMessage('swapNameFormat: input="' + name + '" swapped="' + swapped + '"', 'log');
                return swapped;
            }
        }
        addLogMessage('swapNameFormat: input="' + name + '" unchanged', 'log');
        return name;
    }

    function showCompletionSummary(names, results) {
        const modal = document.getElementById('signatures-loading-modal');
        if (modal) {
            document.body.removeChild(modal);
        }
        let successCount = 0;
        let duplicateCount = 0;
        let failureCount = 0;

        results.forEach(result => {
            if (result.statusType === 'success') {
                successCount++;
            } else if (result.statusType === 'duplicate') {
                duplicateCount++;
            } else if (result.statusType === 'failure') {
                failureCount++;
            }
        });

        const totalProcessed = names.length;

        const summaryModal = document.createElement('div');
        summaryModal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.8);
            z-index: 50000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            padding: 24px;
            width: 520px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        `;

        const titleContainer = document.createElement('div');
        titleContainer.style.cssText = `display: flex; align-items: center; gap: 10px;`;

        const checkIcon = document.createElement('span');
        checkIcon.innerHTML = '✓';
        checkIcon.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            width: 32px;
            height: 32px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            color: #6bcf7f;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Process Complete';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        titleContainer.appendChild(checkIcon);
        titleContainer.appendChild(title);

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => document.body.removeChild(summaryModal);

        header.appendChild(titleContainer);
        header.appendChild(closeButton);

        // Statistics Section
        const statsSection = document.createElement('div');
        statsSection.style.cssText = `
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            margin-bottom: 16px;
        `;

        const createStatBox = (label, value, bgColor) => {
            const box = document.createElement('div');
            box.style.cssText = `
                background: ${bgColor};
                border-radius: 8px;
                padding: 12px 8px;
                text-align: center;
            `;

            const valueEl = document.createElement('div');
            valueEl.textContent = value;
            valueEl.style.cssText = `
                font-size: 24px;
                font-weight: 700;
                color: white;
                line-height: 1;
            `;

            const labelEl = document.createElement('div');
            labelEl.textContent = label;
            labelEl.style.cssText = `
                font-size: 11px;
                color: rgba(255, 255, 255, 0.85);
                margin-top: 4px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            `;

            box.appendChild(valueEl);
            box.appendChild(labelEl);
            return box;
        };

        statsSection.appendChild(createStatBox('Total', totalProcessed, 'rgba(255, 255, 255, 0.15)'));
        statsSection.appendChild(createStatBox('Success', successCount, 'rgba(107, 207, 127, 0.3)'));
        statsSection.appendChild(createStatBox('Duplicate', duplicateCount, 'rgba(255, 217, 61, 0.3)'));
        statsSection.appendChild(createStatBox('Failed', failureCount, 'rgba(255, 107, 107, 0.3)'));

        // Results List Section
        const listHeader = document.createElement('div');
        listHeader.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        `;

        const listTitle = document.createElement('span');
        listTitle.textContent = 'Detailed Results';
        listTitle.style.cssText = `
            color: rgba(255, 255, 255, 0.9);
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        `;

        listHeader.appendChild(listTitle);

        const resultsContainer = document.createElement('div');
        resultsContainer.style.cssText = `
            background: rgba(0, 0, 0, 0.2);
            border-radius: 8px;
            padding: 8px;
            max-height: 250px;
            overflow-y: auto;
        `;

        // Custom scrollbar styling
        const scrollStyle = document.createElement('style');
        scrollStyle.textContent = `
            .completion-results-list::-webkit-scrollbar {
                width: 6px;
            }
            .completion-results-list::-webkit-scrollbar-track {
                background: rgba(0, 0, 0, 0.1);
                border-radius: 3px;
            }
            .completion-results-list::-webkit-scrollbar-thumb {
                background: rgba(255, 255, 255, 0.3);
                border-radius: 3px;
            }
            .completion-results-list::-webkit-scrollbar-thumb:hover {
                background: rgba(255, 255, 255, 0.5);
            }
        `;
        document.head.appendChild(scrollStyle);
        resultsContainer.className = 'completion-results-list';

        results.forEach((result, index) => {
            const row = document.createElement('div');
            row.style.cssText = `
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 10px 12px;
                margin: 4px 0;
                background: rgba(255, 255, 255, 0.08);
                border-radius: 6px;
                transition: background 0.2s ease;
            `;
            row.onmouseover = () => row.style.background = 'rgba(255, 255, 255, 0.12)';
            row.onmouseout = () => row.style.background = 'rgba(255, 255, 255, 0.08)';

            const nameSection = document.createElement('div');
            nameSection.style.cssText = `
                display: flex;
                align-items: center;
                gap: 10px;
                flex: 1;
                min-width: 0;
            `;

            const indexBadge = document.createElement('span');
            indexBadge.textContent = index + 1;
            indexBadge.style.cssText = `
                background: rgba(255, 255, 255, 0.15);
                color: rgba(255, 255, 255, 0.7);
                font-size: 11px;
                font-weight: 600;
                min-width: 24px;
                height: 24px;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
            `;

            const nameText = document.createElement('span');
            nameText.textContent = result.name;
            nameText.style.cssText = `
                color: white;
                font-size: 14px;
                font-weight: 500;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            `;

            nameSection.appendChild(indexBadge);
            nameSection.appendChild(nameText);

            const statusBadge = document.createElement('span');
            statusBadge.textContent = result.status;

            let badgeColor, badgeBg;
            switch (result.statusType) {
                case 'success':
                    badgeColor = '#6bcf7f';
                    badgeBg = 'rgba(107, 207, 127, 0.2)';
                    break;
                case 'duplicate':
                    badgeColor = '#ffd93d';
                    badgeBg = 'rgba(255, 217, 61, 0.2)';
                    break;
                case 'failure':
                    badgeColor = '#ff6b6b';
                    badgeBg = 'rgba(255, 107, 107, 0.2)';
                    break;
                default:
                    badgeColor = 'rgba(255, 255, 255, 0.7)';
                    badgeBg = 'rgba(255, 255, 255, 0.1)';
            }

            statusBadge.style.cssText = `
                color: ${badgeColor};
                background: ${badgeBg};
                font-size: 12px;
                font-weight: 600;
                padding: 4px 10px;
                border-radius: 12px;
                white-space: nowrap;
                flex-shrink: 0;
            `;

            row.appendChild(nameSection);
            row.appendChild(statusBadge);
            resultsContainer.appendChild(row);
        });

        // OK Button
        const okButton = document.createElement('button');
        okButton.textContent = 'Close';
        okButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 12px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
            margin-top: 16px;
            width: 100%;
            text-transform: uppercase;
            letter-spacing: 1px;
        `;
        okButton.onmouseover = () => {
            okButton.style.background = 'rgba(255, 255, 255, 0.3)';
            okButton.style.borderColor = 'rgba(255, 255, 255, 0.5)';
        };
        okButton.onmouseout = () => {
            okButton.style.background = 'rgba(255, 255, 255, 0.2)';
            okButton.style.borderColor = 'rgba(255, 255, 255, 0.3)';
        };
        okButton.onclick = () => document.body.removeChild(summaryModal);

        container.appendChild(header);
        container.appendChild(statsSection);
        container.appendChild(listHeader);
        container.appendChild(resultsContainer);
        container.appendChild(okButton);
        summaryModal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        summaryModal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(summaryModal);

        addLogMessage('showCompletionSummary: displayed summary - Total: ' + totalProcessed + ', Success: ' + successCount + ', Duplicates: ' + duplicateCount + ', Failed: ' + failureCount, 'log');
    }

    function init() {
        document.addEventListener('keydown', (e) => {
            if (e.key === 'F2') {
                e.preventDefault();
                toggleGUI();
            }
        });

        console.log('Florence Basic Automator loaded. Press F2 to toggle GUI.');
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
    if (guiVisible) {
        toggleGUI();
    }
})();