
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

    const ELOG_SCROLL = {
        stepPx: 1000,
        idleDelayMs: 30,
        settleDelayMs: 100,
        maxDurationMs: 120000,
        maxNoProgressIterations: 8,
        userScrollPauseMs: 800,
        viewportOverscanPx: 400,
        retryScanAttempts: 3,
        retryScanDelayMs: 80
    };

    const ELOG_ATTRS = {
        ariaBusyTarget: 'body',
        ariaBusyAttr: 'aria-busy'
    };

    const ELOG_LABELS = {
        progressComplete: 'Scan complete',
        progressNoMore: 'End of list reached',
        progressRescanning: 'Re-scanning',
        progressStopped: 'Scan stopped'
    };

    const ELOG_FORM_SELECTORS = {
        addEntryBtn: '.test-createLogEntryBtn',
        memberInput: '#filtered-select-input.filtered-select__input',
        listContainer: 'ul.filtered-select__list.u-z-index-1060, ul.filtered-select__list, cdk-virtual-scroll-viewport',
        virtualViewport: 'cdk-virtual-scroll-viewport, .cdk-virtual-scroll-viewport',
        optionItem: 'cdk-virtual-scroll-viewport li.filtered-select__list__item, cdk-virtual-scroll-viewport li, .filtered-select__list li, cdk-virtual-scroll-viewport [role="option"], .filtered-select__list [role="option"]',
        saveAndAddAnotherBtn: 'button.btn.btn-primary',
        modalOrFormRoot: '.modal.show, .document-log-entries, body'
    };

    const ELOG_FORM_TIMEOUTS = {
        waitOpenMs: 10000,
        waitListMs: 6000,
        waitOptionRenderMs: 3000,
        waitSaveAfterClickMs: 8000,
        scrollIdleMs: 120,
        settleMs: 250,
        waitFilterMs: 800,
        maxSelectDurationMs: 45000
    };

    const ELOG_FORM_RETRY = {
        openListRetries: 4,
        selectRetriesPerScroll: 2,
        maxScrollPasses: 8,
        saveRetries: 2
    };

    const ELOG_RUN_LABELS = {
        statusPending: 'Pending',
        statusAlready: 'Already Exist',
        statusAdded: 'Completed',
        statusNotInDropdown: 'Not In Dropdown',
        statusSelectionFailed: 'Selection Failed',
        statusSaveFailed: 'Save Failed',
        statusStopped: 'Stopped'
    };

    const DOA_SELECTORS = {
        addEntryBtn: '.test-createLogEntryBtn',
        memberInput: '#filtered-select-input.filtered-select__input[placeholder*="Team Member"]',
        listContainer: 'ul.filtered-select__list.u-z-index-1060, ul.filtered-select__list, cdk-virtual-scroll-viewport',
        virtualViewport: 'cdk-virtual-scroll-viewport, .cdk-virtual-scroll-viewport',
        optionItem: 'cdk-virtual-scroll-viewport li.filtered-select__list__item, cdk-virtual-scroll-viewport li, .filtered-select__list li, cdk-virtual-scroll-viewport [role="option"], .filtered-select__list [role="option"]',
        roleSearchInput: '#filtered-select-input.filtered-select__input[placeholder*="Search"]',
        roleOptionItem: '.filtered-select__list__item, [role="option"], .cdk-virtual-scroll-viewport .filtered-select__list__item',
        roleOptionText: '.filtered-select__list__item__text',
        tasksToggleBtn: 'button.dropdown-toggle.log-entry-form__select-options-dropdown-button',
        tasksMenu: 'ul.dropdown-menu.log-entry-form__select-options-dropdown',
        tasksItem: 'ul.dropdown-menu.log-entry-form__select-options-dropdown li',
        tasksItemCheckbox: 'ul.dropdown-menu.log-entry-form__select-options-dropdown li input[type="checkbox"]',
        mainGridTable: '.document-log-entries__grid-table[role="table"]',
        mainGridRow: 'log-entry-row[role="row"], .document-log-entries__grid-table__row[role="row"]',
        mainGridCell: '[role="cell"]',
        mainTableContainer: '.document-log-entries.document-log-entries__table',
        ariaLiveRegion: '.aria-live-region'
    };

    const DOA_TIMEOUTS = {
        waitOpenMs: 10000,
        waitListMs: 6000,
        waitOptionRenderMs: 3000,
        waitRoleListMs: 6000,
        settleMs: 250,
        waitFilterMs: 800,
        waitTasksMenuMs: 5000,
        waitAfterTasksToggleMs: 200,
        maxSelectDurationMs: 45000,
        scrollIdleMs: 140
    };

    const DOA_RETRY = {
        openListRetries: 4,
        selectRetriesPerScroll: 2,
        maxScrollPasses: 8
    };

    const DOA_LABELS = {
        featureButton: 'Add DoA Log Staff Entries',
        statusPending: 'Pending',
        statusDuplicate: 'Duplicate',
        statusAlready: 'Already Exist',
        statusAdded: 'Completed',
        statusNotInDropdown: 'Not In Dropdown',
        statusSelectionFailed: 'Selection Failed',
        statusSaveFailed: 'Save Failed',
        statusRoleNotFound: 'Role Not Found',
        statusTasksApplied: 'Completed',
        statusStopped: 'Stopped'
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
        abortController: null,
        seenNormalizedNames: new Set(),
        prevAriaBusy: null,
        scrollContainer: null,
        prevScrollTop: 0,
        userScrollHandler: null,
        userScrollPaused: false,
        idleCallbackId: null,
        leftPanelRowIndex: 0,
        lastAutoScrollTime: 0,
        addQueue: [],
        addQueueIndex: 0,
        existingPairs: new Set(),
        counters: { total: 0, added: 0, duplicates: 0, failures: 0, pending: 0 },
        listScrollTop: 0,
        isAddingEntries: false
    };

    let doaState = {
        isRunning: false,
        observers: [],
        timeouts: [],
        intervals: [],
        eventListeners: [],
        idleCallbackIds: [],
        parsedCandidates: [],
        scannedNames: [],
        focusReturnElement: null,
        abortController: null,
        seenNormalizedNames: new Set(),
        prevAriaBusy: null,
        scrollContainer: null,
        prevScrollTop: 0,
        userScrollHandler: null,
        userScrollPaused: false,
        idleCallbackId: null,
        leftPanelRowIndex: 0,
        lastAutoScrollTime: 0,
        addQueue: [],
        addQueueIndex: 0,
        existingPairs: new Set(),
        counters: { total: 0, added: 0, duplicates: 0, failures: 0, pending: 0 },
        listScrollTop: 0,
        roleListScrollTop: 0,
        isAddingEntries: false
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
        elogState.seenNormalizedNames = new Set();
        elogState.prevAriaBusy = null;
        elogState.scrollContainer = null;
        elogState.prevScrollTop = 0;
        elogState.userScrollHandler = null;
        elogState.userScrollPaused = false;
        elogState.idleCallbackId = null;
        elogState.leftPanelRowIndex = 0;
        elogState.lastAutoScrollTime = 0;
        elogState.addQueue = [];
        elogState.addQueueIndex = 0;
        elogState.existingPairs = new Set();
        elogState.counters = { total: 0, added: 0, duplicates: 0, failures: 0, pending: 0 };
        elogState.listScrollTop = 0;
        elogState.isAddingEntries = false;
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
            elogState.isRunning = true;
            showCollectingDataPanel('elog', 'Add Training Log Staff Entries');
            startELogScan();
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
        title.textContent = 'ELog Staff Entries - Review';
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
        const summaryFooter = document.createElement('div');
        summaryFooter.id = 'elog-summary-footer';
        summaryFooter.setAttribute('aria-label', 'Processing summary');
        summaryFooter.style.cssText = 'display: flex; justify-content: space-around; align-items: center; padding: 10px 16px; background: rgba(0, 0, 0, 0.2); border-radius: 8px; margin-top: 12px; flex-shrink: 0;';
        const summaryItems = [
            { id: 'elog-summary-total', label: 'Total', value: '0' },
            { id: 'elog-summary-added', label: 'Added', value: '0' },
            { id: 'elog-summary-duplicates', label: 'Duplicates', value: '0' },
            { id: 'elog-summary-failures', label: 'Failures', value: '0' },
            { id: 'elog-summary-pending', label: 'Pending', value: '0' },
            { id: 'elog-summary-percent', label: 'Progress', value: '0%' }
        ];
        for (let si = 0; si < summaryItems.length; si++) {
            const summaryItem = document.createElement('div');
            summaryItem.style.cssText = 'text-align: center;';
            const valSpan = document.createElement('span');
            valSpan.id = summaryItems[si].id;
            valSpan.textContent = summaryItems[si].value;
            valSpan.style.cssText = 'display: block; color: white; font-size: 16px; font-weight: 700;';
            const labelSpan = document.createElement('span');
            labelSpan.textContent = summaryItems[si].label;
            labelSpan.style.cssText = 'display: block; color: rgba(255, 255, 255, 0.6); font-size: 11px; font-weight: 500; margin-top: 2px;';
            summaryItem.appendChild(valSpan);
            summaryItem.appendChild(labelSpan);
            summaryFooter.appendChild(summaryItem);
        }
        const ariaLiveRegion = document.createElement('div');
        ariaLiveRegion.id = 'elog-aria-live';
        ariaLiveRegion.setAttribute('aria-live', 'polite');
        ariaLiveRegion.setAttribute('aria-atomic', 'true');
        ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        container.appendChild(header);
        container.appendChild(panelsContainer);
        container.appendChild(summaryFooter);
        container.appendChild(ariaLiveRegion);
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
        populateELogLeftPanel();
        finalizeELogRightPanelAfterScan();
        rescanButton.focus();
        addLogMessage('showELogProgressPanel: progress panel displayed with all data populated', 'log');
        updateScanStatus('Scan Complete', 'complete');
        var progressTitle = document.getElementById('elog-progress-title');
        if (progressTitle) {
            progressTitle.textContent = 'ELog Staff Entries - Adding Entries';
        }
        beginAddNonDuplicateELogEntries();
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
            item.setAttribute('data-pairkey', normalizeFirstLastPair(nameObj.display));
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

    function showCollectingDataPanel(featureId, featureName) {
        addLogMessage('showCollectingDataPanel: showing for ' + featureId, 'log');
        var existing = document.getElementById(featureId + '-collecting-modal');
        if (existing && existing.parentNode) {
            existing.parentNode.removeChild(existing);
        }
        var modal = document.createElement('div');
        modal.id = featureId + '-collecting-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 36px 48px; max-width: 420px; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); text-align: center;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', featureId + '-collecting-title');
        var spinner = document.createElement('div');
        spinner.style.cssText = 'width: 48px; height: 48px; border: 4px solid rgba(255, 255, 255, 0.2); border-top-color: white; border-radius: 50%; margin: 0 auto 20px; animation: collectingSpin 0.8s linear infinite;';
        var styleTag = document.getElementById('collecting-spinner-style');
        if (!styleTag) {
            styleTag = document.createElement('style');
            styleTag.id = 'collecting-spinner-style';
            styleTag.textContent = '@keyframes collectingSpin { to { transform: rotate(360deg); } }';
            document.head.appendChild(styleTag);
        }
        var title = document.createElement('h3');
        title.id = featureId + '-collecting-title';
        title.textContent = featureName || 'Please wait';
        title.style.cssText = 'margin: 0 0 8px 0; color: white; font-size: 18px; font-weight: 600;';
        var message = document.createElement('p');
        message.textContent = 'Please wait. Collecting data.';
        message.style.cssText = 'margin: 0; color: rgba(255, 255, 255, 0.85); font-size: 15px; font-weight: 400;';
        container.appendChild(spinner);
        container.appendChild(title);
        container.appendChild(message);
        modal.appendChild(container);
        document.body.appendChild(modal);
        addLogMessage('showCollectingDataPanel: displayed for ' + featureId, 'log');
    }

    function removeCollectingDataPanel(featureId) {
        addLogMessage('removeCollectingDataPanel: removing for ' + featureId, 'log');
        var modal = document.getElementById(featureId + '-collecting-modal');
        if (modal && modal.parentNode) {
            modal.parentNode.removeChild(modal);
        }
    }

    function populateELogLeftPanel() {
        addLogMessage('populateELogLeftPanel: populating with ' + elogState.scannedNames.length + ' names', 'log');
        var leftPanel = document.getElementById('elog-left-panel');
        if (!leftPanel) {
            addLogMessage('populateELogLeftPanel: left panel not found', 'error');
            return;
        }
        leftPanel.innerHTML = '';
        for (var i = 0; i < elogState.scannedNames.length; i++) {
            var item = createListItem(elogState.scannedNames[i], null, null, i + 1);
            leftPanel.appendChild(item);
        }
        addLogMessage('populateELogLeftPanel: populated ' + elogState.scannedNames.length + ' items', 'log');
    }

    function finalizeELogRightPanelAfterScan() {
        addLogMessage('finalizeELogRightPanelAfterScan: updating right panel statuses', 'log');
        var scannedNormSet = new Set();
        for (var si = 0; si < elogState.scannedNames.length; si++) {
            scannedNormSet.add(elogNormalizeName(elogState.scannedNames[si]));
        }
        for (var i = 0; i < elogState.parsedNames.length; i++) {
            var nameObj = elogState.parsedNames[i];
            if (scannedNormSet.has(nameObj.normalized)) {
                nameObj.status = 'Found';
                updateRightPanelItemStatus(nameObj.normalized, 'Found', 'found');
            } else {
                nameObj.status = 'Not Found';
                updateRightPanelItemStatus(nameObj.normalized, 'Not Found', 'notfound');
            }
        }
        addLogMessage('finalizeELogRightPanelAfterScan: completed', 'log');
    }

    function startELogScan() {
        addLogMessage('startELogScan: beginning scan with auto-scroll', 'log');
        elogState.scannedNames = [];
        elogState.seenNormalizedNames = new Set();
        waitForElement(ELOG_SELECTORS.mainTable, ELOG_TIMEOUTS.waitTableMs)
            .then(function(mainTable) {
            addLogMessage('startELogScan: main table found', 'log');
            return waitForElement(ELOG_SELECTORS.gridTable, ELOG_TIMEOUTS.waitGridMs);
        })
            .then(function(gridTable) {
            addLogMessage('startELogScan: grid table found, starting auto-scroll scan', 'log');
            autoScrollScan({
                onRow: function(name, normalized) {
                    addLogMessage('startELogScan: onRow callback for ' + name, 'log');
                },
                onProgress: function(data) {
                    addLogMessage('startELogScan: progress - scanned ' + data.scanned, 'log');
                },
                onDone: function(data) {
                    addLogMessage('startELogScan: done - total=' + data.total + ' reason=' + data.reason, 'log');
                    removeCollectingDataPanel('elog');
                    if (data.reason !== 'stopped' && elogState.isRunning) {
                        addLogMessage('startELogScan: scan complete, showing progress panel', 'log');
                        showELogProgressPanel();
                    }
                },
                onError: function(error) {
                    addLogMessage('startELogScan: error - ' + error.message, 'error');
                    removeCollectingDataPanel('elog');
                    showELogProgressPanel();
                    updateScanStatus('Error', 'error');
                    showInlineNotice('Error during auto-scroll scan: ' + error.message);
                }
            });
        })
            .catch(function(error) {
            addLogMessage('startELogScan: error during scan: ' + error, 'error');
            removeCollectingDataPanel('elog');
            showELogProgressPanel();
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
        if (!elogState.isRunning) { return; }
        try {
            const rows = document.querySelectorAll(ELOG_SELECTORS.row);
            let rowIndex = 0;
            const processedNames = new Set();
            var batchSize = 10;
            const processNextBatch = function() {
                if (!elogState.isRunning) { return; }
                if (rowIndex >= rows.length) { finalizeScan(); return; }
                var end = Math.min(rowIndex + batchSize, rows.length);
                for (var bi = rowIndex; bi < end; bi++) {
                    const row = rows[bi];
                    const extractedName = extractNameFromRow(row, bi + 1);
                    if (extractedName && !processedNames.has(extractedName)) {
                        processedNames.add(extractedName);
                        elogState.scannedNames.push(extractedName);
                        updateLeftPanelList(extractedName, bi + 1);
                        checkAndUpdateRightPanel(extractedName);
                    }
                }
                rowIndex = end;
                const timeoutId = setTimeout(processNextBatch, 0);
                elogState.timeouts.push(timeoutId);
            };
            processNextBatch();
        } catch (error) {
            addLogMessage('scanExistingStaffNames: error: ' + error, 'error');
            showInlineNotice('Error scanning rows: ' + error.message);
        }
    }

    function extractNameFromRow(row, rowNumber) {
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
                    if (nameText) { return nameText; }
                } else {
                    let nameText = primaryElement.textContent.trim().replace(/\s+/g, ' ');
                    if (nameText) { return nameText; }
                }
            }
            const fallbackElement = targetCell.querySelector(ELOG_SELECTORS.nameFallback);
            if (fallbackElement) {
                let nameText = fallbackElement.textContent.trim().replace(/\s+/g, ' ');
                if (nameText) { return nameText; }
            }
            return null;
        } catch (error) {
            addLogMessage('extractNameFromRow: error in row ' + rowNumber + ': ' + error, 'error');
            return null;
        }
    }

    function updateLeftPanelList(name, rowNumber) {
        const leftPanel = document.getElementById('elog-left-panel');
        if (!leftPanel) { addLogMessage('updateLeftPanelList: left panel not found', 'error'); return; }
        const item = createListItem(name, null, null, rowNumber);
        leftPanel.appendChild(item);
        const searchInput = document.getElementById('elog-left-search');
        if (searchInput && searchInput.value.trim()) { filterSubpanelList('elog-left-panel', searchInput.value); }
    }

    function checkAndUpdateRightPanel(scannedName) {
        const normalizedScanned = elogNormalizeName(scannedName);
        for (let i = 0; i < elogState.parsedNames.length; i++) {
            const nameObj = elogState.parsedNames[i];
            if (nameObj.normalized === normalizedScanned && nameObj.status === 'Pending') {
                nameObj.status = 'Found';
                updateRightPanelItemStatus(nameObj.normalized, 'Found', 'found');
            }
        }
    }

    function updateRightPanelItemStatus(normalized, statusText, statusType) {
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
        addLogMessage('performRescan: restarting scan with auto-scroll', 'log');
        for (let i = 0; i < elogState.observers.length; i++) {
            try {
                elogState.observers[i].disconnect();
            } catch (e) {
                addLogMessage('performRescan: error disconnecting observer: ' + e, 'error');
            }
        }
        elogState.observers = [];
        for (let i = 0; i < elogState.timeouts.length; i++) {
            try {
                clearTimeout(elogState.timeouts[i]);
            } catch (e) {
                addLogMessage('performRescan: error clearing timeout: ' + e, 'error');
            }
        }
        elogState.timeouts = [];
        if (elogState.idleCallbackId && typeof cancelIdleCallback === 'function') {
            cancelIdleCallback(elogState.idleCallbackId);
            elogState.idleCallbackId = null;
        }
        elogState.scannedNames = [];
        elogState.seenNormalizedNames = new Set();
        elogState.leftPanelRowIndex = 0;
        elogState.userScrollPaused = false;
        elogState.isAddingEntries = false;
        elogState.addQueue = [];
        elogState.addQueueIndex = 0;
        for (let pi = 0; pi < elogState.parsedNames.length; pi++) {
            elogState.parsedNames[pi].status = 'Pending';
        }
        var progressModal = document.getElementById('elog-progress-modal');
        if (progressModal && progressModal.parentNode) {
            progressModal.parentNode.removeChild(progressModal);
        }
        showCollectingDataPanel('elog', 'Add Training Log Staff Entries');
        addLogMessage('performRescan: starting auto-scroll scan', 'log');
        autoScrollScan({
            onRow: function(name, normalized) {
                addLogMessage('performRescan: onRow callback for ' + name, 'log');
            },
            onProgress: function(data) {
                addLogMessage('performRescan: progress - scanned ' + data.scanned, 'log');
            },
            onDone: function(data) {
                addLogMessage('performRescan: done - total=' + data.total + ' reason=' + data.reason, 'log');
                removeCollectingDataPanel('elog');
                if (data.reason !== 'stopped' && elogState.isRunning) {
                    addLogMessage('performRescan: re-scan complete, showing progress panel', 'log');
                    showELogProgressPanel();
                }
            },
            onError: function(error) {
                addLogMessage('performRescan: error - ' + error.message, 'error');
                removeCollectingDataPanel('elog');
                showELogProgressPanel();
                updateScanStatus('Error', 'error');
                showInlineNotice('Error during re-scan: ' + error.message);
            }
        });
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

    function getFirstAndLast(name) {
        if (!name || !name.trim()) {
            return ['', ''];
        }
        const suffixPattern = /^(jr|sr|ii|iii|iv)\.?$/i;
        const trimmed = name.trim().replace(/\s+/g, ' ');
        const tokens = trimmed.split(' ');
        const filtered = [];
        for (let i = 0; i < tokens.length; i++) {
            if (tokens[i]) {
                filtered.push(tokens[i]);
            }
        }
        if (filtered.length === 0) {
            return ['', ''];
        }
        if (filtered.length === 1) {
            return [filtered[0], filtered[0]];
        }
        const first = filtered[0];
        let last = filtered[filtered.length - 1];
        if (filtered.length > 2 && suffixPattern.test(last)) {
            last = filtered[filtered.length - 2];
        }
        return [first, last];
    }

    function normalizeFirstLastPair(name) {
        const pair = getFirstAndLast(name);
        return elogNormalizeName(pair[0] + ' ' + pair[1]);
    }

    function buildExistingPairsFromScan(scannedNamesArray) {
        addLogMessage('buildExistingPairsFromScan: building from ' + scannedNamesArray.length + ' names', 'log');
        const pairs = new Set();
        for (let i = 0; i < scannedNamesArray.length; i++) {
            const pk = normalizeFirstLastPair(scannedNamesArray[i]);
            if (pk) {
                pairs.add(pk);
            }
        }
        addLogMessage('buildExistingPairsFromScan: built ' + pairs.size + ' pairs', 'log');
        return pairs;
    }

    function buildUserQueueSorted(userNamesArray) {
        addLogMessage('buildUserQueueSorted: building from ' + userNamesArray.length + ' names', 'log');
        const seenPairKeys = new Set();
        const unique = [];
        const duplicateIndices = [];
        for (let i = 0; i < userNamesArray.length; i++) {
            const nameObj = userNamesArray[i];
            const pk = normalizeFirstLastPair(nameObj.display);
            if (seenPairKeys.has(pk)) {
                addLogMessage('buildUserQueueSorted: input duplicate pairKey=' + pk + ' display=' + nameObj.display, 'log');
                duplicateIndices.push(i);
                continue;
            }
            seenPairKeys.add(pk);
            unique.push({
                display: nameObj.display,
                normalized: nameObj.normalized,
                pairKey: pk,
                status: ELOG_RUN_LABELS.statusPending
            });
        }
        unique.sort(function(a, b) {
            const aPair = getFirstAndLast(a.display);
            const bPair = getFirstAndLast(b.display);
            const firstCmp = aPair[0].localeCompare(bPair[0], undefined, { sensitivity: 'base', numeric: true });
            if (firstCmp !== 0) {
                return firstCmp;
            }
            const lastCmp = aPair[1].localeCompare(bPair[1], undefined, { sensitivity: 'base', numeric: true });
            if (lastCmp !== 0) {
                return lastCmp;
            }
            return a.display.localeCompare(b.display, undefined, { sensitivity: 'base', numeric: true });
        });
        addLogMessage('buildUserQueueSorted: unique=' + unique.length + ' duplicates=' + duplicateIndices.length, 'log');
        return { queue: unique, duplicateIndices: duplicateIndices };
    }

    function updateRightPanelStatus(pairKey, newStatus) {
        addLogMessage('updateRightPanelStatus: pairKey=' + pairKey + ' status=' + newStatus, 'log');
        const rightPanel = document.getElementById('elog-right-panel');
        if (!rightPanel) {
            addLogMessage('updateRightPanelStatus: right panel not found', 'error');
            return;
        }
        const items = rightPanel.querySelectorAll('.' + ELOG_CSS_CLASSNAMES.listItem);
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const itemPairKey = item.getAttribute('data-pairkey');
            if (!itemPairKey) {
                continue;
            }
            if (itemPairKey === pairKey) {
                const badge = item.querySelector('.elog-status-badge');
                if (badge) {
                    badge.textContent = newStatus;
                    let badgeColor = 'rgba(255, 255, 255, 0.7)';
                    let badgeBg = 'rgba(255, 255, 255, 0.1)';
                    if (newStatus === ELOG_RUN_LABELS.statusAdded) {
                        badgeColor = '#6bcf7f';
                        badgeBg = 'rgba(107, 207, 127, 0.2)';
                    }
                    else if (newStatus === ELOG_RUN_LABELS.statusAlready) {
                        badgeColor = '#ffa500';
                        badgeBg = 'rgba(255, 165, 0, 0.2)';
                    }
                    else if (newStatus === ELOG_RUN_LABELS.statusNotInDropdown || newStatus === ELOG_RUN_LABELS.statusSelectionFailed || newStatus === ELOG_RUN_LABELS.statusSaveFailed) {
                        badgeColor = '#ff6b6b';
                        badgeBg = 'rgba(255, 107, 107, 0.2)';
                    }
                    else if (newStatus === ELOG_RUN_LABELS.statusStopped) {
                        badgeColor = '#aaa';
                        badgeBg = 'rgba(170, 170, 170, 0.2)';
                    }
                    else if (newStatus === ELOG_RUN_LABELS.statusPending) {
                        badgeColor = '#ffd93d';
                        badgeBg = 'rgba(255, 217, 61, 0.2)';
                    }
                    badge.style.color = badgeColor;
                    badge.style.background = badgeBg;
                }
                addLogMessage('updateRightPanelStatus: updated item for pairKey=' + pairKey, 'log');
                break;
            }
        }
    }

    function updateRightPanelSummary(counters) {
        addLogMessage('updateRightPanelSummary: total=' + counters.total + ' added=' + counters.added + ' dup=' + counters.duplicates + ' fail=' + counters.failures + ' pending=' + counters.pending, 'log');
        const totalEl = document.getElementById('elog-summary-total');
        const addedEl = document.getElementById('elog-summary-added');
        const dupEl = document.getElementById('elog-summary-duplicates');
        const failEl = document.getElementById('elog-summary-failures');
        const pendingEl = document.getElementById('elog-summary-pending');
        const percentEl = document.getElementById('elog-summary-percent');
        if (totalEl) {
            totalEl.textContent = String(counters.total);
        }
        if (addedEl) {
            addedEl.textContent = String(counters.added);
        }
        if (dupEl) {
            dupEl.textContent = String(counters.duplicates);
        }
        if (failEl) {
            failEl.textContent = String(counters.failures);
        }
        if (pendingEl) {
            pendingEl.textContent = String(counters.pending);
        }
        if (percentEl) {
            const processed = counters.total - counters.pending;
            const pct = counters.total > 0 ? Math.round((processed / counters.total) * 100) : 0;
            percentEl.textContent = pct + '%';
        }
        updateAriaLiveRegion('Added: ' + counters.added + ', Duplicates: ' + counters.duplicates + ', Pending: ' + counters.pending);
    }

    function ensureAddEntryFormOpen() {
        addLogMessage('ensureAddEntryFormOpen: checking form state', 'log');
        return new Promise(function(resolve, reject) {
            const memberInput = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
            if (memberInput) {
                addLogMessage('ensureAddEntryFormOpen: member input already present', 'log');
                resolve(memberInput);
                return;
            }
            const addBtn = document.querySelector(ELOG_FORM_SELECTORS.addEntryBtn);
            if (addBtn) {
                addLogMessage('ensureAddEntryFormOpen: clicking Add Entry button', 'log');
                addBtn.click();
            } else {
                addLogMessage('ensureAddEntryFormOpen: Add Entry button not found, waiting for input', 'warn');
            }
            waitForElement(ELOG_FORM_SELECTORS.memberInput, ELOG_FORM_TIMEOUTS.waitOpenMs)
                .then(function(el) {
                addLogMessage('ensureAddEntryFormOpen: member input found', 'log');
                resolve(el);
            })
                .catch(function(err) {
                addLogMessage('ensureAddEntryFormOpen: timeout waiting for member input: ' + err, 'error');
                reject(err);
            });
        });
    }

    function ensureMemberDropdownOpen() {
        addLogMessage('ensureMemberDropdownOpen: checking dropdown', 'log');
        return new Promise(function(resolve, reject) {
            let retries = 0;
            function tryOpen() {
                addLogMessage('ensureMemberDropdownOpen: attempt ' + (retries + 1), 'log');
                const listEl = document.querySelector(ELOG_FORM_SELECTORS.listContainer);
                if (listEl) {
                    addLogMessage('ensureMemberDropdownOpen: list already open', 'log');
                    resolve(listEl);
                    return;
                }
                const inputEl = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
                if (inputEl) {
                    addLogMessage('ensureMemberDropdownOpen: clicking input to open list', 'log');
                    inputEl.click();
                    inputEl.focus();
                }
                waitForElement(ELOG_FORM_SELECTORS.listContainer, ELOG_FORM_TIMEOUTS.waitListMs)
                    .then(function(el) {
                    addLogMessage('ensureMemberDropdownOpen: list opened', 'log');
                    resolve(el);
                })
                    .catch(function(err) {
                    retries++;
                    if (retries < ELOG_FORM_RETRY.openListRetries) {
                        addLogMessage('ensureMemberDropdownOpen: retry ' + retries, 'warn');
                        const tid = setTimeout(tryOpen, 300);
                        elogState.timeouts.push(tid);
                    } else {
                        addLogMessage('ensureMemberDropdownOpen: exhausted retries', 'error');
                        reject(err);
                    }
                });
            }
            tryOpen();
        });
    }

    function findDropdownViewportElement() {
        addLogMessage('findDropdownViewportElement: searching', 'log');
        const el = document.querySelector(ELOG_FORM_SELECTORS.virtualViewport);
        if (el) {
            addLogMessage('findDropdownViewportElement: found', 'log');
        } else {
            addLogMessage('findDropdownViewportElement: not found', 'warn');
        }
        return el;
    }

    function rememberListScrollPosition(viewportEl) {
        if (viewportEl) {
            elogState.listScrollTop = viewportEl.scrollTop;
            addLogMessage('rememberListScrollPosition: saved=' + elogState.listScrollTop, 'log');
        }
    }

    function restoreListScrollPosition(viewportEl) {
        if (viewportEl && elogState.listScrollTop > 0) {
            viewportEl.scrollTop = elogState.listScrollTop;
            addLogMessage('restoreListScrollPosition: restored=' + elogState.listScrollTop, 'log');
        }
    }

    function typeIntoFilteredInput(inputEl, text) {
        addLogMessage('typeIntoFilteredInput: typing "' + text + '"', 'log');
        inputEl.focus();
        inputEl.value = text;
        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function clearFilteredInput(inputEl) {
        addLogMessage('clearFilteredInput: clearing input', 'log');
        inputEl.focus();
        inputEl.value = '';
        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function scanFilteredOptionsForMatch(targetPairKey, selectorSet) {
        var options = document.querySelectorAll(selectorSet.optionItem);
        addLogMessage('scanFilteredOptionsForMatch: checking ' + options.length + ' options for pairKey=' + targetPairKey, 'log');
        for (var oi = 0; oi < options.length; oi++) {
            var optText = options[oi].getAttribute('aria-label') || (options[oi].textContent || '').trim();
            optText = optText.trim();
            var optPairKey = normalizeFirstLastPair(optText);
            if (optPairKey === targetPairKey) {
                addLogMessage('scanFilteredOptionsForMatch: exact match at index ' + oi + ' text=' + optText, 'log');
                return { element: options[oi], matchType: 'exact' };
            }
        }
        var targetParts = targetPairKey.split(' ');
        for (var si = 0; si < options.length; si++) {
            var sOptText = options[si].getAttribute('aria-label') || (options[si].textContent || '').trim();
            sOptText = sOptText.trim();
            var sOptPairKey = normalizeFirstLastPair(sOptText);
            var optParts = sOptPairKey.split(' ');
            if (targetParts.length >= 2 && optParts.length >= 2) {
                if (targetParts[0] === optParts[0] && targetParts[targetParts.length - 1] === optParts[optParts.length - 1]) {
                    addLogMessage('scanFilteredOptionsForMatch: similar match at index ' + si + ' text=' + sOptText, 'log');
                    return { element: options[si], matchType: 'similar' };
                }
            }
        }
        addLogMessage('scanFilteredOptionsForMatch: no match found', 'log');
        return null;
    }

    function attemptSelectByScrollingForName(targetDisplay, targetPairKey) {
        addLogMessage('attemptSelectByScrollingForName: target=' + targetDisplay + ' pairKey=' + targetPairKey, 'log');
        return new Promise(function(resolve) {
            var inputEl = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
            if (!inputEl) {
                addLogMessage('attemptSelectByScrollingForName: member input not found', 'error');
                resolve(false);
                return;
            }
            var nameParts = getFirstAndLast(targetDisplay);
            var firstName = nameParts[0];
            var lastName = nameParts[1];
            addLogMessage('attemptSelectByScrollingForName: firstName=' + firstName + ' lastName=' + lastName, 'log');
            typeIntoFilteredInput(inputEl, firstName);
            var firstNameTid = setTimeout(function() {
                if (!elogState.isRunning) {
                    clearFilteredInput(inputEl);
                    resolve(false);
                    return;
                }
                var match = scanFilteredOptionsForMatch(targetPairKey, ELOG_FORM_SELECTORS);
                if (match) {
                    addLogMessage('attemptSelectByScrollingForName: found via firstName filter (' + match.matchType + ')', 'log');
                    match.element.click();
                    var verifyTid1 = setTimeout(function() {
                        resolve(true);
                    }, ELOG_FORM_TIMEOUTS.settleMs);
                    elogState.timeouts.push(verifyTid1);
                    return;
                }
                addLogMessage('attemptSelectByScrollingForName: not found with firstName, trying lastName', 'log');
                if (lastName && lastName !== firstName) {
                    typeIntoFilteredInput(inputEl, lastName);
                    var lastNameTid = setTimeout(function() {
                        if (!elogState.isRunning) {
                            clearFilteredInput(inputEl);
                            resolve(false);
                            return;
                        }
                        var match2 = scanFilteredOptionsForMatch(targetPairKey, ELOG_FORM_SELECTORS);
                        if (match2) {
                            addLogMessage('attemptSelectByScrollingForName: found via lastName filter (' + match2.matchType + ')', 'log');
                            match2.element.click();
                            var verifyTid2 = setTimeout(function() {
                                resolve(true);
                            }, ELOG_FORM_TIMEOUTS.settleMs);
                            elogState.timeouts.push(verifyTid2);
                            return;
                        }
                        addLogMessage('attemptSelectByScrollingForName: not found with lastName, falling back to scroll', 'log');
                        clearFilteredInput(inputEl);
                        var scrollFallbackTid = setTimeout(function() {
                            scrollSearchForName(targetPairKey, ELOG_FORM_SELECTORS, ELOG_FORM_TIMEOUTS, ELOG_FORM_RETRY, elogState, resolve);
                        }, ELOG_FORM_TIMEOUTS.waitFilterMs);
                        elogState.timeouts.push(scrollFallbackTid);
                    }, ELOG_FORM_TIMEOUTS.waitFilterMs);
                    elogState.timeouts.push(lastNameTid);
                } else {
                    addLogMessage('attemptSelectByScrollingForName: lastName same as firstName, falling back to scroll', 'log');
                    clearFilteredInput(inputEl);
                    var scrollTid = setTimeout(function() {
                        scrollSearchForName(targetPairKey, ELOG_FORM_SELECTORS, ELOG_FORM_TIMEOUTS, ELOG_FORM_RETRY, elogState, resolve);
                    }, ELOG_FORM_TIMEOUTS.waitFilterMs);
                    elogState.timeouts.push(scrollTid);
                }
            }, ELOG_FORM_TIMEOUTS.waitFilterMs);
            elogState.timeouts.push(firstNameTid);
        });
    }

    function scrollSearchForName(targetPairKey, selectors, timeouts, retryConfig, state, resolve) {
        addLogMessage('scrollSearchForName: fallback scroll for pairKey=' + targetPairKey, 'log');
        var viewportEl = document.querySelector(selectors.virtualViewport);
        if (!viewportEl) {
            addLogMessage('scrollSearchForName: no viewport element', 'error');
            resolve(false);
            return;
        }
        var passCount = 0;
        var lastScrollTop = -1;
        var lastOptionSnapshot = '';
        var startTime = Date.now();
        function scanAndScroll() {
            if (!state.isRunning) {
                resolve(false);
                return;
            }
            if (Date.now() - startTime > timeouts.maxSelectDurationMs) {
                addLogMessage('scrollSearchForName: max duration exceeded', 'warn');
                resolve(false);
                return;
            }
            if (passCount >= retryConfig.maxScrollPasses) {
                addLogMessage('scrollSearchForName: max passes reached', 'warn');
                resolve(false);
                return;
            }
            var options = document.querySelectorAll(selectors.optionItem);
            var retryCount = 0;
            function tryScanOptions() {
                options = document.querySelectorAll(selectors.optionItem);
                if (options.length === 0 && retryCount < retryConfig.selectRetriesPerScroll) {
                    retryCount++;
                    var rtid = setTimeout(tryScanOptions, timeouts.waitOptionRenderMs);
                    state.timeouts.push(rtid);
                    return;
                }
                var found = false;
                var optionTexts = [];
                for (var oi = 0; oi < options.length; oi++) {
                    var optText = options[oi].getAttribute('aria-label') || (options[oi].textContent || '').trim();
                    optText = optText.trim();
                    optionTexts.push(optText);
                    var optPairKey = normalizeFirstLastPair(optText);
                    if (optPairKey === targetPairKey) {
                        addLogMessage('scrollSearchForName: match found at index ' + oi + ' text=' + optText, 'log');
                        options[oi].click();
                        found = true;
                        var verifyTid = setTimeout(function() {
                            resolve(true);
                        }, timeouts.settleMs);
                        state.timeouts.push(verifyTid);
                        break;
                    }
                }
                if (!found) {
                    var currentSnapshot = optionTexts.join('|');
                    var currentTop = viewportEl.scrollTop;
                    var noProgress = (currentTop === lastScrollTop && currentSnapshot === lastOptionSnapshot);
                    if (noProgress) {
                        passCount++;
                        addLogMessage('scrollSearchForName: no progress pass ' + passCount, 'log');
                    }
                    lastScrollTop = currentTop;
                    lastOptionSnapshot = currentSnapshot;
                    var stepSize = Math.round(viewportEl.clientHeight * 0.7);
                    var maxScroll = viewportEl.scrollHeight - viewportEl.clientHeight;
                    var newTop = Math.min(currentTop + stepSize, maxScroll);
                    if (newTop <= currentTop && currentTop > 0) {
                        newTop = 0;
                        lastScrollTop = -1;
                        lastOptionSnapshot = '';
                        passCount++;
                    }
                    viewportEl.scrollTop = newTop;
                    var scrollTid = setTimeout(scanAndScroll, timeouts.scrollIdleMs);
                    state.timeouts.push(scrollTid);
                }
            }
            tryScanOptions();
        }
        var initTid = setTimeout(scanAndScroll, timeouts.scrollIdleMs);
        state.timeouts.push(initTid);
    }

    function clickSaveAndAddAnother() {
        addLogMessage('clickSaveAndAddAnother: searching for button', 'log');
        return new Promise(function(resolve) {
            var buttons = document.querySelectorAll(ELOG_FORM_SELECTORS.saveAndAddAnotherBtn);
            var targetBtn = null;
            for (var bi = 0; bi < buttons.length; bi++) {
                var btnText = (buttons[bi].innerText || '').toLowerCase();
                if (btnText.indexOf('save') !== -1 && btnText.indexOf('add another') !== -1) {
                    targetBtn = buttons[bi];
                    break;
                }
            }
            if (!targetBtn) {
                addLogMessage('clickSaveAndAddAnother: button not found', 'error');
                resolve('not_found');
                return;
            }
            if (targetBtn.disabled || targetBtn.getAttribute('disabled') !== null) {
                addLogMessage('clickSaveAndAddAnother: button is disabled', 'warn');
                resolve('disabled');
                return;
            }
            addLogMessage('clickSaveAndAddAnother: clicking button', 'log');
            targetBtn.click();
            var waitStart = Date.now();
            function checkSaveResult() {
                if (Date.now() - waitStart > ELOG_FORM_TIMEOUTS.waitSaveAfterClickMs) {
                    addLogMessage('clickSaveAndAddAnother: timeout waiting for save result', 'warn');
                    resolve('timeout');
                    return;
                }
                var inputEl = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
                if (inputEl && (!inputEl.value || !inputEl.value.trim())) {
                    addLogMessage('clickSaveAndAddAnother: input cleared, save successful', 'log');
                    resolve('success');
                    return;
                }
                var listGone = !document.querySelector(ELOG_FORM_SELECTORS.listContainer);
                if (listGone && inputEl && (!inputEl.value || !inputEl.value.trim())) {
                    addLogMessage('clickSaveAndAddAnother: list closed and input clear', 'log');
                    resolve('success');
                    return;
                }
                var tid = setTimeout(checkSaveResult, 200);
                elogState.timeouts.push(tid);
            }
            var initTid = setTimeout(checkSaveResult, 300);
            elogState.timeouts.push(initTid);
        });
    }

    function processNextStaffFromQueue() {
        addLogMessage('processNextStaffFromQueue: index=' + elogState.addQueueIndex + ' of ' + elogState.addQueue.length, 'log');
        if (!elogState.isRunning) {
            addLogMessage('processNextStaffFromQueue: stopped, marking remaining as Stopped', 'warn');
            for (var si = elogState.addQueueIndex; si < elogState.addQueue.length; si++) {
                if (elogState.addQueue[si].status === ELOG_RUN_LABELS.statusPending) {
                    elogState.addQueue[si].status = ELOG_RUN_LABELS.statusStopped;
                    updateRightPanelStatus(elogState.addQueue[si].pairKey, ELOG_RUN_LABELS.statusStopped);
                    elogState.counters.pending--;
                }
            }
            updateRightPanelSummary(elogState.counters);
            elogState.isAddingEntries = false;
            updateScanStatus('Stopped', 'stopped');
            var title = document.getElementById('elog-progress-title');
            if (title) {
                title.textContent = 'ELog Staff Entries - Stopped';
            }
            if (elogState.focusReturnElement) {
                elogState.focusReturnElement.focus();
            }
            return;
        }
        if (elogState.addQueueIndex >= elogState.addQueue.length) {
            addLogMessage('processNextStaffFromQueue: all candidates processed', 'log');
            elogState.isAddingEntries = false;
            updateScanStatus('Complete', 'complete');
            var titleEl = document.getElementById('elog-progress-title');
            if (titleEl) {
                titleEl.textContent = 'ELog Staff Entries - Complete';
            }
            updateRightPanelSummary(elogState.counters);
            updateAriaLiveRegion('Processing complete. Added: ' + elogState.counters.added + ', Duplicates: ' + elogState.counters.duplicates + ', Failures: ' + elogState.counters.failures);
            if (elogState.counters.added > 0) {
                var existingBtn = document.getElementById('elog-select-checkboxes-btn');
                if (!existingBtn) {
                    var btnContainer = document.getElementById('elog-right-panel');
                    if (btnContainer) {
                        var selectBtn = document.createElement('button');
                        selectBtn.id = 'elog-select-checkboxes-btn';
                        selectBtn.textContent = 'Select Checkboxes';
                        selectBtn.style.cssText = 'display: block; width: 100%; margin-top: 12px; padding: 10px 16px; background: rgba(100, 149, 237, 0.3); color: #6495ed; border: 1px solid rgba(100, 149, 237, 0.4); border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.2s ease;';
                        selectBtn.onmouseover = function() { selectBtn.style.background = 'rgba(100, 149, 237, 0.45)'; };
                        selectBtn.onmouseout = function() { selectBtn.style.background = 'rgba(100, 149, 237, 0.3)'; };
                        selectBtn.onclick = function() {
                            selectCheckboxesForAddedNames();
                        };
                        btnContainer.parentElement.appendChild(selectBtn);
                    }
                }
            }
            var memberInput = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
            if (memberInput) {
                memberInput.focus();
            } else if (elogState.focusReturnElement) {
                elogState.focusReturnElement.focus();
            }
            return;
        }
        var candidate = elogState.addQueue[elogState.addQueueIndex];
        addLogMessage('processNextStaffFromQueue: processing ' + candidate.display + ' pairKey=' + candidate.pairKey, 'log');
        if (elogState.existingPairs.has(candidate.pairKey)) {
            addLogMessage('processNextStaffFromQueue: already exists in table', 'log');
            candidate.status = ELOG_RUN_LABELS.statusAlready;
            elogState.counters.duplicates++;
            elogState.counters.pending--;
            updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusAlready);
            updateRightPanelSummary(elogState.counters);
            elogState.addQueueIndex++;
            var tid1 = setTimeout(processNextStaffFromQueue, 50);
            elogState.timeouts.push(tid1);
            return;
        }
        ensureAddEntryFormOpen()
            .then(function() {
            addLogMessage('processNextStaffFromQueue: form open, opening dropdown', 'log');
            return ensureMemberDropdownOpen();
        })
            .then(function() {
            addLogMessage('processNextStaffFromQueue: dropdown open, selecting name', 'log');
            return attemptSelectByScrollingForName(candidate.display, candidate.pairKey);
        })
            .then(function(selected) {
            if (!selected) {
                addLogMessage('processNextStaffFromQueue: not found in dropdown', 'warn');
                candidate.status = ELOG_RUN_LABELS.statusNotInDropdown;
                elogState.counters.failures++;
                elogState.counters.pending--;
                updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusNotInDropdown);
                updateRightPanelSummary(elogState.counters);
                elogState.addQueueIndex++;
                var tid2 = setTimeout(processNextStaffFromQueue, 50);
                elogState.timeouts.push(tid2);
                return;
            }
            addLogMessage('processNextStaffFromQueue: selected, clicking save', 'log');
            clickSaveAndAddAnother().then(function(saveResult) {
                if (saveResult === 'disabled') {
                    addLogMessage('processNextStaffFromQueue: save disabled, retrying selection', 'warn');
                    ensureMemberDropdownOpen()
                        .then(function() {
                        return attemptSelectByScrollingForName(candidate.display, candidate.pairKey);
                    })
                        .then(function(reselected) {
                        if (!reselected) {
                            candidate.status = ELOG_RUN_LABELS.statusSelectionFailed;
                            elogState.counters.failures++;
                            elogState.counters.pending--;
                            updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusSelectionFailed);
                            updateRightPanelSummary(elogState.counters);
                            elogState.addQueueIndex++;
                            var tid3 = setTimeout(processNextStaffFromQueue, 50);
                            elogState.timeouts.push(tid3);
                            return;
                        }
                        return clickSaveAndAddAnother().then(function(retryResult) {
                            if (retryResult === 'success') {
                                candidate.status = ELOG_RUN_LABELS.statusAdded;
                                elogState.counters.added++;
                                elogState.counters.pending--;
                                updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusAdded);
                            } else {
                                candidate.status = ELOG_RUN_LABELS.statusSelectionFailed;
                                elogState.counters.failures++;
                                elogState.counters.pending--;
                                updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusSelectionFailed);
                            }
                            updateRightPanelSummary(elogState.counters);
                            elogState.addQueueIndex++;
                            var tid4 = setTimeout(processNextStaffFromQueue, 50);
                            elogState.timeouts.push(tid4);
                        });
                    })
                        .catch(function(err) {
                        addLogMessage('processNextStaffFromQueue: retry error: ' + err, 'error');
                        candidate.status = ELOG_RUN_LABELS.statusSelectionFailed;
                        elogState.counters.failures++;
                        elogState.counters.pending--;
                        updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusSelectionFailed);
                        updateRightPanelSummary(elogState.counters);
                        elogState.addQueueIndex++;
                        var tid5 = setTimeout(processNextStaffFromQueue, 50);
                        elogState.timeouts.push(tid5);
                    });
                    return;
                }
                if (saveResult === 'success') {
                    addLogMessage('processNextStaffFromQueue: saved successfully', 'log');
                    candidate.status = ELOG_RUN_LABELS.statusAdded;
                    elogState.counters.added++;
                    elogState.counters.pending--;
                    updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusAdded);
                } else {
                    addLogMessage('processNextStaffFromQueue: save failed result=' + saveResult, 'warn');
                    candidate.status = ELOG_RUN_LABELS.statusSaveFailed;
                    elogState.counters.failures++;
                    elogState.counters.pending--;
                    updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusSaveFailed);
                }
                updateRightPanelSummary(elogState.counters);
                elogState.addQueueIndex++;
                var tid6 = setTimeout(processNextStaffFromQueue, 50);
                elogState.timeouts.push(tid6);
            });
        })
            .catch(function(err) {
            addLogMessage('processNextStaffFromQueue: error: ' + err, 'error');
            candidate.status = ELOG_RUN_LABELS.statusSaveFailed;
            elogState.counters.failures++;
            elogState.counters.pending--;
            updateRightPanelStatus(candidate.pairKey, ELOG_RUN_LABELS.statusSaveFailed);
            updateRightPanelSummary(elogState.counters);
            elogState.addQueueIndex++;
            var tid7 = setTimeout(processNextStaffFromQueue, 50);
            elogState.timeouts.push(tid7);
        });
    }

    function beginAddNonDuplicateELogEntries() {
        addLogMessage('beginAddNonDuplicateELogEntries: starting', 'log');
        elogState.existingPairs = buildExistingPairsFromScan(elogState.scannedNames);
        addLogMessage('beginAddNonDuplicateELogEntries: existingPairs count=' + elogState.existingPairs.size, 'log');
        var result = buildUserQueueSorted(elogState.parsedNames);
        elogState.addQueue = result.queue;
        elogState.addQueueIndex = 0;
        elogState.listScrollTop = 0;
        elogState.isAddingEntries = true;
        for (var di = 0; di < result.duplicateIndices.length; di++) {
            var dupIdx = result.duplicateIndices[di];
            var dupName = elogState.parsedNames[dupIdx];
            if (dupName) {
                var dupPk = normalizeFirstLastPair(dupName.display);
                updateRightPanelStatus(dupPk, ELOG_RUN_LABELS.statusAlready);
            }
        }
        elogState.counters = {
            total: elogState.addQueue.length,
            added: 0,
            duplicates: 0,
            failures: 0,
            pending: elogState.addQueue.length
        };
        updateRightPanelSummary(elogState.counters);
        updateScanStatus('Adding Entries', 'progress');
        var title = document.getElementById('elog-progress-title');
        if (title) {
            title.textContent = 'ELog Staff Entries - Adding Entries';
        }
        updateAriaLiveRegion('Starting to add ' + elogState.addQueue.length + ' entries');
        addLogMessage('beginAddNonDuplicateELogEntries: queue size=' + elogState.addQueue.length + ', starting processing', 'log');
        processNextStaffFromQueue();
    }

    function selectCheckboxesForAddedNames() {
        addLogMessage('selectCheckboxesForAddedNames: starting', 'log');
        var addedPairKeys = new Set();
        for (var qi = 0; qi < elogState.addQueue.length; qi++) {
            if (elogState.addQueue[qi].status === ELOG_RUN_LABELS.statusAdded) {
                addedPairKeys.add(elogState.addQueue[qi].pairKey);
            }
        }
        addLogMessage('selectCheckboxesForAddedNames: addedPairKeys count=' + addedPairKeys.size, 'log');
        if (addedPairKeys.size === 0) {
            addLogMessage('selectCheckboxesForAddedNames: no added entries to select', 'warn');
            return;
        }
        var gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
        if (!gridTable) {
            addLogMessage('selectCheckboxesForAddedNames: grid table not found', 'error');
            return;
        }
        var container = findScrollableContainer(gridTable);
        if (!container) {
            addLogMessage('selectCheckboxesForAddedNames: scrollable container not found', 'error');
            return;
        }
        var btn = document.getElementById('elog-select-checkboxes-btn');
        if (btn) {
            btn.disabled = true;
            btn.textContent = 'Selecting...';
        }
        var checkedPairKeys = new Set();
        var noProgress = 0;
        var lastScrollTop = -1;
        var lastRowSnapshot = '';
        container.scrollTo({ top: 0, behavior: 'auto' });
        function scanAndCheck() {
            var rows = gridTable.querySelectorAll(ELOG_SELECTORS.row);
            var rowTexts = [];
            for (var ri = 0; ri < rows.length; ri++) {
                var row = rows[ri];
                if (row.getAttribute('role') === 'columnheader') { continue; }
                var name = extractNameFromRow(row, ri + 1);
                if (!name) { continue; }
                rowTexts.push(name);
                var pk = normalizeFirstLastPair(name);
                if (addedPairKeys.has(pk) && !checkedPairKeys.has(pk)) {
                    var checkbox = row.querySelector('i[role="checkbox"]');
                    if (checkbox && checkbox.getAttribute('aria-checked') === 'false') {
                        addLogMessage('selectCheckboxesForAddedNames: clicking checkbox for ' + name, 'log');
                        checkbox.click();
                        checkedPairKeys.add(pk);
                    } else if (checkbox && checkbox.getAttribute('aria-checked') === 'true') {
                        addLogMessage('selectCheckboxesForAddedNames: already checked for ' + name, 'log');
                        checkedPairKeys.add(pk);
                    }
                }
            }
            return rowTexts.join('|');
        }
        var initialSnapshot = scanAndCheck();
        function scrollLoop() {
            if (checkedPairKeys.size >= addedPairKeys.size) {
                addLogMessage('selectCheckboxesForAddedNames: all checkboxes selected (' + checkedPairKeys.size + ')', 'log');
                finishSelection();
                return;
            }
            var currTop = container.scrollTop;
            var maxScroll = container.scrollHeight - container.clientHeight;
            var newTop = Math.min(currTop + ELOG_SCROLL.stepPx, maxScroll);
            container.scrollTo({ top: newTop, behavior: 'auto' });
            setTimeout(function() {
                var snapshot = scanAndCheck();
                var currentTop = container.scrollTop;
                if (currentTop === lastScrollTop && snapshot === lastRowSnapshot) {
                    noProgress++;
                } else {
                    noProgress = 0;
                }
                lastScrollTop = currentTop;
                lastRowSnapshot = snapshot;
                if (noProgress >= ELOG_SCROLL.maxNoProgressIterations || currentTop >= maxScroll) {
                    addLogMessage('selectCheckboxesForAddedNames: end of table reached, checked=' + checkedPairKeys.size + ' of ' + addedPairKeys.size, 'log');
                    finishSelection();
                    return;
                }
                scrollLoop();
            }, ELOG_SCROLL.settleDelayMs);
        }
        function finishSelection() {
            addLogMessage('selectCheckboxesForAddedNames: finished, checked ' + checkedPairKeys.size + ' checkboxes', 'log');
            if (btn) {
                btn.textContent = 'Done (' + checkedPairKeys.size + ' selected)';
                btn.disabled = true;
                btn.style.background = 'rgba(107, 207, 127, 0.3)';
                btn.style.color = '#6bcf7f';
            }
        }
        setTimeout(scrollLoop, ELOG_SCROLL.settleDelayMs);
    }

    function findScrollableContainer(gridEl) {
        if (!gridEl) {
            addLogMessage('findScrollableContainer: gridEl is null', 'warn');
            return null;
        }
        let current = gridEl;
        while (current && current !== document.body) {
            const style = window.getComputedStyle(current);
            const overflowY = style.overflowY;
            const scrollDiff = current.scrollHeight - current.clientHeight;
            if (scrollDiff >= 20 && (overflowY === 'scroll' || overflowY === 'auto')) {
                return current;
            }
            current = current.parentElement;
        }
        return gridEl;
    }

    function getRenderedRowCount() {
        const gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
        if (!gridTable) {
            addLogMessage('getRenderedRowCount: grid not found', 'warn');
            return 0;
        }
        const rows = gridTable.querySelectorAll(ELOG_SELECTORS.row);
        let count = 0;
        for (let i = 0; i < rows.length; i++) {
            if (rows[i].getAttribute('role') !== 'columnheader') {
                count++;
            }
        }
        return count;
    }

    function getRenderedLastRowKey() {
        const gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
        if (!gridTable) {
            addLogMessage('getRenderedLastRowKey: grid not found', 'warn');
            return '';
        }
        const rows = gridTable.querySelectorAll(ELOG_SELECTORS.row);
        let lastDataRow = null;
        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i].getAttribute('role') !== 'columnheader') {
                lastDataRow = rows[i];
                break;
            }
        }
        if (!lastDataRow) {
            addLogMessage('getRenderedLastRowKey: no data rows', 'warn');
            return '';
        }
        const cells = lastDataRow.querySelectorAll(ELOG_SELECTORS.cell);
        if (cells.length > 0) {
            const firstCellText = cells[0].textContent.trim();
            if (firstCellText) {
                return firstCellText;
            }
        }
        let hash = 0;
        const content = lastDataRow.textContent || '';
        for (let i = 0; i < content.length; i++) {
            hash = ((hash << 5) - hash) + content.charCodeAt(i);
            hash = hash & hash;
        }
        const key = 'hash_' + hash;
        return key;
    }

    function awaitSettle(container) {
        return new Promise(function(resolve) {
            const gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
            let resolved = false;
            let observer = null;
            const timeoutId = setTimeout(function() {
                if (!resolved) {
                    resolved = true;
                    if (observer) {
                        observer.disconnect();
                        const idx = elogState.observers.indexOf(observer);
                        if (idx > -1) {
                            elogState.observers.splice(idx, 1);
                        }
                    }
                    resolve();
                }
            }, ELOG_SCROLL.settleDelayMs);
            elogState.timeouts.push(timeoutId);
            if (gridTable) {
                observer = new MutationObserver(function() {
                    if (!resolved) {
                        resolved = true;
                        clearTimeout(timeoutId);
                        observer.disconnect();
                        const idx = elogState.observers.indexOf(observer);
                        if (idx > -1) {
                            elogState.observers.splice(idx, 1);
                        }
                        resolve();
                    }
                });
                observer.observe(gridTable, { childList: true, subtree: true });
                elogState.observers.push(observer);
            }
        });
    }

    function setAriaBusyOn() {
        const target = document.querySelector(ELOG_ATTRS.ariaBusyTarget);
        if (target) {
            elogState.prevAriaBusy = target.getAttribute(ELOG_ATTRS.ariaBusyAttr);
            target.setAttribute(ELOG_ATTRS.ariaBusyAttr, 'true');
        }
    }

    function setAriaBusyOff() {
        const target = document.querySelector(ELOG_ATTRS.ariaBusyTarget);
        if (target) {
            if (elogState.prevAriaBusy !== null) {
                target.setAttribute(ELOG_ATTRS.ariaBusyAttr, elogState.prevAriaBusy);
            } else {
                target.removeAttribute(ELOG_ATTRS.ariaBusyAttr);
            }
            elogState.prevAriaBusy = null;
        }
    }

    function computeEndReached(container, noProgress) {
        if (!container) {
            return true;
        }
        const atBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 1;
        if (atBottom && noProgress >= 1) {
            return true;
        }
        return false;
    }

    function restoreViewport(container, prevTop) {
        if (!container) {
            return;
        }
        const diff = Math.abs(container.scrollTop - prevTop);
        if (diff > 100) {
            addLogMessage('restoreViewport: restoring scroll', 'log');
            container.scrollTo({ top: prevTop, behavior: 'auto' });
        }
    }

    function observeUserScrollPause(container) {
        addLogMessage('observeUserScrollPause: setup', 'log');
        if (!container || elogState.userScrollHandler) {
            return;
        }
        elogState.userScrollHandler = function() {
            const timeSinceAuto = Date.now() - elogState.lastAutoScrollTime;
            if (timeSinceAuto > 50 && !elogState.userScrollPaused) {
                elogState.userScrollPaused = true;
                const resumeTimeout = setTimeout(function() {
                    elogState.userScrollPaused = false;
                }, ELOG_SCROLL.userScrollPauseMs);
                elogState.timeouts.push(resumeTimeout);
            }
        };
        container.addEventListener('scroll', elogState.userScrollHandler);
        elogState.eventListeners.push({ element: container, type: 'scroll', handler: elogState.userScrollHandler });
    }

    function scanVisibleRowsOnce(onRow) {
        const gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
        if (!gridTable) {
            addLogMessage('scanVisibleRowsOnce: grid not found', 'warn');
            return 0;
        }
        const rows = gridTable.querySelectorAll(ELOG_SELECTORS.row);
        let newCount = 0;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row.getAttribute('role') === 'columnheader') {
                continue;
            }
            const extractedName = extractNameFromRow(row, i + 1);
            if (!extractedName) {
                continue;
            }
            const normalized = elogNormalizeName(extractedName);
            if (elogState.seenNormalizedNames.has(normalized)) {
                continue;
            }
            elogState.seenNormalizedNames.add(normalized);
            elogState.scannedNames.push(extractedName);
            newCount++;
            if (onRow) {
                onRow(extractedName, normalized);
            }
        }
        return newCount;
    }

    function updateAriaLiveRegion(message) {
        const liveRegion = document.getElementById('elog-aria-live');
        if (liveRegion) {
            liveRegion.textContent = message;
        }
    }

    function autoScrollScan(options) {
        addLogMessage('autoScrollScan: starting', 'log');
        const onRow = options.onRow || function() {};
        const onProgress = options.onProgress || function() {};
        const onDone = options.onDone || function() {};
        const onError = options.onError || function() {};
        const gridTable = document.querySelector(ELOG_SELECTORS.gridTable);
        if (!gridTable) {
            addLogMessage('autoScrollScan: grid not found', 'error');
            onError(new Error('Grid table not found'));
            return;
        }
        const container = findScrollableContainer(gridTable);
        if (!container) {
            addLogMessage('autoScrollScan: container not found', 'error');
            onError(new Error('Container not found'));
            return;
        }
        elogState.scrollContainer = container;
        elogState.prevScrollTop = container.scrollTop;
        elogState.seenNormalizedNames = new Set();
        elogState.leftPanelRowIndex = 0;
        setAriaBusyOn();
        observeUserScrollPause(container);
        const startTime = Date.now();
        let noProgress = 0;
        scanVisibleRowsOnce(function(name) {
            elogState.leftPanelRowIndex++;
        });
        onProgress({ scanned: elogState.scannedNames.length });
        function scrollLoop() {
            if (!elogState.isRunning) {
                updateAriaLiveRegion(ELOG_LABELS.progressStopped);
                finishScan(ELOG_LABELS.progressStopped, 'stopped');
                return;
            }
            if (Date.now() - startTime > ELOG_SCROLL.maxDurationMs) {
                finishScan(ELOG_LABELS.progressComplete, 'timeout');
                return;
            }
            if (elogState.userScrollPaused) {
                const pt = setTimeout(scrollLoop, 100);
                elogState.timeouts.push(pt);
                return;
            }
            const priorKey = getRenderedLastRowKey();
            const priorCount = getRenderedRowCount();
            const currTop = container.scrollTop;
            const maxScroll = container.scrollHeight - container.clientHeight;
            const newTop = Math.min(currTop + ELOG_SCROLL.stepPx, maxScroll);
            elogState.lastAutoScrollTime = Date.now();
            container.scrollTo({ top: newTop, behavior: 'auto' });
            awaitSettle(container).then(function() {
                let attempts = 0;
                function attemptScan() {
                    const rc = getRenderedRowCount();
                    if (rc === 0 && attempts < ELOG_SCROLL.retryScanAttempts) {
                        attempts++;
                        const rt = setTimeout(attemptScan, ELOG_SCROLL.retryScanDelayMs);
                        elogState.timeouts.push(rt);
                        return;
                    }
                    scanVisibleRowsOnce(function(name) {
                        elogState.leftPanelRowIndex++;
                    });
                    onProgress({ scanned: elogState.scannedNames.length });
                    const currKey = getRenderedLastRowKey();
                    const currCount = getRenderedRowCount();
                    if (currKey === priorKey && currCount === priorCount) {
                        noProgress++;
                    } else {
                        noProgress = 0;
                    }
                    if (computeEndReached(container, noProgress)) {
                        finishScan(ELOG_LABELS.progressNoMore, 'endReached');
                        return;
                    }
                    if (noProgress >= ELOG_SCROLL.maxNoProgressIterations) {
                        finishScan(ELOG_LABELS.progressNoMore, 'noProgress');
                        return;
                    }
                    if (typeof requestIdleCallback === 'function') {
                        elogState.idleCallbackId = requestIdleCallback(function() {
                            elogState.idleCallbackId = null;
                            scrollLoop();
                        }, { timeout: ELOG_SCROLL.idleDelayMs * 2 });
                    } else {
                        const it = setTimeout(scrollLoop, ELOG_SCROLL.idleDelayMs);
                        elogState.timeouts.push(it);
                    }
                }
                attemptScan();
            });
        }
        function finishScan(label, reason) {
            addLogMessage('autoScrollScan: done reason=' + reason + ' total=' + elogState.scannedNames.length, 'log');
            setAriaBusyOff();
            onDone({ total: elogState.scannedNames.length, reason: reason });
        }
        const initTimeout = setTimeout(scrollLoop, ELOG_SCROLL.idleDelayMs);
        elogState.timeouts.push(initTimeout);
    }

    function stopELog() {
        addLogMessage('stopELog: stopping all ELog processes', 'log');
        elogState.isRunning = false;
        if (elogState.idleCallbackId && typeof cancelIdleCallback === 'function') {
            addLogMessage('stopELog: cancelling idle callback', 'log');
            cancelIdleCallback(elogState.idleCallbackId);
            elogState.idleCallbackId = null;
        }
        for (let i = 0; i < elogState.observers.length; i++) {
            try {
                elogState.observers[i].disconnect();
            } catch (e) {
                addLogMessage('stopELog: error disconnecting observer: ' + e, 'error');
            }
        }
        elogState.observers = [];
        for (let i = 0; i < elogState.timeouts.length; i++) {
            try {
                clearTimeout(elogState.timeouts[i]);
            } catch (e) {
                addLogMessage('stopELog: error clearing timeout: ' + e, 'error');
            }
        }
        elogState.timeouts = [];
        for (let i = 0; i < elogState.intervals.length; i++) {
            try {
                clearInterval(elogState.intervals[i]);
            } catch (e) {
                addLogMessage('stopELog: error clearing interval: ' + e, 'error');
            }
        }
        elogState.intervals = [];
        for (let i = 0; i < elogState.eventListeners.length; i++) {
            try {
                const listener = elogState.eventListeners[i];
                listener.element.removeEventListener(listener.type, listener.handler);
            } catch (e) {
                addLogMessage('stopELog: error removing event listener: ' + e, 'error');
            }
        }
        elogState.eventListeners = [];
        if (elogState.abortController) {
            elogState.abortController.abort();
            elogState.abortController = null;
        }
        setAriaBusyOff();
        if (elogState.scrollContainer && elogState.prevScrollTop !== undefined) {
            addLogMessage('stopELog: restoring viewport', 'log');
            restoreViewport(elogState.scrollContainer, elogState.prevScrollTop);
        }
        elogState.scrollContainer = null;
        elogState.userScrollHandler = null;
        elogState.userScrollPaused = false;
        if (elogState.isAddingEntries) {
            addLogMessage('stopELog: was adding entries, marking remaining as Stopped', 'log');
            for (let qi = elogState.addQueueIndex; qi < elogState.addQueue.length; qi++) {
                if (elogState.addQueue[qi].status === ELOG_RUN_LABELS.statusPending) {
                    elogState.addQueue[qi].status = ELOG_RUN_LABELS.statusStopped;
                    elogState.counters.pending--;
                }
            }
            elogState.isAddingEntries = false;
        }
        elogState.addQueue = [];
        elogState.addQueueIndex = 0;
        elogState.existingPairs = new Set();
        elogState.listScrollTop = 0;
        const inputModal = document.getElementById('elog-input-modal');
        if (inputModal && inputModal.parentNode) {
            inputModal.parentNode.removeChild(inputModal);
        }
        const progressModal = document.getElementById('elog-progress-modal');
        if (progressModal && progressModal.parentNode) {
            progressModal.parentNode.removeChild(progressModal);
        }
        removeCollectingDataPanel('elog');
        resetELogState();
        addLogMessage('stopELog: cleanup complete', 'log');
    }

    //==========================
    // SET ROLE RESPONSIBILITIES FUNCTIONS
    //==========================
    // This section contains functions to handle Training ELogs feature.
    // This includes a UI user input, parsing a list of names, and adding entries on webpages.
    //==========================
    const RESP_SELECTORS = {
        pageStepRoot: 'doa-log-template-study-roles-step',
        dropListContainer: '.cdk-drop-list.doa-log-form-step__drop-list-container',
        roleColumns: '.doa-log-form-step__column.roles__column',
        roleSearchInput: '#filtered-select-input.filtered-select__input[placeholder*="Search"]',
        roleListContainer: 'ul.filtered-select__list.u-z-index-1060, ul.filtered-select__list',
        virtualViewport: 'cdk-virtual-scroll-viewport.cdk-virtual-scroll-viewport, .filtered-select__list',
        roleOptionItem: '.filtered-select__list__item, [role="option"], .cdk-virtual-scroll-viewport .filtered-select__list__item',
        roleOptionText: '.filtered-select__list__item__text',
        roleOptionCheckboxTri: '.test-checkboxTristate[role="checkbox"]',
        selectedRoleCheckboxInColumn: '[role="checkbox"][aria-checked="true"]',
        responsibilitiesToggleBtn: 'button[dropdowntoggle][data-test="study-responsibilities-dropdown-toggle"].roles__select-options-dropdown-button',
        responsibilitiesMenu: 'ul[role="menu"][data-test="study-responsibilities-list"]',
        responsibilitiesItem: 'ul[role="menu"][data-test="study-responsibilities-list"] li',
        responsibilitiesItemCheckbox: 'ul[role="menu"][data-test="study-responsibilities-list"] li input[type="checkbox"]',
        addStudyRoleBtn: 'button[data-test="add-study-role-button"]',
        mainPanelButtonTarget: '.main-gui-panel',
        ariaLiveRegion: '.aria-live-region'
    };

    const RESP_TIMEOUTS = {
        waitPageMs: 10000,
        waitInputPanelMs: 10000,
        waitProgressPanelMs: 10000,
        waitListOpenMs: 4000,
        waitOptionRenderMs: 1500,
        waitAfterSelectRoleMs: 800,
        waitResponsibilitiesMenuMs: 3000,
        waitAfterToggleResponsibilityMs: 200,
        waitAddRoleResultMs: 4000,
        settleAfterScrollMs: 150,
        idleBetweenScrollsMs: 80,
        maxSelectRoleDurationMs: 30000
    };

    const RESP_RETRY = {
        openListRetries: 4,
        optionScanRetries: 2,
        maxScrollPasses: 8,
        addRoleRetries: 2
    };

    const RESP_LABELS = {
        notOnPageWarning: 'You are not on the Study Role Page.',
        statusPending: 'Pending',
        statusCompleted: 'Completed',
        statusFailed: 'Failed',
        statusStopped: 'Stopped',
        parsing: 'Parsing input',
        scanningExisting: 'Scanning existing roles',
        selectingRole: 'Selecting role',
        selectingResponsibilities: 'Selecting responsibilities',
        addingRole: 'Adding study role',
        done: 'Completed'
    };

    const RESP_COUNTERS = {
        total: 0,
        completed: 0,
        failed: 0,
        pending: 0
    };

    const RESP_REGEX = {
        range: /\b(\d+)\s*(?:to|\-|\u2013|\u2014)\s*(\d+)\b/i,
        number: /\b\d+\b/g,
        quoteCleanup: /["\u201c\u201d]+/g,
        hyphenBreak: /-\s*\n\s*/g,
        lineBreakInRole: /\n+/g,
        whitespace: /\s+/g
    };

    const RESP_ROLE_ALIASES = {
        'pi': 'Principal Investigator',
        'principal investigator': 'Principal Investigator',
        'sub investigator': 'Sub-Investigator',
        'sub-investigator': 'Sub-Investigator',
        'research nurse': 'Research Nurse',
        'nurse': 'Research Nurse',
        'nurse manager': 'Research Nurse',
        'infusion nurse i': 'Research Nurse',
        'infusion nurse ii': 'Research Nurse',
        'research nurse i': 'Research Nurse',
        'research nurse ii': 'Research Nurse',
        'research assistant': 'Research Assistant',
        'research assistant i': 'Research Assistant',
        'research assistant ii': 'Research Assistant',
        'study coordinator': 'Study Coordinator',
        'crc i': 'Study Coordinator',
        'crc ii': 'Study Coordinator',
        'senior crc': 'Study Coordinator',
        'clinical trial lead': 'Study Coordinator',
        'pharmacy': 'Pharmacy',
        'pharmacist': 'Pharmacy',
        'pharmacy technician i': 'Pharmacy',
        'pharmacy assistant': 'Pharmacy',
        'quality assurance': 'Quality Assurance',
        'qa manager': 'Quality Assurance',
        'qa coordinator i': 'Quality Assurance',
        'qa coordinator ii': 'Quality Assurance',
        'qa admin assistant i': 'Quality Assurance',
        'qa admin assistant ii': 'Quality Assurance',
        'laboratory manager': 'Laboratory Technician',
        'laboratory technician': 'Laboratory Technician',
        'laboratory technician i': 'Laboratory Technician',
        'laboratory technician ii': 'Laboratory Technician',
        'data entry': 'Data Entry',
        'data entry i': 'Data Entry',
        'data entry ii': 'Data Entry',
        'data entry manager': 'Data Entry',
        'data coordinator i': 'Data Entry',
        'regulatory specialist i': 'Regulatory Coordinator',
        'regulatory specialist ii': 'Regulatory Coordinator',
        'admin asst (regulatory)': 'Regulatory Coordinator',
        'director of recruitment': 'Recruitment',
        'senior recruitment specialist': 'Recruitment',
        'recruitment specialist i': 'Recruitment',
        'recruitment specialist ii': 'Recruitment',
        'dietary aide': 'Dietary Aide',
        'principle investigator': 'Principal Investigator',
        'principal investigator': 'Principal Investigator',
        'pi': 'Principal Investigator',
        'sub-investigator': 'Sub-Investigator',
        'sub investigator': 'Sub-Investigator',
        'subinvestigator': 'Sub-Investigator',
        'subinvestigator': 'Sub-Investigator',
    };

    const RESP_SCROLL = {
        stepRatio: 0.7,
        userPauseMs: 800
    };

    const RESP_ATTRS = {
        ariaBusyTarget: 'body',
        ariaBusyAttr: 'aria-busy'
    };

    const CLEAN_SELECTORS = {
        mainPanelButtonTarget: '.main-gui-panel',
        ariaLiveRegion: '.aria-live-region'
    };

    const CLEAN_TIMEOUTS = {
        waitInputPanelMs: 10000,
        waitResultsPanelMs: 10000
    };

    const CLEAN_LABELS = {
        featureButton: 'Clean Study Task List',
        inputTitle: 'Clean Study Task List Input',
        resultsTitle: 'Cleaned Study Task List',
        responsibilitiesHeader: 'Responsibilities',
        confirm: 'Confirm',
        clear: 'Clear All',
        close: 'Close',
        downloadXlsx: 'Download .xlsx',
        downloadCsv: 'Download CSV',
        parsing: 'Parsing and cleaning input',
        done: 'Cleaning complete',
        exportSuccess: 'Export file created',
        exportFailed: 'Export failed'
    };

    const CLEAN_REGEX = {
        itemStart: /(^|\s)(\d{1,3})([.)])\s+/g,
        leadingToken: /^(\d{1,3})([.)])\s*/,
        specialChars: /[\\\/:<>"|?*]/g,
        standaloneAnd: /\b(and)\b/gi,
        smartQuotes: /[\u201c\u201d]/g,
        strayQuotes: /"+/g,
        whitespace: /\s+/g,
        duplicateCommas: /,\s*,+/g
    };

    const CLEAN_LIMITS = {
        maxChars: 100
    };

    const CLEAN_ATTRS = {
        ariaBusyTarget: 'body',
        ariaBusyAttr: 'aria-busy'
    };

    var cleanState = {
        isRunning: false,
        stopRequested: false,
        observers: [],
        timeouts: [],
        intervals: [],
        eventListeners: [],
        idleCallbackIds: [],
        focusReturnElement: null,
        prevAriaBusy: null,
        parsedItems: null,
        outputModel: null
    };

    var respState = {
        isRunning: false,
        stopRequested: false,
        observers: [],
        timeouts: [],
        intervals: [],
        eventListeners: [],
        idleCallbackIds: [],
        focusReturnElement: null,
        prevAriaBusy: null,
        parsedRoles: null,
        rolesData: [],
        counters: { total: 0, completed: 0, failed: 0, pending: 0 },
        listScrollTop: 0,
        userScrollHandler: null,
        userScrollPaused: false
    };

    function resetRespState() {
        addLogMessage('resetRespState: resetting state', 'log');
        respState.isRunning = false;
        respState.stopRequested = false;
        respState.observers = [];
        respState.timeouts = [];
        respState.intervals = [];
        respState.eventListeners = [];
        respState.idleCallbackIds = [];
        respState.prevAriaBusy = null;
        respState.parsedRoles = null;
        respState.rolesData = [];
        respState.counters = { total: 0, completed: 0, failed: 0, pending: 0 };
        respState.listScrollTop = 0;
        respState.userScrollHandler = null;
        respState.userScrollPaused = false;
    }

    function respWaitForElement(selector, timeout) {
        addLogMessage('respWaitForElement: waiting for ' + selector, 'log');
        return new Promise(function(resolve, reject) {
            var element = document.querySelector(selector);
            if (element) {
                addLogMessage('respWaitForElement: found immediately', 'log');
                resolve(element);
                return;
            }
            var observer = new MutationObserver(function(mutations, obs) {
                var el = document.querySelector(selector);
                if (el) {
                    obs.disconnect();
                    var idx = respState.observers.indexOf(obs);
                    if (idx > -1) {
                        respState.observers.splice(idx, 1);
                    }
                    resolve(el);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            respState.observers.push(observer);
            var timeoutId = setTimeout(function() {
                observer.disconnect();
                var idx = respState.observers.indexOf(observer);
                if (idx > -1) {
                    respState.observers.splice(idx, 1);
                }
                addLogMessage('respWaitForElement: timeout for ' + selector, 'warn');
                reject(new Error('Timeout waiting for ' + selector));
            }, timeout);
            respState.timeouts.push(timeoutId);
        });
    }

    function respDelay(ms) {
        return new Promise(function(resolve) {
            var tid = setTimeout(resolve, ms);
            respState.timeouts.push(tid);
        });
    }

    function normalizeRoleName(s) {
        if (!s) {
            return { display: '', key: '' };
        }
        var cleaned = s.trim();
        cleaned = cleaned.replace(RESP_REGEX.quoteCleanup, '');
        cleaned = cleaned.replace(RESP_REGEX.hyphenBreak, '-');
        cleaned = cleaned.replace(RESP_REGEX.lineBreakInRole, ' ');
        cleaned = cleaned.replace(RESP_REGEX.whitespace, ' ');
        cleaned = cleaned.trim();
        var display = cleaned;
        var key = cleaned.toLowerCase();
        if (RESP_ROLE_ALIASES[key]) {
            display = RESP_ROLE_ALIASES[key];
            key = display.toLowerCase();
        }
        addLogMessage('normalizeRoleName: display=' + display + ' key=' + key, 'log');
        return { display: display, key: key };
    }

    function expandRanges(tokens) {
        addLogMessage('expandRanges: tokens=' + JSON.stringify(tokens), 'log');
        var result = new Set();
        var i = 0;
        while (i < tokens.length) {
            var current = tokens[i];
            if (i + 2 < tokens.length && /^(to)$/i.test(tokens[i + 1])) {
                var start = parseInt(current, 10);
                var end = parseInt(tokens[i + 2], 10);
                if (!isNaN(start) && !isNaN(end)) {
                    var lo = Math.min(start, end);
                    var hi = Math.max(start, end);
                    for (var n = lo; n <= hi; n++) {
                        result.add(n);
                    }
                    i += 3;
                    continue;
                }
            }
            var num = parseInt(current, 10);
            if (!isNaN(num)) {
                result.add(num);
            }
            i++;
        }
        addLogMessage('expandRanges: result size=' + result.size, 'log');
        return result;
    }

    function parseResponsibilitiesInput(rawText) {
        addLogMessage('parseResponsibilitiesInput: starting parse', 'log');
        try {
            if (!rawText || !rawText.trim()) {
                addLogMessage('parseResponsibilitiesInput: empty input', 'warn');
                return null;
            }
            var text = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            text = text.replace(/[\u201c\u201d\u201e\u201f\u2018\u2019]+/g, '"');
            text = text.replace(/-\s*\n\s*/g, '-');
            var rawLines = text.split('\n');
            var mergedLines = [];
            var pendingQuote = false;
            var buffer = '';
            for (var li = 0; li < rawLines.length; li++) {
                var line = rawLines[li];
                if (pendingQuote) {
                    buffer = buffer + ' ' + line;
                    var quoteCount = 0;
                    for (var ci = 0; ci < buffer.length; ci++) {
                        if (buffer[ci] === '"') {
                            quoteCount++;
                        }
                    }
                    if (quoteCount % 2 === 0) {
                        pendingQuote = false;
                        mergedLines.push(buffer);
                        buffer = '';
                    }
                    continue;
                }
                var qc = 0;
                for (var ci2 = 0; ci2 < line.length; ci2++) {
                    if (line[ci2] === '"') {
                        qc++;
                    }
                }
                if (qc % 2 !== 0) {
                    pendingQuote = true;
                    buffer = line;
                    continue;
                }
                mergedLines.push(line);
            }
            if (buffer) {
                mergedLines.push(buffer);
            }
            addLogMessage('parseResponsibilitiesInput: merged into ' + mergedLines.length + ' lines', 'log');
            var parsedMap = {};
            for (var mi = 0; mi < mergedLines.length; mi++) {
                var mline = mergedLines[mi].trim();
                if (!mline) {
                    continue;
                }
                mline = mline.replace(/"+/g, '');
                var parts = mline.split(/\t+/);
                var rolePart = '';
                var numberPart = '';
                if (parts.length >= 2) {
                    var emailIdx = -1;
                    for (var pi = 0; pi < parts.length; pi++) {
                        if (parts[pi].indexOf('@') !== -1) {
                            emailIdx = pi;
                            break;
                        }
                    }
                    if (emailIdx !== -1 && emailIdx + 1 < parts.length) {
                        rolePart = parts[emailIdx + 1].trim();
                        numberPart = parts.slice(emailIdx + 2).join(' ');
                        addLogMessage('parseResponsibilitiesInput: line ' + mi + ' staff table detected, role=' + rolePart, 'log');
                    } else {
                        rolePart = parts[0].trim();
                        numberPart = parts.slice(1).join(' ');
                    }
                } else {
                    var firstNumMatch = mline.match(/\d/);
                    if (firstNumMatch) {
                        var idx = mline.indexOf(firstNumMatch[0]);
                        var beforeNum = mline.substring(0, idx).trim();
                        var afterNum = mline.substring(idx).trim();
                        if (beforeNum) {
                            rolePart = beforeNum;
                            numberPart = afterNum;
                        } else {
                            addLogMessage('parseResponsibilitiesInput: line ' + mi + ' no role part, skip', 'warn');
                            continue;
                        }
                    } else {
                        addLogMessage('parseResponsibilitiesInput: line ' + mi + ' no numbers, skip', 'warn');
                        continue;
                    }
                }
                rolePart = rolePart.replace(/"+/g, '').trim();
                if (!rolePart) {
                    addLogMessage('parseResponsibilitiesInput: line ' + mi + ' empty role, skip', 'warn');
                    continue;
                }
                var normalized = normalizeRoleName(rolePart);
                if (!normalized.key) {
                    addLogMessage('parseResponsibilitiesInput: line ' + mi + ' normalize fail, skip', 'warn');
                    continue;
                }
                var numTokens = numberPart.replace(/,/g, ' ').split(/\s+/).filter(function(t) {
                    return t.length > 0;
                });
                var numbers = expandRanges(numTokens);
                if (numbers.size === 0) {
                    addLogMessage('parseResponsibilitiesInput: line ' + mi + ' no numbers for ' + normalized.display, 'warn');
                    continue;
                }
                if (!parsedMap[normalized.key]) {
                    parsedMap[normalized.key] = { displayRole: normalized.display, occurrences: [], union: new Set(), intersection: null, excluded: new Set() };
                }
                parsedMap[normalized.key].occurrences.push(numbers);
                addLogMessage('parseResponsibilitiesInput: role=' + normalized.display + ' occ#' + parsedMap[normalized.key].occurrences.length + ' nums=' + Array.from(numbers).join(','), 'log');
            }
            addLogMessage('parseResponsibilitiesInput: parsed ' + Object.keys(parsedMap).length + ' unique roles', 'log');
            return parsedMap;
        } catch (error) {
            addLogMessage('parseResponsibilitiesInput: error: ' + error.message, 'error');
            return null;
        }
    }

    function computeRoleCommonAndExcluded(parsedMap) {
        addLogMessage('computeRoleCommonAndExcluded: computing sets', 'log');
        var rolesData = [];
        var keys = Object.keys(parsedMap);
        for (var ki = 0; ki < keys.length; ki++) {
            var key = keys[ki];
            var entry = parsedMap[key];
            var union = new Set();
            var intersection = null;
            for (var oi = 0; oi < entry.occurrences.length; oi++) {
                var occ = entry.occurrences[oi];
                occ.forEach(function(n) {
                    union.add(n);
                });
                if (intersection === null) {
                    intersection = new Set(occ);
                } else {
                    var newInt = new Set();
                    intersection.forEach(function(n) {
                        if (occ.has(n)) {
                            newInt.add(n);
                        }
                    });
                    intersection = newInt;
                }
            }
            if (intersection === null) {
                intersection = new Set();
            }
            var excluded = new Set();
            union.forEach(function(n) {
                if (!intersection.has(n)) {
                    excluded.add(n);
                }
            });
            entry.union = union;
            entry.intersection = intersection;
            entry.excluded = excluded;
            var status = RESP_LABELS.statusPending;
            if (intersection.size === 0) {
                status = RESP_LABELS.statusFailed;
                addLogMessage('computeRoleCommonAndExcluded: role=' + entry.displayRole + ' empty intersection', 'warn');
            }
            rolesData.push({
                key: key,
                displayRole: entry.displayRole,
                common: Array.from(intersection).sort(function(a, b) {
                    return a - b;
                }),
                excluded: Array.from(excluded).sort(function(a, b) {
                    return a - b;
                }),
                intersection: intersection,
                status: status,
                reason: intersection.size === 0 ? 'No common numbers across occurrences' : ''
            });
            addLogMessage('computeRoleCommonAndExcluded: role=' + entry.displayRole + ' common=[' + rolesData[rolesData.length - 1].common.join(',') + '] excluded=[' + rolesData[rolesData.length - 1].excluded.join(',') + ']', 'log');
        }
        return rolesData;
    }

    function setResponsibilitiesInit() {
        addLogMessage('setResponsibilitiesInit: starting feature', 'log');
        respState.focusReturnElement = document.getElementById('resp-set-btn');
        resetRespState();
        var pageStep = document.querySelector(RESP_SELECTORS.pageStepRoot);
        addLogMessage('setResponsibilitiesInit: checking for page step root', 'log');
        if (!pageStep) {
            addLogMessage('setResponsibilitiesInit: not on Study Role Page', 'warn');
            showRespWarning();
            return;
        }
        addLogMessage('setResponsibilitiesInit: on Study Role Page', 'log');
        showResponsibilitiesInputPanel();
    }

    function showRespWarning() {
        addLogMessage('showRespWarning: creating warning popup', 'log');
        var modal = document.createElement('div');
        modal.id = 'resp-warning-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 30000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); border-radius: 12px; padding: 24px; width: 450px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'alertdialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'resp-warning-title');
        container.setAttribute('aria-describedby', 'resp-warning-message');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;';
        var title = document.createElement('h3');
        title.id = 'resp-warning-title';
        title.textContent = 'Study Role Page Not Found';
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close warning');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.3)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        var closeWarning = function() {
            addLogMessage('showRespWarning: closing warning', 'log');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            if (respState.focusReturnElement) {
                respState.focusReturnElement.focus();
            }
        };
        closeButton.onclick = closeWarning;
        header.appendChild(title);
        header.appendChild(closeButton);
        var messageDiv = document.createElement('p');
        messageDiv.id = 'resp-warning-message';
        messageDiv.textContent = RESP_LABELS.notOnPageWarning + ' Please navigate to the Study Roles step before using this feature.';
        messageDiv.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0; font-size: 14px; line-height: 1.5;';
        var okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: 2px solid rgba(255, 255, 255, 0.3); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.3s ease; margin-top: 20px; width: 100%;';
        okButton.onmouseover = function() {
            okButton.style.background = 'rgba(255, 255, 255, 0.3)';
        };
        okButton.onmouseout = function() {
            okButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        okButton.onclick = closeWarning;
        var keyHandler = function(e) {
            if (e.key === 'Escape') {
                closeWarning();
            }
        };
        document.addEventListener('keydown', keyHandler);
        respState.eventListeners.push({ element: document, type: 'keydown', handler: keyHandler });
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
        addLogMessage('showRespWarning: warning displayed', 'log');
    }

    function showResponsibilitiesInputPanel() {
        addLogMessage('showResponsibilitiesInputPanel: creating input panel', 'log');
        var modal = document.createElement('div');
        modal.id = 'resp-input-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 550px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'resp-input-title');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;';
        var titleEl = document.createElement('h3');
        titleEl.id = 'resp-input-title';
        titleEl.textContent = 'Set Responsibilities';
        titleEl.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600; letter-spacing: 0.2px;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close panel');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = function() {
            addLogMessage('showResponsibilitiesInputPanel: closed by user', 'warn');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            stopResponsibilities();
        };
        header.appendChild(titleEl);
        header.appendChild(closeButton);
        var description = document.createElement('p');
        description.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0 0 12px 0; font-size: 14px; line-height: 1.4;';
        description.append("Rules:");
        description.appendChild(document.createElement('br'));
        var lines = [
            'Make sure the Study Responsibilities Identifier are set to Numbers and NOT Letters',
            'Paste role-to-responsibility assignments below.',
            'Each line should contain the role name, followed by a tab, then the responsibility numbers.',
            'Ranges such as "1 to 8" are supported.',
            'Ensure that no existing Study Roles are already added on the page.',
            'After clicking Confirm, do not click anywhere else on the page, as this will impact the process.'
        ];

        for (var i = 0; i < lines.length; i++) {
            description.appendChild(document.createTextNode('• ' + lines[i]));
            if (i < lines.length - 1) {
                description.appendChild(document.createElement('br'));
            }
        }
        var textarea = document.createElement('textarea');
        textarea.id = 'resp-input-textarea';
        textarea.placeholder = 'PI  1 to 8  13  14  17  21\nStudy Coordinator  1 6 7 8 10 12 13 14 17 33';
        textarea.setAttribute('aria-label', 'Role responsibilities input');
        textarea.style.cssText = 'width: 100%; height: 200px; padding: 12px 14px; border: 2px solid rgba(255, 255, 255, 0.35); border-radius: 10px; background: rgba(255, 255, 255, 0.95); color: #1e293b; font-size: 14px; font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; resize: vertical; outline: none; transition: all 0.25s ease; box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset; box-sizing: border-box;';
        textarea.onfocus = function() {
            textarea.style.borderColor = '#8ea0ff';
            textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)';
        };
        textarea.onblur = function() {
            textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)';
            textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset';
        };
        var confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.disabled = true;
        confirmButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; opacity: 0.5;';
        textarea.oninput = function() {
            if (textarea.value.trim().length > 0) {
                confirmButton.disabled = false;
                confirmButton.style.opacity = '1';
                confirmButton.style.cursor = 'pointer';
            } else {
                confirmButton.disabled = true;
                confirmButton.style.opacity = '0.5';
                confirmButton.style.cursor = 'not-allowed';
            }
        };
        confirmButton.onmouseover = function() {
            if (!confirmButton.disabled) {
                confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
            }
        };
        confirmButton.onmouseout = function() {
            confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        confirmButton.onclick = function() {
            addLogMessage('showResponsibilitiesInputPanel: Confirm clicked', 'log');
            var parsedMap = parseResponsibilitiesInput(textarea.value);
            if (!parsedMap || Object.keys(parsedMap).length === 0) {
                addLogMessage('showResponsibilitiesInputPanel: no valid roles', 'warn');
                return;
            }
            respState.parsedRoles = parsedMap;
            var rd = computeRoleCommonAndExcluded(parsedMap);
            respState.rolesData = rd;
            addLogMessage('showResponsibilitiesInputPanel: parsed ' + rd.length + ' roles', 'log');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            showResponsibilitiesProgressPanel(rd);
            processRolesWorkflow(rd);
        };
        var clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease;';
        clearButton.onmouseover = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        clearButton.onmouseout = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        clearButton.onclick = function() {
            addLogMessage('showResponsibilitiesInputPanel: Clear All clicked', 'log');
            textarea.value = '';
            respState.parsedRoles = null;
            confirmButton.disabled = true;
            confirmButton.style.opacity = '0.5';
            confirmButton.style.cursor = 'not-allowed';
        };
        var buttonContainer = document.createElement('div');
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
        addLogMessage('showResponsibilitiesInputPanel: displayed', 'log');
    }

    function getRespBadgeColors(status) {
        if (status === RESP_LABELS.statusCompleted) {
            return { color: '#6bcf7f', bg: 'rgba(107, 207, 127, 0.2)' };
        }
        if (status === RESP_LABELS.statusFailed) {
            return { color: '#ff6b6b', bg: 'rgba(255, 107, 107, 0.2)' };
        }
        if (status === RESP_LABELS.statusStopped) {
            return { color: '#aaa', bg: 'rgba(170, 170, 170, 0.2)' };
        }
        return { color: '#ffd93d', bg: 'rgba(255, 217, 61, 0.2)' };
    }

    function createRespRoleRow(roleData, index) {
        var item = document.createElement('div');
        item.className = 'resp-role-item';
        item.setAttribute('data-role-key', roleData.key);
        item.style.cssText = 'display: flex; flex-direction: column; padding: 10px 12px; margin: 4px 0; background: rgba(255, 255, 255, 0.08); border-radius: 6px; transition: background 0.2s ease;';
        item.onmouseover = function() {
            item.style.background = 'rgba(255, 255, 255, 0.12)';
        };
        item.onmouseout = function() {
            item.style.background = 'rgba(255, 255, 255, 0.08)';
        };
        var topRow = document.createElement('div');
        topRow.style.cssText = 'display: flex; justify-content: space-between; align-items: center;';
        var leftSection = document.createElement('div');
        leftSection.style.cssText = 'display: flex; align-items: center; gap: 8px; flex: 1; min-width: 0;';
        var indexBadge = document.createElement('span');
        indexBadge.textContent = String(index + 1);
        indexBadge.style.cssText = 'background: rgba(255, 255, 255, 0.15); color: rgba(255, 255, 255, 0.7); font-size: 11px; font-weight: 600; min-width: 24px; height: 24px; border-radius: 12px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;';
        var nameText = document.createElement('span');
        nameText.textContent = roleData.displayRole;
        nameText.style.cssText = 'color: white; font-size: 13px; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;';
        leftSection.appendChild(indexBadge);
        leftSection.appendChild(nameText);
        var statusBadge = document.createElement('span');
        statusBadge.className = 'resp-status-badge';
        statusBadge.textContent = roleData.status;
        var badgeColors = getRespBadgeColors(roleData.status);
        statusBadge.style.cssText = 'color: ' + badgeColors.color + '; background: ' + badgeColors.bg + '; font-size: 11px; font-weight: 600; padding: 3px 8px; border-radius: 10px; white-space: nowrap; flex-shrink: 0;';
        topRow.appendChild(leftSection);
        topRow.appendChild(statusBadge);
        var detailRow = document.createElement('div');
        detailRow.className = 'resp-detail-row';
        detailRow.style.cssText = 'margin-top: 6px; font-size: 11px; color: rgba(255, 255, 255, 0.7); line-height: 1.4;';
        var commonLabel = document.createElement('div');
        commonLabel.innerHTML = '<strong style="color: rgba(255,255,255,0.85);">Common:</strong> ' + (roleData.common.length > 0 ? roleData.common.join(', ') : 'None');
        detailRow.appendChild(commonLabel);
        if (roleData.excluded.length > 0) {
            var excludedLabel = document.createElement('div');
            excludedLabel.innerHTML = '<strong style="color: rgba(255,255,255,0.85);">Excluded:</strong> ' + roleData.excluded.join(', ');
            detailRow.appendChild(excludedLabel);
        }
        if (roleData.reason) {
            var reasonLabel = document.createElement('div');
            reasonLabel.style.cssText = 'color: #ff6b6b; margin-top: 2px;';
            reasonLabel.textContent = roleData.reason;
            detailRow.appendChild(reasonLabel);
        }
        var liveRegion = document.createElement('span');
        liveRegion.setAttribute('aria-live', 'polite');
        liveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        item.appendChild(topRow);
        item.appendChild(detailRow);
        item.appendChild(liveRegion);
        return item;
    }
function showResponsibilitiesProgressPanel(rolesData) {
    addLogMessage('showResponsibilitiesProgressPanel: creating', 'log');
    respState.isRunning = true;
    respState.stopRequested = false;
    respState.counters = { total: rolesData.length, completed: 0, failed: 0, pending: 0 };
    for (var ci = 0; ci < rolesData.length; ci++) {
        if (rolesData[ci].status === RESP_LABELS.statusPending) {
            respState.counters.pending++;
        } else if (rolesData[ci].status === RESP_LABELS.statusFailed) {
            respState.counters.failed++;
        } else if (rolesData[ci].status === RESP_LABELS.statusCompleted) {
            respState.counters.completed++;
        }
    }
    var modal = document.createElement('div');
    modal.id = 'resp-progress-modal';
    modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
    var container = document.createElement('div');
    container.id = 'resp-progress-container';
    container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 750px; max-width: 95%; max-height: 80vh; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative; display: flex; flex-direction: column;';
    container.setAttribute('role', 'dialog');
    container.setAttribute('aria-modal', 'true');
    container.setAttribute('aria-labelledby', 'resp-progress-title');
    var header = document.createElement('div');
    header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0; order: 0;';
    var titleContainer = document.createElement('div');
    titleContainer.style.cssText = 'display: flex; align-items: center; gap: 12px;';
    var title = document.createElement('h3');
    title.id = 'resp-progress-title';
    title.textContent = 'Set Responsibilities - Processing';
    title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
    var statusBadge = document.createElement('span');
    statusBadge.id = 'resp-status-badge';
    statusBadge.textContent = 'In Progress';
    statusBadge.style.cssText = 'background: rgba(255, 255, 255, 0.3); color: #ffd93d; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
    titleContainer.appendChild(title);
    titleContainer.appendChild(statusBadge);
    var closeButton = document.createElement('button');
    closeButton.innerHTML = '\u2715';
    closeButton.setAttribute('aria-label', 'Close and stop');
    closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
    closeButton.onmouseover = function() {
        closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
    };
    closeButton.onmouseout = function() {
        closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
    };
    closeButton.onclick = function() {
        addLogMessage('showResponsibilitiesProgressPanel: closed', 'warn');
        stopResponsibilities();
    };
    header.appendChild(titleContainer);
    header.appendChild(closeButton);
    container.appendChild(header);
    addLogMessage('showResponsibilitiesProgressPanel: header appended', 'log');

    var descriptionContainer = document.createElement('div');
    descriptionContainer.style.cssText = 'margin-bottom: 12px; padding: 8px 0; border-bottom: 1px solid rgba(255, 255, 255, 0.1);';
    var description = document.createElement('div');
    var bulletPoints = [
        'Do not click anywhere on the page or outside of the page. It will affect the process as doing so closes the dropdown menu.',
        'If you need to make changes, do it at the end of the process: delete the role and create another one.'
    ];
    description.innerHTML = bulletPoints.map(function(point) {
        return '• ' + point;
    }).join('<br>');
    description.style.cssText = 'color: rgba(255, 255, 255, 0.85); font-size: 13px; line-height: 1.4; font-weight: 400;';
    descriptionContainer.appendChild(description);
    container.appendChild(descriptionContainer);
    addLogMessage('showResponsibilitiesProgressPanel: description appended', 'log');

    var searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.id = 'resp-progress-search';
    searchInput.placeholder = 'Search roles...';
    searchInput.setAttribute('aria-label', 'Search roles');
    searchInput.style.cssText = 'width: 100%; padding: 8px 12px; border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 6px; background: rgba(255, 255, 255, 0.1); color: white; font-size: 13px; outline: none; box-sizing: border-box; margin-bottom: 12px; flex-shrink: 0;';
    searchInput.oninput = function() {
        var term = searchInput.value.toLowerCase().trim();
        var items = document.querySelectorAll('.resp-role-item');
        for (var si = 0; si < items.length; si++) {
            var text = items[si].textContent.toLowerCase();
            if (!term || text.indexOf(term) !== -1) {
                items[si].style.display = 'flex';
            } else {
                items[si].style.display = 'none';
            }
        }
    };
    container.appendChild(searchInput);
    addLogMessage('showResponsibilitiesProgressPanel: search input appended', 'log');

    var listContainer = document.createElement('div');
    listContainer.id = 'resp-roles-list';
    listContainer.style.cssText = 'flex: 1; overflow-y: auto; min-height: 150px; max-height: 400px;';
    for (var ri = 0; ri < rolesData.length; ri++) {
        listContainer.appendChild(createRespRoleRow(rolesData[ri], ri));
    }
    container.appendChild(listContainer);
    addLogMessage('showResponsibilitiesProgressPanel: list appended', 'log');

    var summaryFooter = document.createElement('div');
    summaryFooter.id = 'resp-summary-footer';
    summaryFooter.style.cssText = 'display: flex; justify-content: space-around; align-items: center; padding: 10px 16px; background: rgba(0, 0, 0, 0.2); border-radius: 8px; margin-top: 12px; flex-shrink: 0;';
    var summaryItems = [
        { id: 'resp-summary-total', label: 'Total', value: String(respState.counters.total) },
        { id: 'resp-summary-completed', label: 'Completed', value: String(respState.counters.completed) },
        { id: 'resp-summary-failed', label: 'Failed', value: String(respState.counters.failed) },
        { id: 'resp-summary-pending', label: 'Pending', value: String(respState.counters.pending) },
        { id: 'resp-summary-percent', label: 'Progress', value: '0%' }
    ];
    for (var si2 = 0; si2 < summaryItems.length; si2++) {
        var sItem = document.createElement('div');
        sItem.style.cssText = 'text-align: center;';
        var vSpan = document.createElement('span');
        vSpan.id = summaryItems[si2].id;
        vSpan.textContent = summaryItems[si2].value;
        vSpan.style.cssText = 'display: block; color: white; font-size: 16px; font-weight: 700;';
        var lSpan = document.createElement('span');
        lSpan.textContent = summaryItems[si2].label;
        lSpan.style.cssText = 'display: block; color: rgba(255, 255, 255, 0.6); font-size: 11px; font-weight: 500; margin-top: 2px;';
        sItem.appendChild(vSpan);
        sItem.appendChild(lSpan);
        summaryFooter.appendChild(sItem);
    }
    container.appendChild(summaryFooter);
    addLogMessage('showResponsibilitiesProgressPanel: summary appended', 'log');

    var ariaLiveRegion = document.createElement('div');
    ariaLiveRegion.id = 'resp-aria-live';
    ariaLiveRegion.setAttribute('aria-live', 'polite');
    ariaLiveRegion.setAttribute('aria-atomic', 'true');
    ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
    container.appendChild(ariaLiveRegion);
    addLogMessage('showResponsibilitiesProgressPanel: aria-live appended', 'log');

    modal.appendChild(container);
    container.style.position = 'fixed';
    container.style.top = '50%';
    container.style.left = '50%';
    container.style.transform = 'translate(-50%, -50%)';
    modal.style.pointerEvents = 'none';
    container.style.pointerEvents = 'auto';
    makeDraggable(container, header);
    document.body.appendChild(modal);
    closeButton.focus();
    addLogMessage('showResponsibilitiesProgressPanel: displayed', 'log');
}


    function updateRespRoleStatus(roleKey, newStatus, reason) {
        addLogMessage('updateRespRoleStatus: role=' + roleKey + ' status=' + newStatus, 'log');
        var items = document.querySelectorAll('.resp-role-item');
        for (var i = 0; i < items.length; i++) {
            if (items[i].getAttribute('data-role-key') === roleKey) {
                var badge = items[i].querySelector('.resp-status-badge');
                if (badge) {
                    badge.textContent = newStatus;
                    var colors = getRespBadgeColors(newStatus);
                    badge.style.color = colors.color;
                    badge.style.background = colors.bg;
                }
                if (reason) {
                    var dr = items[i].querySelector('.resp-detail-row');
                    if (dr) {
                        var re = document.createElement('div');
                        re.style.cssText = 'color: #ff6b6b; margin-top: 2px;';
                        re.textContent = reason;
                        dr.appendChild(re);
                    }
                }
                var lr = items[i].querySelector('[aria-live]');
                if (lr) {
                    lr.textContent = roleKey + ' ' + newStatus;
                }
                break;
            }
        }
    }

    function updateRespSummary() {
        addLogMessage('updateRespSummary: t=' + respState.counters.total + ' c=' + respState.counters.completed + ' f=' + respState.counters.failed + ' p=' + respState.counters.pending, 'log');
        var el1 = document.getElementById('resp-summary-total');
        var el2 = document.getElementById('resp-summary-completed');
        var el3 = document.getElementById('resp-summary-failed');
        var el4 = document.getElementById('resp-summary-pending');
        var el5 = document.getElementById('resp-summary-percent');
        if (el1) {
            el1.textContent = String(respState.counters.total);
        }
        if (el2) {
            el2.textContent = String(respState.counters.completed);
        }
        if (el3) {
            el3.textContent = String(respState.counters.failed);
        }
        if (el4) {
            el4.textContent = String(respState.counters.pending);
        }
        if (el5) {
            var processed = respState.counters.completed + respState.counters.failed;
            var pct = respState.counters.total > 0 ? Math.round((processed / respState.counters.total) * 100) : 0;
            el5.textContent = pct + '%';
        }
        updateRespAriaLive('Completed: ' + respState.counters.completed + ', Failed: ' + respState.counters.failed + ', Pending: ' + respState.counters.pending);
    }

    function updateRespAriaLive(message) {
        var lr = document.getElementById('resp-aria-live');
        if (lr) {
            lr.textContent = message;
        }
    }

    function scanExistingCompletedRoles() {
        addLogMessage('scanExistingCompletedRoles: scanning', 'log');
        try {
            var dropList = document.querySelector(RESP_SELECTORS.dropListContainer);
            if (!dropList) {
                addLogMessage('scanExistingCompletedRoles: no drop list container', 'warn');
                return;
            }
            var columns = dropList.querySelectorAll(RESP_SELECTORS.roleColumns);
            addLogMessage('scanExistingCompletedRoles: ' + columns.length + ' columns', 'log');
            for (var colIdx = 0; colIdx < columns.length; colIdx++) {
                var checked = columns[colIdx].querySelectorAll(RESP_SELECTORS.selectedRoleCheckboxInColumn);
                for (var bi = 0; bi < checked.length; bi++) {
                    var ariaLabel = checked[bi].getAttribute('aria-label') || '';
                    var parentItem = checked[bi].closest('.filtered-select__list__item');
                    var roleText = ariaLabel || (parentItem ? parentItem.textContent || '' : '');
                    if (!roleText) {
                        continue;
                    }
                    var normalized = normalizeRoleName(roleText);
                    for (var ri = 0; ri < respState.rolesData.length; ri++) {
                        if (respState.rolesData[ri].key === normalized.key && respState.rolesData[ri].status === RESP_LABELS.statusPending) {
                            respState.rolesData[ri].status = RESP_LABELS.statusCompleted;
                            respState.counters.completed++;
                            respState.counters.pending--;
                            updateRespRoleStatus(respState.rolesData[ri].key, RESP_LABELS.statusCompleted, '');
                            addLogMessage('scanExistingCompletedRoles: ' + normalized.display + ' already done', 'log');
                        }
                    }
                }
            }
            updateRespSummary();
            updateRespAriaLive('Existing roles scanned');
        } catch (error) {
            addLogMessage('scanExistingCompletedRoles: error: ' + error.message, 'error');
        }
    }

    function findAvailableRoleColumn() {
        addLogMessage('findAvailableRoleColumn: searching for last empty column', 'log');
        return new Promise(function(resolve, reject) {
            if (respState.stopRequested) {
                reject(new Error('Stopped'));
                return;
            }
            var dropList = document.querySelector(RESP_SELECTORS.dropListContainer);
            if (!dropList) {
                addLogMessage('findAvailableRoleColumn: drop list container not found', 'warn');
                reject(new Error('Drop list container not found'));
                return;
            }
            var columns = dropList.querySelectorAll(RESP_SELECTORS.roleColumns);
            addLogMessage('findAvailableRoleColumn: found ' + columns.length + ' columns in drop list', 'log');
            var lastEmpty = null;
            for (var ci = columns.length - 1; ci >= 0; ci--) {
                var selectedCheckbox = columns[ci].querySelector(RESP_SELECTORS.selectedRoleCheckboxInColumn);
                if (!selectedCheckbox) {
                    lastEmpty = columns[ci];
                    addLogMessage('findAvailableRoleColumn: last empty column at index ' + ci, 'log');
                    break;
                }
            }
            if (lastEmpty) {
                resolve(lastEmpty);
                return;
            }
            addLogMessage('findAvailableRoleColumn: no empty columns, clicking Add Study Role', 'log');
            var addBtn = document.querySelector(RESP_SELECTORS.addStudyRoleBtn);
            if (!addBtn) {
                reject(new Error('Add Study Role button not found'));
                return;
            }
            addBtn.click();
            var waitTid = setTimeout(function() {
                var updatedColumns = dropList.querySelectorAll(RESP_SELECTORS.roleColumns);
                addLogMessage('findAvailableRoleColumn: after add, ' + updatedColumns.length + ' columns', 'log');
                if (updatedColumns.length === 0) {
                    reject(new Error('No columns found after adding study role'));
                    return;
                }
                var lastCol = updatedColumns[updatedColumns.length - 1];
                var sel = lastCol.querySelector(RESP_SELECTORS.selectedRoleCheckboxInColumn);
                if (!sel) {
                    addLogMessage('findAvailableRoleColumn: new last column at index ' + (updatedColumns.length - 1), 'log');
                    resolve(lastCol);
                    return;
                }
                reject(new Error('No empty column after adding study role'));
            }, RESP_TIMEOUTS.waitAddRoleResultMs);
            respState.timeouts.push(waitTid);
        });
    }

    function ensureRoleListOpenForColumn(columnEl) {
        addLogMessage('ensureRoleListOpenForColumn: ensuring open in target column', 'log');
        return new Promise(function(resolve, reject) {
            var retries = 0;
            function tryOpen() {
                if (respState.stopRequested) {
                    reject(new Error('Stopped'));
                    return;
                }
                addLogMessage('ensureRoleListOpenForColumn: attempt ' + (retries + 1), 'log');
                var globalList = document.querySelector(RESP_SELECTORS.roleListContainer);
                if (globalList) {
                    addLogMessage('ensureRoleListOpenForColumn: global list already open', 'log');
                    resolve(globalList);
                    return;
                }
                var inputEl = columnEl.querySelector(RESP_SELECTORS.roleSearchInput);
                if (!inputEl) {
                    inputEl = columnEl.querySelector('.filtered-select__input, input[placeholder*="Search"]');
                }
                if (!inputEl) {
                    var allInputs = columnEl.querySelectorAll('input');
                    if (allInputs.length > 0) {
                        inputEl = allInputs[0];
                    }
                }
                if (inputEl) {
                    addLogMessage('ensureRoleListOpenForColumn: clicking input in column', 'log');
                    inputEl.click();
                    inputEl.focus();
                } else {
                    addLogMessage('ensureRoleListOpenForColumn: no input found in column, clicking column itself', 'warn');
                    columnEl.click();
                }
                var waitStart = Date.now();
                function pollForList() {
                    var el = document.querySelector(RESP_SELECTORS.roleListContainer);
                    if (el) {
                        addLogMessage('ensureRoleListOpenForColumn: global list appeared', 'log');
                        resolve(el);
                        return;
                    }
                    var vp = document.querySelector('cdk-virtual-scroll-viewport');
                    if (vp && vp.scrollHeight > 0 && vp.clientHeight > 0) {
                        addLogMessage('ensureRoleListOpenForColumn: found active viewport as list proxy', 'log');
                        resolve(vp);
                        return;
                    }
                    if (Date.now() - waitStart > RESP_TIMEOUTS.waitListOpenMs) {
                        retries++;
                        if (retries < RESP_RETRY.openListRetries) {
                            var tid = setTimeout(tryOpen, 300);
                            respState.timeouts.push(tid);
                        } else {
                            reject(new Error('Could not open role list in column'));
                        }
                        return;
                    }
                    var pt = setTimeout(pollForList, 200);
                    respState.timeouts.push(pt);
                }
                pollForList();
            }
            tryOpen();
        });
    }

    function findRoleListViewport() {
        var allViewports = document.querySelectorAll('cdk-virtual-scroll-viewport');
        addLogMessage('findRoleListViewport: found ' + allViewports.length + ' global cdk-virtual-scroll-viewport elements', 'log');
        if (allViewports.length > 0) {
            var best = null;
            for (var vi = allViewports.length - 1; vi >= 0; vi--) {
                var vp = allViewports[vi];
                if (vp.scrollHeight > 0 && vp.clientHeight > 0) {
                    best = vp;
                    addLogMessage('findRoleListViewport: using viewport index ' + vi + ' scrollH=' + vp.scrollHeight + ' clientH=' + vp.clientHeight, 'log');
                    break;
                }
            }
            if (best) {
                return best;
            }
            var last = allViewports[allViewports.length - 1];
            addLogMessage('findRoleListViewport: fallback to last viewport index ' + (allViewports.length - 1) + ' scrollH=' + last.scrollHeight + ' clientH=' + last.clientHeight, 'log');
            return last;
        }
        var listEl = document.querySelector(RESP_SELECTORS.roleListContainer);
        if (listEl) {
            addLogMessage('findRoleListViewport: fallback to global list container scrollH=' + listEl.scrollHeight + ' clientH=' + listEl.clientHeight, 'log');
            return listEl;
        }
        addLogMessage('findRoleListViewport: no viewport found anywhere', 'warn');
        return null;
    }

    function respRememberListScrollPosition(viewportEl) {
        if (viewportEl) {
            respState.listScrollTop = viewportEl.scrollTop;
            addLogMessage('respRememberScroll: ' + respState.listScrollTop, 'log');
        }
    }

    function respRestoreListScrollPosition(viewportEl) {
        if (viewportEl && respState.listScrollTop > 0) {
            viewportEl.scrollTop = respState.listScrollTop;
            addLogMessage('respRestoreScroll: ' + respState.listScrollTop, 'log');
        }
    }

    function attemptSelectByScrollingForRole(columnEl, targetRoleDisplay, targetRoleKey) {
        addLogMessage('attemptSelectByScrollingForRole: target=' + targetRoleDisplay + ' key=' + targetRoleKey, 'log');
        return new Promise(function(resolve) {
            var viewportEl = findRoleListViewport();
            if (!viewportEl) {
                addLogMessage('attemptSelectByScrollingForRole: no viewport found', 'error');
                resolve(false);
                return;
            }
            addLogMessage('attemptSelectByScrollingForRole: viewport scrollH=' + viewportEl.scrollHeight + ' clientH=' + viewportEl.clientHeight, 'log');
            viewportEl.scrollTop = 0;
            var passCount = 0;
            var lastScrollTop = -1;
            var lastSnapshot = '';
            var startTime = Date.now();

            function scanAndScroll() {
                if (respState.stopRequested) {
                    respRememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                if (Date.now() - startTime > RESP_TIMEOUTS.maxSelectRoleDurationMs) {
                    addLogMessage('attemptSelectByScrollingForRole: max duration exceeded', 'warn');
                    respRememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                if (passCount >= RESP_RETRY.maxScrollPasses) {
                    addLogMessage('attemptSelectByScrollingForRole: max passes reached (' + passCount + ')', 'warn');
                    respRememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                var options = document.querySelectorAll(RESP_SELECTORS.roleOptionItem);
                var retryCount = 0;

                function tryScan() {
                    options = document.querySelectorAll(RESP_SELECTORS.roleOptionItem);
                    if (options.length === 0 && retryCount < RESP_RETRY.optionScanRetries) {
                        retryCount++;
                        addLogMessage('attemptSelectByScrollingForRole: empty render, retry ' + retryCount, 'log');
                        var rtid = setTimeout(tryScan, RESP_TIMEOUTS.waitOptionRenderMs);
                        respState.timeouts.push(rtid);
                        return;
                    }
                    addLogMessage('attemptSelectByScrollingForRole: scanning ' + options.length + ' options at scrollTop=' + viewportEl.scrollTop, 'log');
                    var found = false;
                    var optTexts = [];
                    for (var oi = 0; oi < options.length; oi++) {
                        var textEl = options[oi].querySelector(RESP_SELECTORS.roleOptionText);
                        var optText = textEl ? textEl.textContent.trim() : (options[oi].textContent || '').trim();
                        optTexts.push(optText);
                        var norm = normalizeRoleName(optText);
                        if (norm.key === targetRoleKey) {
                            addLogMessage('attemptSelectByScrollingForRole: MATCH at index ' + oi + ' text=' + optText, 'log');
                            var cb = options[oi].querySelector(RESP_SELECTORS.roleOptionCheckboxTri);
                            if (cb) {
                                cb.click();
                                addLogMessage('attemptSelectByScrollingForRole: clicked checkbox tristate', 'log');
                            } else {
                                options[oi].click();
                                addLogMessage('attemptSelectByScrollingForRole: clicked option item directly', 'log');
                            }
                            found = true;
                            respRememberListScrollPosition(viewportEl);
                            var vt = setTimeout(function() {
                                addLogMessage('attemptSelectByScrollingForRole: closing role dropdown after selection', 'log');
                                document.body.click();
                                var closeTid = setTimeout(function() {
                                    resolve(true);
                                }, 300);
                                respState.timeouts.push(closeTid);
                            }, RESP_TIMEOUTS.waitAfterSelectRoleMs);
                            respState.timeouts.push(vt);
                            break;
                        }
                    }
                    if (!found) {
                        var snap = optTexts.join('|');
                        var curTop = viewportEl.scrollTop;
                        if (curTop === lastScrollTop && snap === lastSnapshot) {
                            passCount++;
                            addLogMessage('attemptSelectByScrollingForRole: no progress pass ' + passCount + ' opts=' + options.length, 'log');
                        }
                        lastScrollTop = curTop;
                        lastSnapshot = snap;
                        var step = Math.round(viewportEl.clientHeight * RESP_SCROLL.stepRatio);
                        if (step < 50) {
                            step = 200;
                            addLogMessage('attemptSelectByScrollingForRole: step too small, using fallback 200px', 'warn');
                        }
                        var maxS = viewportEl.scrollHeight - viewportEl.clientHeight;
                        var newTop = Math.min(curTop + step, maxS);
                        if (newTop <= curTop && curTop > 0) {
                            addLogMessage('attemptSelectByScrollingForRole: at bottom, wrapping to top', 'log');
                            newTop = 0;
                            lastScrollTop = -1;
                            lastSnapshot = '';
                            passCount++;
                        }
                        viewportEl.scrollTop = newTop;
                        addLogMessage('attemptSelectByScrollingForRole: scrolled to ' + newTop + ' (max=' + maxS + ' step=' + step + ')', 'log');
                        var st = setTimeout(scanAndScroll, RESP_TIMEOUTS.idleBetweenScrollsMs);
                        respState.timeouts.push(st);
                    }
                }
                tryScan();
            }
            var initTid = setTimeout(scanAndScroll, RESP_TIMEOUTS.settleAfterScrollMs);
            respState.timeouts.push(initTid);
        });
    }

    function dismissExistingResponsibilitiesMenu() {
        return new Promise(function(resolve) {
            var allMenus = document.querySelectorAll(RESP_SELECTORS.responsibilitiesMenu);
            if (allMenus.length === 0) {
                resolve();
                return;
            }
            addLogMessage('dismissExistingResponsibilitiesMenu: found ' + allMenus.length + ' open menu(s), closing', 'log');
            var allToggles = document.querySelectorAll(RESP_SELECTORS.responsibilitiesToggleBtn);
            for (var ti = 0; ti < allToggles.length; ti++) {
                var btn = allToggles[ti];
                var expanded = btn.getAttribute('aria-expanded');
                if (expanded === 'true') {
                    addLogMessage('dismissExistingResponsibilitiesMenu: clicking expanded toggle at index ' + ti, 'log');
                    btn.click();
                }
            }
            if (allToggles.length === 0) {
                document.body.click();
            }
            var attempts = 0;
            function waitGone() {
                var still = document.querySelectorAll(RESP_SELECTORS.responsibilitiesMenu);
                if (still.length === 0) {
                    addLogMessage('dismissExistingResponsibilitiesMenu: all menus closed', 'log');
                    resolve();
                    return;
                }
                attempts++;
                if (attempts >= 10) {
                    addLogMessage('dismissExistingResponsibilitiesMenu: menus still visible after retries, clicking body to close', 'warn');
                    document.body.click();
                    var finalTid = setTimeout(function() {
                        addLogMessage('dismissExistingResponsibilitiesMenu: proceeding after final close attempt', 'log');
                        resolve();
                    }, 300);
                    respState.timeouts.push(finalTid);
                    return;
                }
                if (attempts === 5) {
                    addLogMessage('dismissExistingResponsibilitiesMenu: mid-retry body click', 'log');
                    document.body.click();
                }
                var tid = setTimeout(waitGone, 200);
                respState.timeouts.push(tid);
            }
            var initTid = setTimeout(waitGone, 300);
            respState.timeouts.push(initTid);
        });
    }

    function openResponsibilitiesDropdown(columnEl) {
        addLogMessage('openResponsibilitiesDropdown: opening responsibilities dropdown', 'log');
        return dismissExistingResponsibilitiesMenu().then(function() {
            return new Promise(function(resolve, reject) {
                try {
                    var targetColumn = columnEl;
                    addLogMessage('openResponsibilitiesDropdown: using passed columnEl directly', 'log');
                    var toggleBtn = targetColumn.querySelector(RESP_SELECTORS.responsibilitiesToggleBtn);
                    if (!toggleBtn) {
                        toggleBtn = targetColumn.querySelector('button[dropdowntoggle]');
                    }
                    if (!toggleBtn) {
                        toggleBtn = targetColumn.querySelector('.roles__select-options-dropdown-button');
                    }
                    if (!toggleBtn) {
                        var allToggles = document.querySelectorAll(RESP_SELECTORS.responsibilitiesToggleBtn);
                        if (allToggles.length > 0) {
                            toggleBtn = allToggles[allToggles.length - 1];
                            addLogMessage('openResponsibilitiesDropdown: using last global toggle button', 'log');
                        }
                    }
                    if (!toggleBtn) {
                        addLogMessage('openResponsibilitiesDropdown: toggle button not found anywhere', 'error');
                        reject(new Error('Responsibilities toggle button not found'));
                        return;
                    }
                    var menusBefore = document.querySelectorAll(RESP_SELECTORS.responsibilitiesMenu).length;
                    addLogMessage('openResponsibilitiesDropdown: menus before click: ' + menusBefore + ', clicking toggle button', 'log');
                    toggleBtn.click();

                    function waitForLastMenu(timeoutMs, retryClick) {
                        var elapsed = 0;
                        var interval = 100;
                        function poll() {
                            var allMenus = document.querySelectorAll(RESP_SELECTORS.responsibilitiesMenu);
                            var lastMenu = allMenus.length > 0 ? allMenus[allMenus.length - 1] : null;
                            if (allMenus.length > menusBefore && lastMenu) {
                                addLogMessage('openResponsibilitiesDropdown: new menu appeared (count ' + menusBefore + ' -> ' + allMenus.length + ')', 'log');
                                resolve(lastMenu);
                                return;
                            }
                            if (lastMenu && allMenus.length >= 1) {
                                var isVisible = lastMenu.offsetParent !== null || lastMenu.offsetHeight > 0;
                                if (isVisible) {
                                    addLogMessage('openResponsibilitiesDropdown: last menu is visible (count=' + allMenus.length + ')', 'log');
                                    resolve(lastMenu);
                                    return;
                                }
                            }
                            elapsed += interval;
                            if (elapsed >= timeoutMs) {
                                if (retryClick) {
                                    addLogMessage('openResponsibilitiesDropdown: timeout, retrying click', 'warn');
                                    toggleBtn.click();
                                    waitForLastMenu(timeoutMs, false);
                                } else {
                                    if (lastMenu) {
                                        addLogMessage('openResponsibilitiesDropdown: timeout but using last menu', 'warn');
                                        resolve(lastMenu);
                                    } else {
                                        reject(new Error('Timeout waiting for responsibilities menu'));
                                    }
                                }
                                return;
                            }
                            var tid = setTimeout(poll, interval);
                            respState.timeouts.push(tid);
                        }
                        var initTid = setTimeout(poll, interval);
                        respState.timeouts.push(initTid);
                    }
                    waitForLastMenu(RESP_TIMEOUTS.waitResponsibilitiesMenuMs, true);
                } catch (error) {
                    addLogMessage('openResponsibilitiesDropdown: error: ' + error.message, 'error');
                    reject(error);
                }
            });
        });
    }

    function selectResponsibilitiesByNumbers(numbersSet, menuEl) {
        addLogMessage('selectResponsibilitiesByNumbers: selecting ' + numbersSet.size + ' numbers: [' + Array.from(numbersSet).join(',') + ']', 'log');
        return new Promise(function(resolve) {
            try {
                var remaining = new Set(numbersSet);
                var scrollPassCount = 0;
                var maxScrollPasses = 10;
                var lastSnapshot = '';

                function getMenuContainer() {
                    if (menuEl && menuEl.isConnected) {
                        return menuEl;
                    }
                    var allMenus = document.querySelectorAll(RESP_SELECTORS.responsibilitiesMenu);
                    return allMenus.length > 0 ? allMenus[allMenus.length - 1] : null;
                }

                function processVisibleItems(callback) {
                    var container = getMenuContainer();
                    var items = container ? container.querySelectorAll('li') : document.querySelectorAll(RESP_SELECTORS.responsibilitiesItem);
                    addLogMessage('selectResponsibilitiesByNumbers: ' + items.length + ' visible items, ' + remaining.size + ' remaining', 'log');
                    var idx = 0;

                    function nextItem() {
                        if (respState.stopRequested) {
                            callback();
                            return;
                        }
                        if (idx >= items.length) {
                            callback();
                            return;
                        }
                        var item = items[idx];
                        var text = (item.textContent || '').trim();
                        var numMatch = text.match(/(\d+)/);
                        if (numMatch) {
                            var num = parseInt(numMatch[1], 10);
                            if (remaining.has(num)) {
                                var checkbox = item.querySelector('input[type="checkbox"]');
                                if (checkbox && !checkbox.checked) {
                                    addLogMessage('selectResponsibilitiesByNumbers: toggling checkbox for ' + num, 'log');
                                    checkbox.click();
                                    remaining.delete(num);
                                    idx++;
                                    var tid = setTimeout(nextItem, RESP_TIMEOUTS.waitAfterToggleResponsibilityMs);
                                    respState.timeouts.push(tid);
                                    return;
                                } else if (checkbox && checkbox.checked) {
                                    addLogMessage('selectResponsibilitiesByNumbers: ' + num + ' already checked', 'log');
                                    remaining.delete(num);
                                } else if (!checkbox) {
                                    addLogMessage('selectResponsibilitiesByNumbers: no checkbox for ' + num + ', clicking item', 'log');
                                    item.click();
                                    remaining.delete(num);
                                    idx++;
                                    var tid2 = setTimeout(nextItem, RESP_TIMEOUTS.waitAfterToggleResponsibilityMs);
                                    respState.timeouts.push(tid2);
                                    return;
                                }
                            }
                        }
                        idx++;
                        var tid3 = setTimeout(nextItem, 20);
                        respState.timeouts.push(tid3);
                    }
                    nextItem();
                }

                function scrollAndProcess() {
                    if (respState.stopRequested || remaining.size === 0) {
                        addLogMessage('selectResponsibilitiesByNumbers: done, remaining=' + remaining.size, 'log');
                        resolve();
                        return;
                    }
                    if (scrollPassCount >= maxScrollPasses) {
                        addLogMessage('selectResponsibilitiesByNumbers: max scroll passes, remaining=' + remaining.size, 'warn');
                        resolve();
                        return;
                    }
                    processVisibleItems(function() {
                        if (remaining.size === 0) {
                            addLogMessage('selectResponsibilitiesByNumbers: all selected', 'log');
                            resolve();
                            return;
                        }
                        var menuEl = getMenuContainer();
                        if (!menuEl) {
                            addLogMessage('selectResponsibilitiesByNumbers: menu gone, resolving', 'warn');
                            resolve();
                            return;
                        }
                        var items = menuEl.querySelectorAll('li');
                        var snap = '';
                        for (var si = 0; si < items.length; si++) {
                            snap += (items[si].textContent || '').trim() + '|';
                        }
                        if (snap === lastSnapshot) {
                            scrollPassCount++;
                            addLogMessage('selectResponsibilitiesByNumbers: no new items pass ' + scrollPassCount, 'log');
                        }
                        lastSnapshot = snap;
                        var scrollStep = Math.round(menuEl.clientHeight * 0.7);
                        if (scrollStep < 50) {
                            scrollStep = 200;
                        }
                        var maxScroll = menuEl.scrollHeight - menuEl.clientHeight;
                        var newTop = Math.min(menuEl.scrollTop + scrollStep, maxScroll);
                        if (newTop <= menuEl.scrollTop && maxScroll > 0) {
                            addLogMessage('selectResponsibilitiesByNumbers: at bottom of menu, done scrolling', 'log');
                            resolve();
                            return;
                        }
                        menuEl.scrollTop = newTop;
                        addLogMessage('selectResponsibilitiesByNumbers: scrolled menu to ' + newTop + ' (max=' + maxScroll + ')', 'log');
                        scrollPassCount++;
                        var tid4 = setTimeout(scrollAndProcess, RESP_TIMEOUTS.settleAfterScrollMs);
                        respState.timeouts.push(tid4);
                    });
                }
                scrollAndProcess();
            } catch (error) {
                addLogMessage('selectResponsibilitiesByNumbers: error: ' + error.message, 'error');
                resolve();
            }
        });
    }

    function clickAddStudyRole() {
        addLogMessage('clickAddStudyRole: clicking', 'log');
        return new Promise(function(resolve) {
            var retries = 0;

            function tryClick() {
                if (respState.stopRequested) {
                    resolve(false);
                    return;
                }
                var addBtn = document.querySelector(RESP_SELECTORS.addStudyRoleBtn);
                if (!addBtn) {
                    addLogMessage('clickAddStudyRole: button not found', 'warn');
                    resolve(false);
                    return;
                }
                if (addBtn.disabled || addBtn.getAttribute('disabled') !== null) {
                    if (retries < RESP_RETRY.addRoleRetries) {
                        retries++;
                        addLogMessage('clickAddStudyRole: button disabled, retry ' + retries, 'log');
                        var tid = setTimeout(tryClick, 1000);
                        respState.timeouts.push(tid);
                        return;
                    }
                    addLogMessage('clickAddStudyRole: button still disabled after retries', 'warn');
                    resolve(false);
                    return;
                }
                addBtn.click();
                addLogMessage('clickAddStudyRole: clicked, waiting for result', 'log');
                var wt = setTimeout(function() {
                    resolve(true);
                }, RESP_TIMEOUTS.waitAddRoleResultMs);
                respState.timeouts.push(wt);
            }
            tryClick();
        });
    }

    function respSetAriaBusyOn() {
        var target = document.querySelector(RESP_ATTRS.ariaBusyTarget);
        if (target) {
            respState.prevAriaBusy = target.getAttribute(RESP_ATTRS.ariaBusyAttr);
            target.setAttribute(RESP_ATTRS.ariaBusyAttr, 'true');
            addLogMessage('respSetAriaBusyOn: set aria-busy=true', 'log');
        }
    }

    function respSetAriaBusyOff() {
        var target = document.querySelector(RESP_ATTRS.ariaBusyTarget);
        if (target) {
            if (respState.prevAriaBusy !== null) {
                target.setAttribute(RESP_ATTRS.ariaBusyAttr, respState.prevAriaBusy);
            } else {
                target.removeAttribute(RESP_ATTRS.ariaBusyAttr);
            }
            respState.prevAriaBusy = null;
            addLogMessage('respSetAriaBusyOff: restored aria-busy', 'log');
        }
    }

    function processRolesWorkflow(rolesData) {
        addLogMessage('processRolesWorkflow: starting with ' + rolesData.length + ' roles', 'log');
        respState.isRunning = true;
        respState.stopRequested = false;
        respSetAriaBusyOn();
        updateRespAriaLive(RESP_LABELS.scanningExisting);
        scanExistingCompletedRoles();
        var roleIndex = 0;

        function processNextRole() {
            if (respState.stopRequested) {
                addLogMessage('processRolesWorkflow: stop requested, marking remaining roles', 'log');
                for (var si = roleIndex; si < rolesData.length; si++) {
                    if (rolesData[si].status === RESP_LABELS.statusPending) {
                        rolesData[si].status = RESP_LABELS.statusStopped;
                        respState.counters.pending--;
                        updateRespRoleStatus(rolesData[si].key, RESP_LABELS.statusStopped, '');
                    }
                }
                updateRespSummary();
                respSetAriaBusyOff();
                var t1 = document.getElementById('resp-progress-title');
                if (t1) {
                    t1.textContent = 'Set Responsibilities - Stopped';
                }
                var b1 = document.getElementById('resp-status-badge');
                if (b1) {
                    b1.textContent = 'Stopped';
                    b1.style.color = '#aaa';
                }
                updateRespAriaLive('Process stopped');
                respState.isRunning = false;
                return;
            }
            if (roleIndex >= rolesData.length) {
                addLogMessage('processRolesWorkflow: all roles processed', 'log');
                respSetAriaBusyOff();
                respState.isRunning = false;
                var t2 = document.getElementById('resp-progress-title');
                if (t2) {
                    t2.textContent = 'Set Responsibilities - Complete';
                }
                var b2 = document.getElementById('resp-status-badge');
                if (b2) {
                    b2.textContent = 'Complete';
                    b2.style.background = 'rgba(107, 207, 127, 0.3)';
                    b2.style.color = '#6bcf7f';
                }
                updateRespSummary();
                updateRespAriaLive('Complete. Done: ' + respState.counters.completed + ', Failed: ' + respState.counters.failed);
                return;
            }
            var role = rolesData[roleIndex];
            if (role.status !== RESP_LABELS.statusPending) {
                addLogMessage('processRolesWorkflow: skipping ' + role.displayRole + ' (status=' + role.status + ')', 'log');
                roleIndex++;
                var sk = setTimeout(processNextRole, 50);
                respState.timeouts.push(sk);
                return;
            }
            addLogMessage('processRolesWorkflow: processing role ' + (roleIndex + 1) + '/' + rolesData.length + ': ' + role.displayRole, 'log');
            updateRespAriaLive(RESP_LABELS.selectingRole + ': ' + role.displayRole);
            var currentColumnEl = null;

            findAvailableRoleColumn()
                .then(function(columnEl) {
                currentColumnEl = columnEl;
                if (respState.stopRequested) {
                    throw new Error('Stopped');
                }
                addLogMessage('processRolesWorkflow: got column for ' + role.displayRole, 'log');
                return ensureRoleListOpenForColumn(columnEl);
            })
                .then(function() {
                if (respState.stopRequested) {
                    throw new Error('Stopped');
                }
                addLogMessage('processRolesWorkflow: list open, attempting scroll select for ' + role.displayRole, 'log');
                return attemptSelectByScrollingForRole(currentColumnEl, role.displayRole, role.key);
            })
                .then(function(selected) {
                if (respState.stopRequested) {
                    throw new Error('Stopped');
                }
                if (!selected) {
                    addLogMessage('processRolesWorkflow: role not found in dropdown: ' + role.displayRole, 'warn');
                    role.status = RESP_LABELS.statusFailed;
                    role.reason = 'Role not in dropdown';
                    respState.counters.failed++;
                    respState.counters.pending--;
                    updateRespRoleStatus(role.key, RESP_LABELS.statusFailed, 'Role not in dropdown');
                    updateRespSummary();
                    roleIndex++;
                    var ft = setTimeout(processNextRole, 200);
                    respState.timeouts.push(ft);
                    return;
                }
                addLogMessage('processRolesWorkflow: role selected, opening responsibilities for ' + role.displayRole, 'log');
                updateRespAriaLive(RESP_LABELS.selectingResponsibilities + ': ' + role.displayRole);
                return openResponsibilitiesDropdown(currentColumnEl)
                    .then(function(menu) {
                    if (respState.stopRequested) {
                        throw new Error('Stopped');
                    }
                    addLogMessage('processRolesWorkflow: responsibilities menu open for ' + role.displayRole, 'log');
                    return selectResponsibilitiesByNumbers(role.intersection, menu);
                })
                    .then(function() {
                    if (respState.stopRequested) {
                        throw new Error('Stopped');
                    }
                    addLogMessage('processRolesWorkflow: responsibilities selected, clicking Add Study Role for ' + role.displayRole, 'log');
                    updateRespAriaLive(RESP_LABELS.addingRole + ': ' + role.displayRole);
                    return dismissExistingResponsibilitiesMenu();
                })
                    .then(function() {
                    if (respState.stopRequested) {
                        throw new Error('Stopped');
                    }
                    return clickAddStudyRole();
                })
                    .then(function(addOk) {
                    if (addOk) {
                        addLogMessage('processRolesWorkflow: ' + role.displayRole + ' completed', 'log');
                        role.status = RESP_LABELS.statusCompleted;
                        respState.counters.completed++;
                        respState.counters.pending--;
                        updateRespRoleStatus(role.key, RESP_LABELS.statusCompleted, '');
                    } else {
                        addLogMessage('processRolesWorkflow: ' + role.displayRole + ' add failed', 'warn');
                        role.status = RESP_LABELS.statusFailed;
                        role.reason = 'Add Study Role failed';
                        respState.counters.failed++;
                        respState.counters.pending--;
                        updateRespRoleStatus(role.key, RESP_LABELS.statusFailed, 'Add Study Role failed');
                    }
                    updateRespSummary();
                    roleIndex++;
                    var nt = setTimeout(processNextRole, 200);
                    respState.timeouts.push(nt);
                });
            })
                .catch(function(err) {
                addLogMessage('processRolesWorkflow: error for ' + role.displayRole + ': ' + err.message, 'error');
                if (respState.stopRequested) {
                    role.status = RESP_LABELS.statusStopped;
                    updateRespRoleStatus(role.key, RESP_LABELS.statusStopped, '');
                    respState.counters.pending--;
                } else {
                    role.status = RESP_LABELS.statusFailed;
                    role.reason = err.message;
                    respState.counters.failed++;
                    respState.counters.pending--;
                    updateRespRoleStatus(role.key, RESP_LABELS.statusFailed, err.message);
                }
                updateRespSummary();
                roleIndex++;
                var et = setTimeout(processNextRole, 200);
                respState.timeouts.push(et);
            });
        }
        processNextRole();
    }

    function resetCleanState() {
        addLogMessage('resetCleanState: resetting state', 'log');
        cleanState.isRunning = false;
        cleanState.stopRequested = false;
        cleanState.observers = [];
        cleanState.timeouts = [];
        cleanState.intervals = [];
        cleanState.eventListeners = [];
        cleanState.idleCallbackIds = [];
        cleanState.prevAriaBusy = null;
        cleanState.parsedItems = null;
        cleanState.outputModel = null;
    }

    function cleanSetAriaBusyOn() {
        addLogMessage('cleanSetAriaBusyOn: setting aria-busy', 'log');
        var target = document.querySelector(CLEAN_ATTRS.ariaBusyTarget);
        if (target) {
            cleanState.prevAriaBusy = target.getAttribute(CLEAN_ATTRS.ariaBusyAttr);
            target.setAttribute(CLEAN_ATTRS.ariaBusyAttr, 'true');
            addLogMessage('cleanSetAriaBusyOn: prev=' + cleanState.prevAriaBusy, 'log');
        }
    }

    function cleanSetAriaBusyOff() {
        addLogMessage('cleanSetAriaBusyOff: restoring aria-busy', 'log');
        var target = document.querySelector(CLEAN_ATTRS.ariaBusyTarget);
        if (target) {
            if (cleanState.prevAriaBusy !== null) {
                target.setAttribute(CLEAN_ATTRS.ariaBusyAttr, cleanState.prevAriaBusy);
            } else {
                target.removeAttribute(CLEAN_ATTRS.ariaBusyAttr);
            }
            cleanState.prevAriaBusy = null;
        }
    }

    function updateCleanAriaLive(message) {
        addLogMessage('updateCleanAriaLive: ' + message, 'log');
        var lr = document.getElementById('clean-aria-live');
        if (lr) {
            lr.textContent = message;
        }
    }

    function normalizeItemText(s) {
        addLogMessage('normalizeItemText: input length=' + s.length, 'log');
        var result = s;
        result = result.replace(CLEAN_REGEX.smartQuotes, '"');
        result = result.replace(CLEAN_REGEX.strayQuotes, '');
        result = result.replace(CLEAN_REGEX.whitespace, ' ');
        result = result.trim();
        addLogMessage('normalizeItemText: output length=' + result.length, 'log');
        return result;
    }

    function replaceSpecialCharactersToComma(s) {
        addLogMessage('replaceSpecialCharactersToComma: input length=' + s.length, 'log');
        var count = 0;
        var result = s.replace(CLEAN_REGEX.specialChars, function() {
            count++;
            return ',';
        });
        result = result.replace(CLEAN_REGEX.duplicateCommas, ',');
        result = result.replace(/,\s*/g, ', ');
        result = result.replace(/\s+/g, ' ');
        result = result.trim();
        addLogMessage('replaceSpecialCharactersToComma: replacements=' + count, 'log');
        return result;
    }

    function replaceStandaloneAndWithAmpersand(s) {
        addLogMessage('replaceStandaloneAndWithAmpersand: input length=' + s.length, 'log');
        var count = 0;
        var result = s.replace(CLEAN_REGEX.standaloneAnd, function() {
            count++;
            return '&';
        });
        addLogMessage('replaceStandaloneAndWithAmpersand: replacements=' + count, 'log');
        return result;
    }

    function splitRawIntoItems(rawText) {
        addLogMessage('splitRawIntoItems: starting split', 'log');
        var items = [];
        var normalized = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        normalized = normalized.replace(CLEAN_REGEX.smartQuotes, '"');
        normalized = normalized.replace(CLEAN_REGEX.strayQuotes, '');
        var pattern = /(?:^|[\s"'])(\d{1,3})([.)])\s+/g;
        var boundaries = [];
        var match;
        while ((match = pattern.exec(normalized)) !== null) {
            var tokenStart = match.index;
            if (match[1] !== undefined && match.index < normalized.length) {
                var prefix = match[0];
                var leadingWhitespace = prefix.length - prefix.trimStart().length;
                tokenStart = match.index + leadingWhitespace;
                if (/["'\s]/.test(normalized[match.index]) && match.index !== 0) {
                    tokenStart = match.index + 1;
                } else if (match.index === 0) {
                    tokenStart = 0;
                }
            }
            boundaries.push({
                index: tokenStart,
                number: match[1],
                separator: match[2],
                fullMatchEnd: match.index + match[0].length
            });
        }
        addLogMessage('splitRawIntoItems: found ' + boundaries.length + ' boundaries', 'log');
        for (var i = 0; i < boundaries.length; i++) {
            var start = boundaries[i].fullMatchEnd;
            var end;
            if (i + 1 < boundaries.length) {
                end = boundaries[i + 1].index;
            } else {
                end = normalized.length;
            }
            var text = normalized.substring(start, end).trim();
            text = text.replace(/\n+/g, ' ');
            text = text.replace(CLEAN_REGEX.whitespace, ' ');
            items.push({
                number: boundaries[i].number,
                separator: boundaries[i].separator,
                rawText: text
            });
        }
        addLogMessage('splitRawIntoItems: extracted ' + items.length + ' items', 'log');
        return items;
    }

    function parseAndCleanResponsibilities(rawText) {
        addLogMessage('parseAndCleanResponsibilities: start', 'log');
        try {
            cleanSetAriaBusyOn();
            updateCleanAriaLive(CLEAN_LABELS.parsing);
            var rawItems = splitRawIntoItems(rawText);
            var results = [];
            var tooLongCount = 0;
            var tooShortCount = 0;
            for (var i = 0; i < rawItems.length; i++) {
                var item = rawItems[i];
                var cleaned = normalizeItemText(item.rawText);
                cleaned = replaceSpecialCharactersToComma(cleaned);
                cleaned = replaceStandaloneAndWithAmpersand(cleaned);
                cleaned = cleaned.trim();
                cleaned = cleaned.replace(CLEAN_REGEX.whitespace, ' ');
                var numberToken = item.number + item.separator;
                var fullText = numberToken + ' ' + cleaned;
                var textLength = cleaned.length;
                var isTooLong = textLength > CLEAN_LIMITS.maxChars;
                var isTooShort = textLength < CLEAN_LIMITS.maxChars;
                if (isTooLong) {
                    tooLongCount++;
                }
                if (isTooShort) {
                    tooShortCount++;
                }
                results.push({
                    numberToken: numberToken,
                    text: cleaned,
                    fullTextWithNumber: fullText,
                    tooLong: isTooLong,
                    tooShort: isTooShort,
                    length: textLength
                });
            }
            addLogMessage('parseAndCleanResponsibilities: parsed ' + results.length + ' items, tooLong=' + tooLongCount + ', tooShort=' + tooShortCount, 'log');
            updateCleanAriaLive(CLEAN_LABELS.done);
            cleanSetAriaBusyOff();
            return {
                items: results,
                summary: {
                    total: results.length,
                    tooLong: tooLongCount,
                    tooShort: tooShortCount
                }
            };
        } catch (e) {
            addLogMessage('parseAndCleanResponsibilities: error: ' + e, 'error');
            cleanSetAriaBusyOff();
            return {
                items: [],
                summary: { total: 0, tooLong: 0, tooShort: 0 }
            };
        }
    }

    function buildOutputModel(items) {
        addLogMessage('buildOutputModel: building model with ' + items.length + ' items', 'log');
        var headerLine = CLEAN_LABELS.responsibilitiesHeader;
        var cleanLines = [];
        var flags = [];
        var tooLongCount = 0;
        var tooShortCount = 0;
        for (var i = 0; i < items.length; i++) {
            cleanLines.push(items[i].text);
            flags.push({
                tooLong: items[i].tooLong,
                tooShort: items[i].tooShort,
                length: items[i].length
            });
            if (items[i].tooLong) {
                tooLongCount++;
            }
            if (items[i].tooShort) {
                tooShortCount++;
            }
        }
        addLogMessage('buildOutputModel: tooLong=' + tooLongCount + ', tooShort=' + tooShortCount, 'log');
        return {
            headerLine: headerLine,
            cleanLines: cleanLines,
            flags: flags,
            copyText: headerLine + '\n' + cleanLines.join('\n'),
            tooLongCount: tooLongCount,
            tooShortCount: tooShortCount
        };
    }

    function exportResponsibilitiesToXlsxOrCsv(model) {
        addLogMessage('exportResponsibilitiesToXlsxOrCsv: starting export', 'log');
        try {
            if (typeof window.XLSX !== 'undefined' && window.XLSX) {
                addLogMessage('exportResponsibilitiesToXlsxOrCsv: using XLSX library', 'log');
                var wsData = [[model.headerLine]];
                for (var i = 0; i < model.cleanLines.length; i++) {
                    wsData.push([model.cleanLines[i]]);
                }
                var wb = window.XLSX.utils.book_new();
                var ws = window.XLSX.utils.aoa_to_sheet(wsData);
                window.XLSX.utils.book_append_sheet(wb, ws, 'Responsibilities');
                var timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
                var filename = 'Responsibilities_' + timestamp + '.xlsx';
                window.XLSX.writeFile(wb, filename);
                addLogMessage('exportResponsibilitiesToXlsxOrCsv: xlsx exported as ' + filename, 'log');
                updateCleanAriaLive(CLEAN_LABELS.exportSuccess);
                return true;
            } else {
                addLogMessage('exportResponsibilitiesToXlsxOrCsv: XLSX not available, falling back to CSV', 'log');
                var csvRows = [];
                csvRows.push('"' + model.headerLine.replace(/"/g, '""') + '"');
                for (var j = 0; j < model.cleanLines.length; j++) {
                    csvRows.push('"' + model.cleanLines[j].replace(/"/g, '""') + '"');
                }
                var csvContent = csvRows.join('\r\n');
                var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                var url = URL.createObjectURL(blob);
                var timestamp2 = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
                var filename2 = 'Responsibilities_' + timestamp2 + '.csv';
                var link = document.createElement('a');
                link.href = url;
                link.download = filename2;
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                addLogMessage('exportResponsibilitiesToXlsxOrCsv: csv exported as ' + filename2, 'log');
                updateCleanAriaLive(CLEAN_LABELS.exportSuccess);
                return true;
            }
        } catch (e) {
            addLogMessage('exportResponsibilitiesToXlsxOrCsv: export failed: ' + e, 'error');
            updateCleanAriaLive(CLEAN_LABELS.exportFailed);
            return false;
        }
    }

    function showCleanInputPanel() {
        addLogMessage('showCleanInputPanel: creating input panel', 'log');
        var modal = document.createElement('div');
        modal.id = 'clean-input-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 550px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'clean-input-title');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;';
        var title = document.createElement('h3');
        title.id = 'clean-input-title';
        title.textContent = CLEAN_LABELS.inputTitle;
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600; letter-spacing: 0.2px;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close input panel');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = function() {
            addLogMessage('showCleanInputPanel: closed by user', 'log');
            stopCleanResponsibility();
        };
        header.appendChild(title);
        header.appendChild(closeButton);
        var description = document.createElement('p');
        description.textContent = 'Paste the raw responsibility list below. Items should start with a number followed by . or ) and text.';
        description.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0 0 12px 0; font-size: 14px; line-height: 1.4;';
        var textareaLabel = document.createElement('label');
        textareaLabel.setAttribute('for', 'clean-input-textarea');
        textareaLabel.textContent = 'Responsibilities text';
        textareaLabel.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        var textarea = document.createElement('textarea');
        textarea.id = 'clean-input-textarea';
        textarea.placeholder = '1. First responsibility\n2. Second responsibility\n3. Third responsibility';
        textarea.style.cssText = 'width: 100%; height: 200px; padding: 12px 14px; border: 2px solid rgba(255, 255, 255, 0.35); border-radius: 10px; background: rgba(255, 255, 255, 0.95); color: #1e293b; font-size: 14px; font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; resize: vertical; outline: none; transition: all 0.25s ease; box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset; box-sizing: border-box;';
        textarea.onfocus = function() {
            textarea.style.borderColor = '#8ea0ff';
            textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)';
        };
        textarea.onblur = function() {
            textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)';
            textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset';
        };
        var confirmButton = document.createElement('button');
        confirmButton.textContent = CLEAN_LABELS.confirm;
        confirmButton.disabled = true;
        confirmButton.setAttribute('aria-label', 'Confirm and parse responsibilities');
        confirmButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; opacity: 0.5;';
        var updateConfirmState = function() {
            var hasInput = textarea.value.trim().length > 0;
            confirmButton.disabled = !hasInput;
            if (hasInput) {
                confirmButton.style.opacity = '1';
                confirmButton.style.cursor = 'pointer';
            } else {
                confirmButton.style.opacity = '0.5';
                confirmButton.style.cursor = 'not-allowed';
            }
        };
        textarea.addEventListener('input', updateConfirmState);
        cleanState.eventListeners.push({ element: textarea, type: 'input', handler: updateConfirmState });
        confirmButton.onmouseover = function() {
            if (!confirmButton.disabled) {
                confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
            }
        };
        confirmButton.onmouseout = function() {
            confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        confirmButton.onclick = function() {
            if (confirmButton.disabled) {
                return;
            }
            addLogMessage('showCleanInputPanel: Confirm clicked', 'log');
            var rawText = textarea.value;
            if (!rawText || !rawText.trim()) {
                addLogMessage('showCleanInputPanel: empty input after confirm', 'warn');
                return;
            }
            var parseResult = parseAndCleanResponsibilities(rawText);
            if (parseResult.items.length === 0) {
                addLogMessage('showCleanInputPanel: no items parsed, showing notice', 'warn');
                var notice = document.createElement('div');
                notice.textContent = 'No numbered items found. Ensure items start with a number followed by . or ) and text.';
                notice.style.cssText = 'color: #ffd93d; font-size: 13px; margin-top: 8px; padding: 8px; background: rgba(255, 217, 61, 0.15); border-radius: 6px;';
                notice.setAttribute('role', 'alert');
                var existingNotice = container.querySelector('[role="alert"]');
                if (existingNotice) {
                    container.removeChild(existingNotice);
                }
                container.appendChild(notice);
                return;
            }
            cleanState.parsedItems = parseResult;
            var outputModel = buildOutputModel(parseResult.items);
            cleanState.outputModel = outputModel;
            if (modal.parentNode) {
                modal.parentNode.removeChild(modal);
            }
            showCleanResultsPanel(outputModel);
        };
        var clearButton = document.createElement('button');
        clearButton.textContent = CLEAN_LABELS.clear;
        clearButton.setAttribute('aria-label', 'Clear all input');
        clearButton.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease; backdrop-filter: blur(2px);';
        clearButton.onmouseover = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        clearButton.onmouseout = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        clearButton.onclick = function() {
            addLogMessage('showCleanInputPanel: Clear All clicked', 'log');
            textarea.value = '';
            cleanState.parsedItems = null;
            cleanState.outputModel = null;
            updateConfirmState();
            textarea.focus();
        };
        var buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = 'display: flex; gap: 12px; margin-top: 20px; justify-content: flex-end;';
        buttonContainer.appendChild(clearButton);
        buttonContainer.appendChild(confirmButton);
        var ariaLiveRegion = document.createElement('div');
        ariaLiveRegion.id = 'clean-aria-live';
        ariaLiveRegion.setAttribute('aria-live', 'polite');
        ariaLiveRegion.setAttribute('aria-atomic', 'true');
        ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        container.appendChild(header);
        container.appendChild(description);
        container.appendChild(textareaLabel);
        container.appendChild(textarea);
        container.appendChild(buttonContainer);
        container.appendChild(ariaLiveRegion);
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
        addLogMessage('showCleanInputPanel: panel displayed', 'log');
        var escHandler = function(e) {
            if (e.key === 'Escape') {
                addLogMessage('showCleanInputPanel: Escape pressed', 'log');
                stopCleanResponsibility();
            }
        };
        document.addEventListener('keydown', escHandler);
        cleanState.eventListeners.push({ element: document, type: 'keydown', handler: escHandler });
    }

    function showCleanResultsPanel(model) {
        addLogMessage('showCleanResultsPanel: creating results panel', 'log');
        var modal = document.createElement('div');
        modal.id = 'clean-results-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 650px; max-width: 90%; max-height: 85vh; display: flex; flex-direction: column; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'clean-results-title');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0;';
        var title = document.createElement('h3');
        title.id = 'clean-results-title';
        title.textContent = CLEAN_LABELS.resultsTitle;
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600; letter-spacing: 0.2px;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close results panel');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = function() {
            addLogMessage('showCleanResultsPanel: closed by user', 'log');
            stopCleanResponsibility();
        };
        header.appendChild(title);
        header.appendChild(closeButton);
        var summaryBar = document.createElement('div');
        summaryBar.style.cssText = 'display: flex; gap: 12px; margin-bottom: 12px; flex-shrink: 0; flex-wrap: wrap;';
        var totalBadge = document.createElement('span');
        totalBadge.textContent = 'Total: ' + model.cleanLines.length;
        totalBadge.style.cssText = 'background: rgba(255, 255, 255, 0.15); color: white; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
        summaryBar.appendChild(totalBadge);
        if (model.tooLongCount > 0) {
            var longBadge = document.createElement('span');
            longBadge.textContent = 'Over ' + CLEAN_LIMITS.maxChars + ' chars: ' + model.tooLongCount;
            longBadge.style.cssText = 'background: rgba(255, 107, 107, 0.25); color: #ff6b6b; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
            summaryBar.appendChild(longBadge);
        }
        if (model.tooShortCount > 0) {
            var shortBadge = document.createElement('span');
            shortBadge.textContent = 'Under ' + CLEAN_LIMITS.maxChars + ' chars: ' + model.tooShortCount;
            shortBadge.style.cssText = 'background: rgba(107, 207, 127, 0.25); color: #6bcf7f; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
            summaryBar.appendChild(shortBadge);
        }
        var outputScrollStyle = document.createElement('style');
        outputScrollStyle.textContent = '.clean-results-output::-webkit-scrollbar { width: 6px; } .clean-results-output::-webkit-scrollbar-track { background: rgba(0, 0, 0, 0.1); border-radius: 3px; } .clean-results-output::-webkit-scrollbar-thumb { background: rgba(255, 255, 255, 0.3); border-radius: 3px; } .clean-results-output::-webkit-scrollbar-thumb:hover { background: rgba(255, 255, 255, 0.5); } .clean-results-warning-icon { user-select: none; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; pointer-events: auto; }';
        document.head.appendChild(outputScrollStyle);
        var outputArea = document.createElement('div');
        outputArea.style.cssText = 'background: rgba(0, 0, 0, 0.2); border-radius: 8px; padding: 12px; overflow-y: auto; flex: 1; min-height: 100px; max-height: 400px;';
        outputArea.className = 'clean-results-output';
        var copyBlock = document.createElement('div');
        copyBlock.id = 'clean-copy-block';
        var headerLine = document.createElement('div');
        headerLine.textContent = model.headerLine;
        headerLine.style.cssText = 'color: white; font-size: 15px; font-weight: 700; margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1px solid rgba(255, 255, 255, 0.2);';
        copyBlock.appendChild(headerLine);
        for (var i = 0; i < model.cleanLines.length; i++) {
            var row = document.createElement('div');
            row.style.cssText = 'display: flex; align-items: flex-start; gap: 6px; padding: 4px 0; color: white; font-size: 14px; line-height: 1.5;';
            var textSpan = document.createElement('span');
            textSpan.textContent = model.cleanLines[i];
            textSpan.style.cssText = 'flex: 1; word-break: break-word;';
            row.appendChild(textSpan);
            if (model.flags[i].tooLong) {
                var warnIcon = document.createElement('span');
                warnIcon.className = 'clean-results-warning-icon';
                warnIcon.textContent = '\u26A0';
                warnIcon.setAttribute('aria-hidden', 'true');
                warnIcon.setAttribute('title', 'Exceeds ' + CLEAN_LIMITS.maxChars + ' character limit. Length: ' + model.flags[i].length);
                warnIcon.style.cssText = 'color: #ffd93d; font-size: 14px; flex-shrink: 0; cursor: help; user-select: none; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none;';
                row.appendChild(warnIcon);
            }
            outputArea.appendChild(row);
        }
        var hasXlsx = typeof window.XLSX !== 'undefined' && window.XLSX;
        var downloadButton = document.createElement('button');
        downloadButton.textContent = hasXlsx ? CLEAN_LABELS.downloadXlsx : CLEAN_LABELS.downloadCsv;
        downloadButton.setAttribute('aria-label', hasXlsx ? 'Download as Excel file' : 'Download as CSV file');
        downloadButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; flex: 1;';
        downloadButton.onmouseover = function() {
            downloadButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
        };
        downloadButton.onmouseout = function() {
            downloadButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        downloadButton.onclick = function() {
            addLogMessage('showCleanResultsPanel: Download clicked', 'log');
            var success = exportResponsibilitiesToXlsxOrCsv(model);
            if (success) {
                addLogMessage('showCleanResultsPanel: export succeeded', 'log');
            } else {
                addLogMessage('showCleanResultsPanel: export failed', 'error');
                var exportNotice = document.createElement('div');
                exportNotice.textContent = CLEAN_LABELS.exportFailed;
                exportNotice.style.cssText = 'color: #ff6b6b; font-size: 13px; margin-top: 8px; padding: 8px; background: rgba(255, 107, 107, 0.15); border-radius: 6px;';
                exportNotice.setAttribute('role', 'alert');
                var existingNotice = container.querySelector('[role="alert"]');
                if (existingNotice) {
                    container.removeChild(existingNotice);
                }
                container.appendChild(exportNotice);
            }
        };
        var closeBtn = document.createElement('button');
        closeBtn.textContent = CLEAN_LABELS.close;
        closeBtn.setAttribute('aria-label', 'Close results');
        closeBtn.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease; backdrop-filter: blur(2px); flex: 1;';
        closeBtn.onmouseover = function() {
            closeBtn.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        closeBtn.onmouseout = function() {
            closeBtn.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        closeBtn.onclick = function() {
            addLogMessage('showCleanResultsPanel: Close clicked', 'log');
            stopCleanResponsibility();
        };
        var bottomButtons = document.createElement('div');
        bottomButtons.style.cssText = 'display: flex; gap: 12px; margin-top: 16px; flex-shrink: 0;';
        bottomButtons.appendChild(downloadButton);
        bottomButtons.appendChild(closeBtn);
        var ariaLiveRegion = document.createElement('div');
        ariaLiveRegion.id = 'clean-aria-live';
        ariaLiveRegion.setAttribute('aria-live', 'polite');
        ariaLiveRegion.setAttribute('aria-atomic', 'true');
        ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        container.appendChild(header);
        container.appendChild(summaryBar);
        container.appendChild(outputArea);
        container.appendChild(bottomButtons);
        container.appendChild(ariaLiveRegion);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);
        downloadButton.focus();
        addLogMessage('showCleanResultsPanel: panel displayed with ' + model.cleanLines.length + ' items', 'log');
        var escHandler = function(e) {
            if (e.key === 'Escape') {
                addLogMessage('showCleanResultsPanel: Escape pressed', 'log');
                stopCleanResponsibility();
            }
        };
        document.addEventListener('keydown', escHandler);
        cleanState.eventListeners.push({ element: document, type: 'keydown', handler: escHandler });
    }

    function cleanResponsibilityInit() {
        addLogMessage('cleanResponsibilityInit: starting feature', 'log');
        cleanState.focusReturnElement = document.getElementById('clean-resp-btn');
        resetCleanState();
        cleanState.isRunning = true;
        showCleanInputPanel();
    }

    function stopCleanResponsibility() {
        addLogMessage('stopCleanResponsibility: stopping', 'log');
        cleanState.stopRequested = true;
        cleanState.isRunning = false;
        for (var i = 0; i < cleanState.idleCallbackIds.length; i++) {
            try {
                if (typeof cancelIdleCallback === 'function') {
                    cancelIdleCallback(cleanState.idleCallbackIds[i]);
                }
            } catch (e) {
                addLogMessage('stopCleanResponsibility: error canceling idle callback: ' + e, 'error');
            }
        }
        cleanState.idleCallbackIds = [];
        for (var i2 = 0; i2 < cleanState.observers.length; i2++) {
            try {
                cleanState.observers[i2].disconnect();
            } catch (e2) {
                addLogMessage('stopCleanResponsibility: error disconnecting observer: ' + e2, 'error');
            }
        }
        cleanState.observers = [];
        for (var i3 = 0; i3 < cleanState.timeouts.length; i3++) {
            try {
                clearTimeout(cleanState.timeouts[i3]);
            } catch (e3) {
                addLogMessage('stopCleanResponsibility: error clearing timeout: ' + e3, 'error');
            }
        }
        cleanState.timeouts = [];
        for (var i4 = 0; i4 < cleanState.intervals.length; i4++) {
            try {
                clearInterval(cleanState.intervals[i4]);
            } catch (e4) {
                addLogMessage('stopCleanResponsibility: error clearing interval: ' + e4, 'error');
            }
        }
        cleanState.intervals = [];
        for (var i5 = 0; i5 < cleanState.eventListeners.length; i5++) {
            try {
                var l = cleanState.eventListeners[i5];
                l.element.removeEventListener(l.type, l.handler);
            } catch (e5) {
                addLogMessage('stopCleanResponsibility: error removing listener: ' + e5, 'error');
            }
        }
        cleanState.eventListeners = [];
        cleanSetAriaBusyOff();
        var im = document.getElementById('clean-input-modal');
        if (im && im.parentNode) {
            im.parentNode.removeChild(im);
        }
        var rm = document.getElementById('clean-results-modal');
        if (rm && rm.parentNode) {
            rm.parentNode.removeChild(rm);
        }
        if (cleanState.focusReturnElement) {
            cleanState.focusReturnElement.focus();
        }
        updateCleanAriaLive('Clean Responsibility stopped');
        resetCleanState();
        addLogMessage('stopCleanResponsibility: cleanup complete', 'log');
    }

    function resetDoAState() {
        addLogMessage('resetDoAState: resetting state', 'log');
        doaState.isRunning = false;
        doaState.observers = [];
        doaState.timeouts = [];
        doaState.intervals = [];
        doaState.eventListeners = [];
        doaState.idleCallbackIds = [];
        doaState.parsedCandidates = [];
        doaState.scannedNames = [];
        doaState.seenNormalizedNames = new Set();
        doaState.prevAriaBusy = null;
        doaState.scrollContainer = null;
        doaState.prevScrollTop = 0;
        doaState.userScrollHandler = null;
        doaState.userScrollPaused = false;
        doaState.idleCallbackId = null;
        doaState.leftPanelRowIndex = 0;
        doaState.lastAutoScrollTime = 0;
        doaState.addQueue = [];
        doaState.addQueueIndex = 0;
        doaState.existingPairs = new Set();
        doaState.counters = { total: 0, added: 0, duplicates: 0, failures: 0, pending: 0 };
        doaState.listScrollTop = 0;
        doaState.roleListScrollTop = 0;
        doaState.isAddingEntries = false;
    }

    function doaSetAriaBusyOn() {
        addLogMessage('doaSetAriaBusyOn: setting aria-busy', 'log');
        var target = document.querySelector(ELOG_ATTRS.ariaBusyTarget);
        if (target) {
            doaState.prevAriaBusy = target.getAttribute(ELOG_ATTRS.ariaBusyAttr);
            target.setAttribute(ELOG_ATTRS.ariaBusyAttr, 'true');
            addLogMessage('doaSetAriaBusyOn: prev=' + doaState.prevAriaBusy, 'log');
        }
    }

    function doaSetAriaBusyOff() {
        addLogMessage('doaSetAriaBusyOff: restoring aria-busy', 'log');
        var target = document.querySelector(ELOG_ATTRS.ariaBusyTarget);
        if (target) {
            if (doaState.prevAriaBusy !== null) {
                target.setAttribute(ELOG_ATTRS.ariaBusyAttr, doaState.prevAriaBusy);
            } else {
                target.removeAttribute(ELOG_ATTRS.ariaBusyAttr);
            }
            doaState.prevAriaBusy = null;
        }
    }

    function doaWaitForElement(selector, timeout) {
        addLogMessage('doaWaitForElement: waiting for ' + selector, 'log');
        return new Promise(function(resolve, reject) {
            var element = document.querySelector(selector);
            if (element) {
                addLogMessage('doaWaitForElement: found immediately', 'log');
                resolve(element);
                return;
            }
            var observer = new MutationObserver(function(mutations, obs) {
                var el = document.querySelector(selector);
                if (el) {
                    obs.disconnect();
                    var idx = doaState.observers.indexOf(obs);
                    if (idx > -1) {
                        doaState.observers.splice(idx, 1);
                    }
                    resolve(el);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            doaState.observers.push(observer);
            var timeoutId = setTimeout(function() {
                observer.disconnect();
                var idx = doaState.observers.indexOf(observer);
                if (idx > -1) {
                    doaState.observers.splice(idx, 1);
                }
                addLogMessage('doaWaitForElement: timeout for ' + selector, 'warn');
                reject(new Error('Timeout waiting for ' + selector));
            }, timeout);
            doaState.timeouts.push(timeoutId);
        });
    }

    function doaDelay(ms) {
        return new Promise(function(resolve) {
            var tid = setTimeout(resolve, ms);
            doaState.timeouts.push(tid);
        });
    }

    function updateDoAAriaLive(message) {
        addLogMessage('updateDoAAriaLive: ' + message, 'log');
        var liveRegion = document.getElementById('doa-aria-live');
        if (liveRegion) {
            liveRegion.textContent = message;
        }
    }

    function expandResponsibilityTokensToSet(tokens) {
        addLogMessage('expandResponsibilityTokensToSet: tokens count=' + tokens.length, 'log');
        var result = new Set();
        var i = 0;
        while (i < tokens.length) {
            var current = tokens[i];
            if (i + 2 < tokens.length && /^(to)$/i.test(tokens[i + 1])) {
                var start = parseInt(current, 10);
                var end = parseInt(tokens[i + 2], 10);
                if (!isNaN(start) && !isNaN(end)) {
                    var lo = Math.min(start, end);
                    var hi = Math.max(start, end);
                    for (var n = lo; n <= hi; n++) {
                        result.add(n);
                    }
                    i += 3;
                    continue;
                }
            }
            if (i + 1 < tokens.length) {
                var dashMatch = (current + tokens[i + 1]).match(/^(\d+)[\-\u2013\u2014](\d+)$/);
                if (dashMatch) {
                    var dStart = parseInt(dashMatch[1], 10);
                    var dEnd = parseInt(dashMatch[2], 10);
                    if (!isNaN(dStart) && !isNaN(dEnd)) {
                        var dLo = Math.min(dStart, dEnd);
                        var dHi = Math.max(dStart, dEnd);
                        for (var dn = dLo; dn <= dHi; dn++) {
                            result.add(dn);
                        }
                        i += 2;
                        continue;
                    }
                }
            }
            var singleDash = current.match(/^(\d+)[\-\u2013\u2014](\d+)$/);
            if (singleDash) {
                var sdStart = parseInt(singleDash[1], 10);
                var sdEnd = parseInt(singleDash[2], 10);
                if (!isNaN(sdStart) && !isNaN(sdEnd)) {
                    var sdLo = Math.min(sdStart, sdEnd);
                    var sdHi = Math.max(sdStart, sdEnd);
                    for (var sdn = sdLo; sdn <= sdHi; sdn++) {
                        result.add(sdn);
                    }
                    i++;
                    continue;
                }
            }
            var num = parseInt(current, 10);
            if (!isNaN(num)) {
                result.add(num);
            }
            i++;
        }
        addLogMessage('expandResponsibilityTokensToSet: result size=' + result.size, 'log');
        return result;
    }

    function parseDoAEntriesInput(rawText) {
        addLogMessage('parseDoAEntriesInput: starting parse', 'log');
        if (!rawText || !rawText.trim()) {
            addLogMessage('parseDoAEntriesInput: empty input', 'warn');
            return [];
        }
        var text = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        text = text.replace(/[\u201c\u201d\u201e\u201f\u2018\u2019]+/g, '"');
        text = text.replace(/-\s*\n\s*/g, '-');
        var rawLines = text.split('\n');
        var mergedLines = [];
        var pendingQuote = false;
        var buffer = '';
        for (var li = 0; li < rawLines.length; li++) {
            var line = rawLines[li];
            if (pendingQuote) {
                buffer = buffer + ' ' + line;
                var quoteCount = 0;
                for (var ci = 0; ci < buffer.length; ci++) {
                    if (buffer[ci] === '"') {
                        quoteCount++;
                    }
                }
                if (quoteCount % 2 === 0) {
                    pendingQuote = false;
                    mergedLines.push(buffer);
                    buffer = '';
                }
                continue;
            }
            var qc = 0;
            for (var ci2 = 0; ci2 < line.length; ci2++) {
                if (line[ci2] === '"') {
                    qc++;
                }
            }
            if (qc % 2 !== 0) {
                pendingQuote = true;
                buffer = line;
                continue;
            }
            mergedLines.push(line);
        }
        if (buffer) {
            mergedLines.push(buffer);
        }
        addLogMessage('parseDoAEntriesInput: merged into ' + mergedLines.length + ' lines', 'log');
        var results = [];
        var seenPairKeys = new Set();
        for (var mi = 0; mi < mergedLines.length; mi++) {
            var mline = mergedLines[mi].trim();
            if (!mline) {
                continue;
            }
            mline = mline.replace(/"+/g, '');
            var parts = mline.split(/\t+/);
            if (parts.length < 2) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' too few columns, skip', 'log');
                continue;
            }
            var emailIdx = -1;
            for (var pi = 0; pi < parts.length; pi++) {
                if (parts[pi].indexOf('@') !== -1) {
                    emailIdx = pi;
                    break;
                }
            }
            if (emailIdx === -1 || emailIdx + 1 >= parts.length) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' no email column or no role after email, skip', 'log');
                continue;
            }
            var namePart = parts[0].trim();
            if (!namePart) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' empty name, skip', 'log');
                continue;
            }
            var rolePart = parts[emailIdx + 1].trim();
            if (!rolePart) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' empty role, skip', 'log');
                continue;
            }
            var numberPart = parts.slice(emailIdx + 2).join(' ');
            var numTokens = numberPart.replace(/,/g, ' ').split(/\s+/).filter(function(t) {
                return t.length > 0;
            });
            var numbersSet = expandResponsibilityTokensToSet(numTokens);
            if (numbersSet.size === 0) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' no numbers for ' + namePart + ', skip', 'log');
                continue;
            }
            var displayName = namePart.replace(/\s+/g, ' ').trim();
            var nameTokens = displayName.split(' ');
            var normalizedDisplay = [];
            for (var nt = 0; nt < nameTokens.length; nt++) {
                var token = nameTokens[nt];
                if (token) {
                    normalizedDisplay.push(token.charAt(0).toUpperCase() + token.slice(1).toLowerCase());
                }
            }
            displayName = normalizedDisplay.join(' ');
            var normalized = elogNormalizeName(displayName);
            var pairKey = normalizeFirstLastPair(displayName);
            var roleNorm = normalizeRoleName(rolePart);
            if (!roleNorm.key) {
                addLogMessage('parseDoAEntriesInput: line ' + mi + ' role normalize failed for ' + rolePart + ', skip', 'log');
                continue;
            }
            var sortedNums = Array.from(numbersSet).sort(function(a, b) {
                return a - b;
            });
            addLogMessage('parseDoAEntriesInput: line ' + mi + ' name=' + displayName + ' role=' + roleNorm.display + ' nums=[' + sortedNums.join(',') + ']', 'log');
            results.push({
                display: displayName,
                normalized: normalized,
                pairKey: pairKey,
                roleDisplay: roleNorm.display,
                roleKey: roleNorm.key,
                numbersSet: numbersSet
            });
        }
        addLogMessage('parseDoAEntriesInput: parsed ' + results.length + ' candidates', 'log');
        return results;
    }

    function buildDoAQueueSorted(parsedCandidates) {
        addLogMessage('buildDoAQueueSorted: building from ' + parsedCandidates.length + ' candidates', 'log');
        var seenPairKeys = new Set();
        var unique = [];
        var duplicateIndices = [];
        for (var i = 0; i < parsedCandidates.length; i++) {
            var candidate = parsedCandidates[i];
            if (seenPairKeys.has(candidate.pairKey)) {
                addLogMessage('buildDoAQueueSorted: input duplicate pairKey=' + candidate.pairKey + ' display=' + candidate.display, 'log');
                duplicateIndices.push(i);
                continue;
            }
            seenPairKeys.add(candidate.pairKey);
            unique.push({
                display: candidate.display,
                normalized: candidate.normalized,
                pairKey: candidate.pairKey,
                roleDisplay: candidate.roleDisplay,
                roleKey: candidate.roleKey,
                numbersSet: candidate.numbersSet,
                status: DOA_LABELS.statusPending
            });
        }
        unique.sort(function(a, b) {
            var aPair = getFirstAndLast(a.display);
            var bPair = getFirstAndLast(b.display);
            var firstCmp = aPair[0].localeCompare(bPair[0], undefined, { sensitivity: 'base', numeric: true });
            if (firstCmp !== 0) {
                return firstCmp;
            }
            var lastCmp = aPair[1].localeCompare(bPair[1], undefined, { sensitivity: 'base', numeric: true });
            if (lastCmp !== 0) {
                return lastCmp;
            }
            return a.display.localeCompare(b.display, undefined, { sensitivity: 'base', numeric: true });
        });
        addLogMessage('buildDoAQueueSorted: unique=' + unique.length + ' duplicates=' + duplicateIndices.length, 'log');
        return { queue: unique, duplicateIndices: duplicateIndices };
    }

    function showDoAInputPanel() {
        addLogMessage('showDoAInputPanel: creating input panel', 'log');
        var modal = document.createElement('div');
        modal.id = 'doa-input-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 600px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'doa-input-title');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;';
        var title = document.createElement('h3');
        title.id = 'doa-input-title';
        title.textContent = DOA_LABELS.featureButton;
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600; letter-spacing: 0.2px;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close panel');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = function() {
            addLogMessage('showDoAInputPanel: closed by user', 'warn');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            stopDoA();
            if (doaState.focusReturnElement) {
                doaState.focusReturnElement.focus();
            }
        };
        header.appendChild(title);
        header.appendChild(closeButton);
        var description = document.createElement('p');
        description.textContent = 'Paste the DoA staff table below. Each row should contain name, codes, degree, email, role, and responsibility numbers (tab-separated).';
        description.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0 0 12px 0; font-size: 14px; line-height: 1.4;';
        var textarea = document.createElement('textarea');
        textarea.id = 'doa-entries-input';
        textarea.placeholder = 'Paste tab-separated staff table here...\nExample:\nPETER WINKLE\tPW\tMD\tp.winkle@cenexel.com\tPrincipal Investigator\tPI\t1 to 8\t13\t14';
        textarea.setAttribute('aria-label', 'DoA staff entries input');
        textarea.style.cssText = 'width: 100%; height: 200px; padding: 12px 14px; border: 2px solid rgba(255, 255, 255, 0.35); border-radius: 10px; background: rgba(255, 255, 255, 0.95); color: #1e293b; font-size: 13px; font-family: Consolas, monospace, Segoe UI, Tahoma, Geneva, Verdana, sans-serif; resize: vertical; outline: none; transition: all 0.25s ease; box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset; box-sizing: border-box;';
        textarea.onfocus = function() {
            textarea.style.borderColor = '#8ea0ff';
            textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)';
        };
        textarea.onblur = function() {
            textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)';
            textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset';
        };
        var confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.disabled = true;
        confirmButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; opacity: 0.5;';
        var updateConfirmState = function() {
            var parsed = parseDoAEntriesInput(textarea.value);
            if (parsed.length > 0) {
                confirmButton.disabled = false;
                confirmButton.style.opacity = '1';
                confirmButton.style.cursor = 'pointer';
            } else {
                confirmButton.disabled = true;
                confirmButton.style.opacity = '0.5';
                confirmButton.style.cursor = 'not-allowed';
            }
        };
        textarea.oninput = updateConfirmState;
        confirmButton.onmouseover = function() {
            if (!confirmButton.disabled) {
                confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
            }
        };
        confirmButton.onmouseout = function() {
            confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        confirmButton.onclick = function() {
            addLogMessage('showDoAInputPanel: Confirm clicked', 'log');
            var parsed = parseDoAEntriesInput(textarea.value);
            if (parsed.length === 0) {
                addLogMessage('showDoAInputPanel: no valid entries parsed', 'warn');
                return;
            }
            doaState.parsedCandidates = parsed;
            addLogMessage('showDoAInputPanel: parsed ' + parsed.length + ' candidates', 'log');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            doaState.isRunning = true;
            showCollectingDataPanel('doa', DOA_LABELS.featureButton);
            startDoAScan();
        };
        var clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease; backdrop-filter: blur(2px);';
        clearButton.onmouseover = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        clearButton.onmouseout = function() {
            clearButton.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        clearButton.onclick = function() {
            addLogMessage('showDoAInputPanel: Clear All clicked', 'log');
            textarea.value = '';
            doaState.parsedCandidates = [];
            confirmButton.disabled = true;
            confirmButton.style.opacity = '0.5';
            confirmButton.style.cursor = 'not-allowed';
        };
        var keyHandler = function(e) {
            if (e.key === 'Escape') {
                addLogMessage('showDoAInputPanel: Escape pressed', 'log');
                if (modal.parentNode) {
                    document.body.removeChild(modal);
                }
                stopDoA();
                if (doaState.focusReturnElement) {
                    doaState.focusReturnElement.focus();
                }
            }
        };
        document.addEventListener('keydown', keyHandler);
        doaState.eventListeners.push({ element: document, type: 'keydown', handler: keyHandler });
        var buttonContainer = document.createElement('div');
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
        addLogMessage('showDoAInputPanel: input panel displayed', 'log');
    }

    function initializeDoARightPanel() {
        addLogMessage('initializeDoARightPanel: initializing with parsed candidates', 'log');
        var rightPanel = document.getElementById('doa-right-panel');
        if (!rightPanel) {
            addLogMessage('initializeDoARightPanel: right panel not found', 'error');
            return;
        }
        rightPanel.innerHTML = '';
        for (var i = 0; i < doaState.parsedCandidates.length; i++) {
            var candidate = doaState.parsedCandidates[i];
            var numsArr = Array.from(candidate.numbersSet).sort(function(a, b) {
                return a - b;
            });
            var labelText = candidate.display + ' | ' + candidate.roleDisplay + ' | [' + numsArr.join(', ') + ']';
            var item = createListItem(labelText, DOA_LABELS.statusPending, 'pending', i + 1);
            item.setAttribute('data-pairkey', candidate.pairKey);
            rightPanel.appendChild(item);
        }
        addLogMessage('initializeDoARightPanel: added ' + doaState.parsedCandidates.length + ' items', 'log');
    }

    function updateDoARightPanelStatus(pairKey, newStatus) {
        addLogMessage('updateDoARightPanelStatus: pairKey=' + pairKey + ' status=' + newStatus, 'log');
        var rightPanel = document.getElementById('doa-right-panel');
        if (!rightPanel) {
            addLogMessage('updateDoARightPanelStatus: right panel not found', 'error');
            return;
        }
        var items = rightPanel.querySelectorAll('.' + ELOG_CSS_CLASSNAMES.listItem);
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.getAttribute('data-pairkey') === pairKey) {
                var badge = item.querySelector('.elog-status-badge');
                if (badge) {
                    badge.textContent = newStatus;
                    var badgeColor = 'rgba(255, 255, 255, 0.7)';
                    var badgeBg = 'rgba(255, 255, 255, 0.1)';
                    if (newStatus === DOA_LABELS.statusAdded) {
                        badgeColor = '#6bcf7f';
                        badgeBg = 'rgba(107, 207, 127, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusAlready) {
                        badgeColor = '#ffa500';
                        badgeBg = 'rgba(255, 165, 0, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusDuplicate) {
                        badgeColor = '#ffa500';
                        badgeBg = 'rgba(255, 165, 0, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusNotInDropdown || newStatus === DOA_LABELS.statusSelectionFailed || newStatus === DOA_LABELS.statusSaveFailed || newStatus === DOA_LABELS.statusRoleNotFound) {
                        badgeColor = '#ff6b6b';
                        badgeBg = 'rgba(255, 107, 107, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusTasksApplied) {
                        badgeColor = '#6bcf7f';
                        badgeBg = 'rgba(107, 207, 127, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusStopped) {
                        badgeColor = '#aaa';
                        badgeBg = 'rgba(170, 170, 170, 0.2)';
                    } else if (newStatus === DOA_LABELS.statusPending) {
                        badgeColor = '#ffd93d';
                        badgeBg = 'rgba(255, 217, 61, 0.2)';
                    }
                    badge.style.color = badgeColor;
                    badge.style.background = badgeBg;
                }
                break;
            }
        }
    }

    function updateDoARightPanelSummary(counters) {
        addLogMessage('updateDoARightPanelSummary: total=' + counters.total + ' added=' + counters.added + ' dup=' + counters.duplicates + ' fail=' + counters.failures + ' pending=' + counters.pending, 'log');
        var totalEl = document.getElementById('doa-summary-total');
        var addedEl = document.getElementById('doa-summary-added');
        var dupEl = document.getElementById('doa-summary-duplicates');
        var failEl = document.getElementById('doa-summary-failures');
        var pendingEl = document.getElementById('doa-summary-pending');
        var percentEl = document.getElementById('doa-summary-percent');
        if (totalEl) {
            totalEl.textContent = String(counters.total);
        }
        if (addedEl) {
            addedEl.textContent = String(counters.added);
        }
        if (dupEl) {
            dupEl.textContent = String(counters.duplicates);
        }
        if (failEl) {
            failEl.textContent = String(counters.failures);
        }
        if (pendingEl) {
            pendingEl.textContent = String(counters.pending);
        }
        if (percentEl) {
            var processed = counters.total - counters.pending;
            var pct = counters.total > 0 ? Math.round((processed / counters.total) * 100) : 0;
            percentEl.textContent = pct + '%';
        }
        updateDoAAriaLive('Added: ' + counters.added + ', Duplicates: ' + counters.duplicates + ', Pending: ' + counters.pending);
    }

    function updateDoAScanStatus(statusText, statusType) {
        addLogMessage('updateDoAScanStatus: ' + statusText, 'log');
        var badge = document.getElementById('doa-status-badge');
        var titleEl = document.getElementById('doa-progress-title');
        if (badge) {
            badge.textContent = statusText;
            if (statusType === 'complete') {
                badge.style.background = 'rgba(107, 207, 127, 0.3)';
                badge.style.color = '#6bcf7f';
            } else if (statusType === 'error') {
                badge.style.background = 'rgba(255, 107, 107, 0.3)';
                badge.style.color = '#ff6b6b';
            } else if (statusType === 'stopped') {
                badge.style.background = 'rgba(170, 170, 170, 0.3)';
                badge.style.color = '#aaa';
            } else {
                badge.style.background = 'rgba(255, 217, 61, 0.3)';
                badge.style.color = '#ffd93d';
            }
        }
        if (titleEl && statusType === 'complete') {
            titleEl.textContent = 'DoA Log Staff Entries - Complete';
        }
    }

    function showDoAProgressPanel() {
        addLogMessage('showDoAProgressPanel: creating progress panel', 'log');
        doaState.isRunning = true;
        var modal = document.createElement('div');
        modal.id = 'doa-progress-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 20000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.id = 'doa-progress-container';
        container.style.cssText = 'background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 24px; width: 900px; max-width: 95%; max-height: 80vh; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative; display: flex; flex-direction: column;';
        container.setAttribute('role', 'dialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'doa-progress-title');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0;';
        var titleContainer = document.createElement('div');
        titleContainer.style.cssText = 'display: flex; align-items: center; gap: 12px;';
        var titleEl = document.createElement('h3');
        titleEl.id = 'doa-progress-title';
        titleEl.textContent = 'DoA Log Staff Entries - Review';
        titleEl.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
        var statusBadge = document.createElement('span');
        statusBadge.id = 'doa-status-badge';
        statusBadge.textContent = 'In Progress';
        statusBadge.style.cssText = 'background: rgba(255, 255, 255, 0.3); color: #ffd93d; font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 12px;';
        titleContainer.appendChild(titleEl);
        titleContainer.appendChild(statusBadge);
        var headerButtons = document.createElement('div');
        headerButtons.style.cssText = 'display: flex; gap: 8px; align-items: center;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close and stop');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = function() {
            addLogMessage('showDoAProgressPanel: closed by user', 'warn');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            stopDoA();
            if (doaState.focusReturnElement) {
                doaState.focusReturnElement.focus();
            }
        };
        headerButtons.appendChild(closeButton);
        header.appendChild(titleContainer);
        header.appendChild(headerButtons);
        var panelsContainer = document.createElement('div');
        panelsContainer.style.cssText = 'display: grid; grid-template-columns: 1fr 1fr; gap: 16px; flex: 1; min-height: 0; overflow: hidden;';
        var leftPanel = createSubpanel('Scanned Log Entries', 'doa-left-panel', 'doa-left-search');
        var rightPanel = createSubpanel('DoA Entries Status', 'doa-right-panel', 'doa-right-search');
        panelsContainer.appendChild(leftPanel);
        panelsContainer.appendChild(rightPanel);
        var summaryFooter = document.createElement('div');
        summaryFooter.id = 'doa-summary-footer';
        summaryFooter.setAttribute('aria-label', 'Processing summary');
        summaryFooter.style.cssText = 'display: flex; justify-content: space-around; align-items: center; padding: 10px 16px; background: rgba(0, 0, 0, 0.2); border-radius: 8px; margin-top: 12px; flex-shrink: 0;';
        var summaryItems = [
            { id: 'doa-summary-total', label: 'Total', value: '0' },
            { id: 'doa-summary-added', label: 'Added', value: '0' },
            { id: 'doa-summary-duplicates', label: 'Duplicates', value: '0' },
            { id: 'doa-summary-failures', label: 'Failures', value: '0' },
            { id: 'doa-summary-pending', label: 'Pending', value: '0' },
            { id: 'doa-summary-percent', label: 'Progress', value: '0%' }
        ];
        for (var si = 0; si < summaryItems.length; si++) {
            var summaryItem = document.createElement('div');
            summaryItem.style.cssText = 'text-align: center;';
            var valSpan = document.createElement('span');
            valSpan.id = summaryItems[si].id;
            valSpan.textContent = summaryItems[si].value;
            valSpan.style.cssText = 'display: block; color: white; font-size: 16px; font-weight: 700;';
            var labelSpan = document.createElement('span');
            labelSpan.textContent = summaryItems[si].label;
            labelSpan.style.cssText = 'display: block; color: rgba(255, 255, 255, 0.6); font-size: 11px; font-weight: 500; margin-top: 2px;';
            summaryItem.appendChild(valSpan);
            summaryItem.appendChild(labelSpan);
            summaryFooter.appendChild(summaryItem);
        }
        var ariaLiveRegion = document.createElement('div');
        ariaLiveRegion.id = 'doa-aria-live';
        ariaLiveRegion.setAttribute('aria-live', 'polite');
        ariaLiveRegion.setAttribute('aria-atomic', 'true');
        ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        container.appendChild(header);
        container.appendChild(panelsContainer);
        container.appendChild(summaryFooter);
        container.appendChild(ariaLiveRegion);
        modal.appendChild(container);
        container.style.position = 'fixed';
        container.style.top = '50%';
        container.style.left = '50%';
        container.style.transform = 'translate(-50%, -50%)';
        modal.style.pointerEvents = 'none';
        container.style.pointerEvents = 'auto';
        makeDraggable(container, header);
        document.body.appendChild(modal);
        initializeDoARightPanel();
        populateDoALeftPanel();
        closeButton.focus();
        addLogMessage('showDoAProgressPanel: progress panel displayed with all data populated', 'log');
        updateDoAScanStatus('Scan Complete', 'complete');
        var progressTitle = document.getElementById('doa-progress-title');
        if (progressTitle) {
            progressTitle.textContent = 'DoA Log Staff Entries - Adding Entries';
        }
        beginAddDoALogEntries();
    }

    function updateDoALeftPanelList(name, rowNumber) {
        var leftPanel = document.getElementById('doa-left-panel');
        if (!leftPanel) {
            return;
        }
        var item = createListItem(name, null, null, rowNumber);
        leftPanel.appendChild(item);
        var searchInput = document.getElementById('doa-left-search');
        if (searchInput && searchInput.value.trim()) {
            filterSubpanelList('doa-left-panel', searchInput.value);
        }
    }

    function populateDoALeftPanel() {
        addLogMessage('populateDoALeftPanel: populating with ' + doaState.scannedNames.length + ' names', 'log');
        var leftPanel = document.getElementById('doa-left-panel');
        if (!leftPanel) {
            addLogMessage('populateDoALeftPanel: left panel not found', 'error');
            return;
        }
        leftPanel.innerHTML = '';
        for (var i = 0; i < doaState.scannedNames.length; i++) {
            var item = createListItem(doaState.scannedNames[i], null, null, i + 1);
            leftPanel.appendChild(item);
        }
        addLogMessage('populateDoALeftPanel: populated ' + doaState.scannedNames.length + ' items', 'log');
    }

    function checkAndUpdateDoARightPanel(scannedName) {
    }

    function doaScanVisibleRowsOnce(onRow) {
        var gridTable = document.querySelector(DOA_SELECTORS.mainGridTable);
        if (!gridTable) {
            return 0;
        }
        var rows = gridTable.querySelectorAll(DOA_SELECTORS.mainGridRow);
        var newCount = 0;
        var batchNames = [];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            if (row.getAttribute('role') === 'columnheader') {
                continue;
            }
            var cells = row.querySelectorAll(DOA_SELECTORS.mainGridCell);
            if (cells.length <= ELOG_SELECTORS.nameCellIndex) {
                continue;
            }
            var targetCell = cells[ELOG_SELECTORS.nameCellIndex];
            var primaryElement = targetCell.querySelector(ELOG_SELECTORS.namePrimary);
            var extractedName = null;
            if (primaryElement) {
                var brElement = primaryElement.querySelector('br');
                if (brElement) {
                    var nameText = '';
                    for (var ni = 0; ni < primaryElement.childNodes.length; ni++) {
                        var node = primaryElement.childNodes[ni];
                        if (node.nodeName === 'BR') {
                            break;
                        }
                        if (node.nodeType === Node.TEXT_NODE) {
                            nameText += node.textContent;
                        }
                    }
                    extractedName = nameText.trim().replace(/\s+/g, ' ');
                } else {
                    extractedName = primaryElement.textContent.trim().replace(/\s+/g, ' ');
                }
            }
            if (!extractedName) {
                var fallbackElement = targetCell.querySelector(ELOG_SELECTORS.nameFallback);
                if (fallbackElement) {
                    extractedName = fallbackElement.textContent.trim().replace(/\s+/g, ' ');
                }
            }
            if (!extractedName) {
                continue;
            }
            var normalized = elogNormalizeName(extractedName);
            if (doaState.seenNormalizedNames.has(normalized)) {
                continue;
            }
            doaState.seenNormalizedNames.add(normalized);
            doaState.scannedNames.push(extractedName);
            newCount++;
            batchNames.push(extractedName);
            if (onRow) {
                onRow(extractedName, normalized);
            }
        }
        if (batchNames.length > 0) {
            addLogMessage('doaScanVisibleRowsOnce: batch of ' + batchNames.length + ' new names', 'log');
        }
        return newCount;
    }

    function doaObserveUserScrollPause(container) {
        if (!container || doaState.userScrollHandler) {
            return;
        }
        doaState.userScrollHandler = function() {
            var timeSinceAuto = Date.now() - doaState.lastAutoScrollTime;
            if (timeSinceAuto > 50 && !doaState.userScrollPaused) {
                doaState.userScrollPaused = true;
                var resumeTimeout = setTimeout(function() {
                    doaState.userScrollPaused = false;
                }, ELOG_SCROLL.userScrollPauseMs);
                doaState.timeouts.push(resumeTimeout);
            }
        };
        container.addEventListener('scroll', doaState.userScrollHandler);
        doaState.eventListeners.push({ element: container, type: 'scroll', handler: doaState.userScrollHandler });
    }

    function doaAwaitSettle(container) {
        return new Promise(function(resolve) {
            var gridTable = document.querySelector(DOA_SELECTORS.mainGridTable);
            var resolved = false;
            var observer = null;
            var timeoutId = setTimeout(function() {
                if (!resolved) {
                    resolved = true;
                    if (observer) {
                        observer.disconnect();
                        var idx = doaState.observers.indexOf(observer);
                        if (idx > -1) {
                            doaState.observers.splice(idx, 1);
                        }
                    }
                    resolve();
                }
            }, ELOG_SCROLL.settleDelayMs);
            doaState.timeouts.push(timeoutId);
            if (gridTable) {
                observer = new MutationObserver(function() {
                    if (!resolved) {
                        resolved = true;
                        clearTimeout(timeoutId);
                        observer.disconnect();
                        var idx = doaState.observers.indexOf(observer);
                        if (idx > -1) {
                            doaState.observers.splice(idx, 1);
                        }
                        resolve();
                    }
                });
                observer.observe(gridTable, { childList: true, subtree: true });
                doaState.observers.push(observer);
            }
        });
    }

    function doaGetRenderedLastRowKey() {
        var gridTable = document.querySelector(DOA_SELECTORS.mainGridTable);
        if (!gridTable) {
            return '';
        }
        var rows = gridTable.querySelectorAll(DOA_SELECTORS.mainGridRow);
        var lastDataRow = null;
        for (var i = rows.length - 1; i >= 0; i--) {
            if (rows[i].getAttribute('role') !== 'columnheader') {
                lastDataRow = rows[i];
                break;
            }
        }
        if (!lastDataRow) {
            return '';
        }
        var cells = lastDataRow.querySelectorAll(DOA_SELECTORS.mainGridCell);
        if (cells.length > 0) {
            var firstCellText = cells[0].textContent.trim();
            if (firstCellText) {
                return firstCellText;
            }
        }
        var hash = 0;
        var content = lastDataRow.textContent || '';
        for (var ci = 0; ci < content.length; ci++) {
            hash = ((hash << 5) - hash) + content.charCodeAt(ci);
            hash = hash & hash;
        }
        return 'hash_' + hash;
    }

    function doaGetRenderedRowCount() {
        var gridTable = document.querySelector(DOA_SELECTORS.mainGridTable);
        if (!gridTable) {
            return 0;
        }
        var rows = gridTable.querySelectorAll(DOA_SELECTORS.mainGridRow);
        var count = 0;
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].getAttribute('role') !== 'columnheader') {
                count++;
            }
        }
        return count;
    }

    function startDoAScan() {
        addLogMessage('startDoAScan: beginning scan with auto-scroll', 'log');
        doaState.scannedNames = [];
        doaState.seenNormalizedNames = new Set();
        doaWaitForElement(DOA_SELECTORS.mainTableContainer, ELOG_TIMEOUTS.waitTableMs)
            .then(function() {
            addLogMessage('startDoAScan: main table found', 'log');
            return doaWaitForElement(DOA_SELECTORS.mainGridTable, ELOG_TIMEOUTS.waitGridMs);
        })
            .then(function(gridTable) {
            addLogMessage('startDoAScan: grid table found, starting auto-scroll scan', 'log');
            doaAutoScrollScan({
                onRow: function(name, normalized) {
                },
                onProgress: function(data) {
                },
                onDone: function(data) {
                    addLogMessage('startDoAScan: done - total=' + data.total + ' reason=' + data.reason, 'log');
                    removeCollectingDataPanel('doa');
                    if (data.reason !== 'stopped' && doaState.isRunning) {
                        addLogMessage('startDoAScan: scan complete, showing progress panel', 'log');
                        showDoAProgressPanel();
                    }
                },
                onError: function(error) {
                    addLogMessage('startDoAScan: error - ' + error.message, 'error');
                    removeCollectingDataPanel('doa');
                    showDoAProgressPanel();
                    updateDoAScanStatus('Error', 'error');
                }
            });
        })
            .catch(function(error) {
            addLogMessage('startDoAScan: error during scan: ' + error, 'error');
            removeCollectingDataPanel('doa');
            showDoAProgressPanel();
            updateDoAScanStatus('Error', 'error');
        });
    }

    function doaAutoScrollScan(options) {
        addLogMessage('doaAutoScrollScan: starting', 'log');
        var onRow = options.onRow || function() {};
        var onProgress = options.onProgress || function() {};
        var onDone = options.onDone || function() {};
        var onError = options.onError || function() {};
        var gridTable = document.querySelector(DOA_SELECTORS.mainGridTable);
        if (!gridTable) {
            addLogMessage('doaAutoScrollScan: grid not found', 'error');
            onError(new Error('Grid table not found'));
            return;
        }
        var container = findScrollableContainer(gridTable);
        if (!container) {
            addLogMessage('doaAutoScrollScan: container not found', 'error');
            onError(new Error('Container not found'));
            return;
        }
        doaState.scrollContainer = container;
        doaState.prevScrollTop = container.scrollTop;
        doaState.seenNormalizedNames = new Set();
        doaState.leftPanelRowIndex = 0;
        doaSetAriaBusyOn();
        doaObserveUserScrollPause(container);
        var startTime = Date.now();
        var noProgress = 0;
        doaScanVisibleRowsOnce(function(name) {
            doaState.leftPanelRowIndex++;
        });
        onProgress({ scanned: doaState.scannedNames.length });
        function scrollLoop() {
            if (!doaState.isRunning) {
                updateDoAAriaLive('Scan stopped');
                finishScan('Scan stopped', 'stopped');
                return;
            }
            if (Date.now() - startTime > ELOG_SCROLL.maxDurationMs) {
                finishScan('Scan complete', 'timeout');
                return;
            }
            if (doaState.userScrollPaused) {
                var pt = setTimeout(scrollLoop, 100);
                doaState.timeouts.push(pt);
                return;
            }
            var priorKey = doaGetRenderedLastRowKey();
            var priorCount = doaGetRenderedRowCount();
            var currTop = container.scrollTop;
            var maxScroll = container.scrollHeight - container.clientHeight;
            var newTop = Math.min(currTop + ELOG_SCROLL.stepPx, maxScroll);
            doaState.lastAutoScrollTime = Date.now();
            container.scrollTo({ top: newTop, behavior: 'auto' });
            doaAwaitSettle(container).then(function() {
                var attempts = 0;
                function attemptScan() {
                    var rc = doaGetRenderedRowCount();
                    if (rc === 0 && attempts < ELOG_SCROLL.retryScanAttempts) {
                        attempts++;
                        var rt = setTimeout(attemptScan, ELOG_SCROLL.retryScanDelayMs);
                        doaState.timeouts.push(rt);
                        return;
                    }
                    doaScanVisibleRowsOnce(function(name) {
                        doaState.leftPanelRowIndex++;
                    });
                    onProgress({ scanned: doaState.scannedNames.length });
                    var currKey = doaGetRenderedLastRowKey();
                    var currCount = doaGetRenderedRowCount();
                    if (currKey === priorKey && currCount === priorCount) {
                        noProgress++;
                    } else {
                        noProgress = 0;
                    }
                    if (computeEndReached(container, noProgress)) {
                        finishScan('Scan complete', 'endReached');
                        return;
                    }
                    if (noProgress >= ELOG_SCROLL.maxNoProgressIterations) {
                        finishScan('Scan complete', 'noProgress');
                        return;
                    }
                    if (typeof requestIdleCallback === 'function') {
                        var icbId = requestIdleCallback(function() {
                            var idx = doaState.idleCallbackIds.indexOf(icbId);
                            if (idx > -1) {
                                doaState.idleCallbackIds.splice(idx, 1);
                            }
                            scrollLoop();
                        }, { timeout: ELOG_SCROLL.idleDelayMs * 2 });
                        doaState.idleCallbackIds.push(icbId);
                    } else {
                        var it = setTimeout(scrollLoop, ELOG_SCROLL.idleDelayMs);
                        doaState.timeouts.push(it);
                    }
                }
                attemptScan();
            });
        }
        function finishScan(label, reason) {
            addLogMessage('doaAutoScrollScan: done reason=' + reason + ' total=' + doaState.scannedNames.length, 'log');
            doaSetAriaBusyOff();
            onDone({ total: doaState.scannedNames.length, reason: reason });
        }
        var initTimeout = setTimeout(scrollLoop, ELOG_SCROLL.idleDelayMs);
        doaState.timeouts.push(initTimeout);
    }

    function ensureDoAAddEntryFormOpen() {
        addLogMessage('ensureDoAAddEntryFormOpen: checking form state', 'log');
        return new Promise(function(resolve, reject) {
            var memberInput = document.querySelector(DOA_SELECTORS.memberInput);
            if (memberInput) {
                addLogMessage('ensureDoAAddEntryFormOpen: member input already present', 'log');
                resolve(memberInput);
                return;
            }
            var addBtn = document.querySelector(DOA_SELECTORS.addEntryBtn);
            if (addBtn) {
                addLogMessage('ensureDoAAddEntryFormOpen: clicking Add Entry button', 'log');
                addBtn.click();
            } else {
                addLogMessage('ensureDoAAddEntryFormOpen: Add Entry button not found', 'warn');
            }
            doaWaitForElement(DOA_SELECTORS.memberInput, DOA_TIMEOUTS.waitOpenMs)
                .then(function(el) {
                addLogMessage('ensureDoAAddEntryFormOpen: member input found', 'log');
                resolve(el);
            })
                .catch(function(err) {
                addLogMessage('ensureDoAAddEntryFormOpen: timeout: ' + err, 'error');
                reject(err);
            });
        });
    }

    function ensureDoAMemberDropdownOpen() {
        addLogMessage('ensureDoAMemberDropdownOpen: checking dropdown', 'log');
        return new Promise(function(resolve, reject) {
            var retries = 0;
            function tryOpen() {
                var listEl = document.querySelector(DOA_SELECTORS.listContainer);
                if (listEl) {
                    addLogMessage('ensureDoAMemberDropdownOpen: list already open', 'log');
                    resolve(listEl);
                    return;
                }
                var inputEl = document.querySelector(DOA_SELECTORS.memberInput);
                if (inputEl) {
                    addLogMessage('ensureDoAMemberDropdownOpen: clicking input to open list', 'log');
                    inputEl.click();
                    inputEl.focus();
                }
                doaWaitForElement(DOA_SELECTORS.listContainer, DOA_TIMEOUTS.waitListMs)
                    .then(function(el) {
                    addLogMessage('ensureDoAMemberDropdownOpen: list opened', 'log');
                    resolve(el);
                })
                    .catch(function(err) {
                    retries++;
                    if (retries < DOA_RETRY.openListRetries) {
                        addLogMessage('ensureDoAMemberDropdownOpen: retry ' + retries, 'warn');
                        var tid = setTimeout(tryOpen, 300);
                        doaState.timeouts.push(tid);
                    } else {
                        addLogMessage('ensureDoAMemberDropdownOpen: exhausted retries', 'error');
                        reject(err);
                    }
                });
            }
            tryOpen();
        });
    }

    function findDoADropdownViewportElement() {
        var el = document.querySelector(DOA_SELECTORS.virtualViewport);
        if (el) {
            addLogMessage('findDoADropdownViewportElement: found', 'log');
        } else {
            addLogMessage('findDoADropdownViewportElement: not found', 'warn');
        }
        return el;
    }

    function rememberDoAListScrollPosition(viewportEl) {
        if (viewportEl) {
            doaState.listScrollTop = viewportEl.scrollTop;
        }
    }

    function restoreDoAListScrollPosition(viewportEl) {
        if (viewportEl && doaState.listScrollTop > 0) {
            viewportEl.scrollTop = doaState.listScrollTop;
        }
    }

    function attemptDoASelectByScrollingForName(targetDisplay, targetPairKey) {
        addLogMessage('attemptDoASelectByScrollingForName: target=' + targetDisplay + ' pairKey=' + targetPairKey, 'log');
        return new Promise(function(resolve) {
            var inputEl = document.querySelector(DOA_SELECTORS.memberInput);
            if (!inputEl) {
                addLogMessage('attemptDoASelectByScrollingForName: member input not found', 'error');
                resolve(false);
                return;
            }
            var nameParts = getFirstAndLast(targetDisplay);
            var firstName = nameParts[0];
            var lastName = nameParts[1];
            addLogMessage('attemptDoASelectByScrollingForName: firstName=' + firstName + ' lastName=' + lastName, 'log');
            typeIntoFilteredInput(inputEl, firstName);
            var firstNameTid = setTimeout(function() {
                if (!doaState.isRunning) {
                    clearFilteredInput(inputEl);
                    resolve(false);
                    return;
                }
                var match = scanFilteredOptionsForMatch(targetPairKey, DOA_SELECTORS);
                if (match) {
                    addLogMessage('attemptDoASelectByScrollingForName: found via firstName filter (' + match.matchType + ')', 'log');
                    match.element.click();
                    var verifyTid1 = setTimeout(function() {
                        resolve(true);
                    }, DOA_TIMEOUTS.settleMs);
                    doaState.timeouts.push(verifyTid1);
                    return;
                }
                addLogMessage('attemptDoASelectByScrollingForName: not found with firstName, trying lastName', 'log');
                if (lastName && lastName !== firstName) {
                    typeIntoFilteredInput(inputEl, lastName);
                    var lastNameTid = setTimeout(function() {
                        if (!doaState.isRunning) {
                            clearFilteredInput(inputEl);
                            resolve(false);
                            return;
                        }
                        var match2 = scanFilteredOptionsForMatch(targetPairKey, DOA_SELECTORS);
                        if (match2) {
                            addLogMessage('attemptDoASelectByScrollingForName: found via lastName filter (' + match2.matchType + ')', 'log');
                            match2.element.click();
                            var verifyTid2 = setTimeout(function() {
                                resolve(true);
                            }, DOA_TIMEOUTS.settleMs);
                            doaState.timeouts.push(verifyTid2);
                            return;
                        }
                        addLogMessage('attemptDoASelectByScrollingForName: not found with lastName, falling back to scroll', 'log');
                        clearFilteredInput(inputEl);
                        var scrollFallbackTid = setTimeout(function() {
                            scrollSearchForName(targetPairKey, DOA_SELECTORS, DOA_TIMEOUTS, DOA_RETRY, doaState, resolve);
                        }, DOA_TIMEOUTS.waitFilterMs);
                        doaState.timeouts.push(scrollFallbackTid);
                    }, DOA_TIMEOUTS.waitFilterMs);
                    doaState.timeouts.push(lastNameTid);
                } else {
                    addLogMessage('attemptDoASelectByScrollingForName: lastName same as firstName, falling back to scroll', 'log');
                    clearFilteredInput(inputEl);
                    var scrollTid = setTimeout(function() {
                        scrollSearchForName(targetPairKey, DOA_SELECTORS, DOA_TIMEOUTS, DOA_RETRY, doaState, resolve);
                    }, DOA_TIMEOUTS.waitFilterMs);
                    doaState.timeouts.push(scrollTid);
                }
            }, DOA_TIMEOUTS.waitFilterMs);
            doaState.timeouts.push(firstNameTid);
        });
    }

    function selectRoleForCandidate(roleDisplay, roleKey) {
        addLogMessage('selectRoleForCandidate: roleDisplay=' + roleDisplay + ' roleKey=' + roleKey, 'log');
        return new Promise(function(resolve) {
            var roleInput = document.querySelector(DOA_SELECTORS.roleSearchInput);
            if (!roleInput) {
                addLogMessage('selectRoleForCandidate: role search input not found', 'warn');
                resolve(false);
                return;
            }
            roleInput.click();
            roleInput.focus();
            doaWaitForElement(DOA_SELECTORS.listContainer, DOA_TIMEOUTS.waitRoleListMs)
                .then(function() {
                addLogMessage('selectRoleForCandidate: role list opened', 'log');
                var viewportEl = document.querySelector(DOA_SELECTORS.virtualViewport);
                if (viewportEl && doaState.roleListScrollTop > 0) {
                    viewportEl.scrollTop = doaState.roleListScrollTop;
                }
                var passCount = 0;
                var lastScrollTop = -1;
                var lastOptionSnapshot = '';
                var startTime = Date.now();
                function scanRoles() {
                    if (!doaState.isRunning) {
                        resolve(false);
                        return;
                    }
                    if (Date.now() - startTime > DOA_TIMEOUTS.maxSelectDurationMs) {
                        addLogMessage('selectRoleForCandidate: max duration exceeded', 'warn');
                        resolve(false);
                        return;
                    }
                    if (passCount >= DOA_RETRY.maxScrollPasses) {
                        addLogMessage('selectRoleForCandidate: max passes reached', 'warn');
                        resolve(false);
                        return;
                    }
                    var roleOptions = document.querySelectorAll(DOA_SELECTORS.roleOptionItem);
                    var found = false;
                    var optionTexts = [];
                    for (var ri = 0; ri < roleOptions.length; ri++) {
                        var textEl = roleOptions[ri].querySelector(DOA_SELECTORS.roleOptionText);
                        var optText = textEl ? textEl.textContent.trim() : roleOptions[ri].textContent.trim();
                        optionTexts.push(optText);
                        var normalized = normalizeRoleName(optText);
                        if (normalized.key === roleKey) {
                            addLogMessage('selectRoleForCandidate: match found at index ' + ri + ' text=' + optText, 'log');
                            roleOptions[ri].click();
                            found = true;
                            var vp = document.querySelector(DOA_SELECTORS.virtualViewport);
                            if (vp) {
                                doaState.roleListScrollTop = vp.scrollTop;
                            }
                            var verifyTid = setTimeout(function() {
                                resolve(true);
                            }, DOA_TIMEOUTS.settleMs);
                            doaState.timeouts.push(verifyTid);
                            break;
                        }
                    }
                    if (!found) {
                        var vp2 = document.querySelector(DOA_SELECTORS.virtualViewport);
                        if (!vp2) {
                            addLogMessage('selectRoleForCandidate: no viewport for role scroll', 'warn');
                            resolve(false);
                            return;
                        }
                        var currentSnapshot = optionTexts.join('|');
                        var currentTop = vp2.scrollTop;
                        var noProgressFlag = (currentTop === lastScrollTop && currentSnapshot === lastOptionSnapshot);
                        if (noProgressFlag) {
                            passCount++;
                        }
                        lastScrollTop = currentTop;
                        lastOptionSnapshot = currentSnapshot;
                        var stepSize = Math.round(vp2.clientHeight * 0.7);
                        var maxScroll = vp2.scrollHeight - vp2.clientHeight;
                        var newTop = Math.min(currentTop + stepSize, maxScroll);
                        if (newTop <= currentTop && currentTop > 0) {
                            newTop = 0;
                            lastScrollTop = -1;
                            lastOptionSnapshot = '';
                            passCount++;
                        }
                        vp2.scrollTop = newTop;
                        var scrollTid = setTimeout(scanRoles, DOA_TIMEOUTS.scrollIdleMs);
                        doaState.timeouts.push(scrollTid);
                    }
                }
                var initTid = setTimeout(scanRoles, DOA_TIMEOUTS.scrollIdleMs);
                doaState.timeouts.push(initTid);
            })
                .catch(function(err) {
                addLogMessage('selectRoleForCandidate: timeout opening role list: ' + err, 'error');
                resolve(false);
            });
        });
    }

    function openAndApplyStudyTasks(numbersSet) {
        addLogMessage('openAndApplyStudyTasks: numbersSet size=' + numbersSet.size, 'log');
        return new Promise(function(resolve) {
            var toggleBtn = document.querySelector(DOA_SELECTORS.tasksToggleBtn);
            if (!toggleBtn) {
                addLogMessage('openAndApplyStudyTasks: toggle button not found', 'warn');
                resolve(false);
                return;
            }
            addLogMessage('openAndApplyStudyTasks: clicking toggle button', 'log');
            toggleBtn.click();
            doaWaitForElement(DOA_SELECTORS.tasksMenu, DOA_TIMEOUTS.waitTasksMenuMs)
                .then(function(menu) {
                addLogMessage('openAndApplyStudyTasks: tasks menu opened', 'log');
                var items = menu.querySelectorAll(DOA_SELECTORS.tasksItem);
                addLogMessage('openAndApplyStudyTasks: found ' + items.length + ' task items', 'log');
                var toggleQueue = [];
                for (var ti = 0; ti < items.length; ti++) {
                    var itemText = items[ti].textContent.trim();
                    var leadingNumMatch = itemText.match(/^(\d+)\./);
                    if (!leadingNumMatch) {
                        continue;
                    }
                    var taskNum = parseInt(leadingNumMatch[1], 10);
                    var checkbox = items[ti].querySelector(DOA_SELECTORS.tasksItemCheckbox);
                    if (!checkbox) {
                        continue;
                    }
                    var isChecked = checkbox.checked;
                    var shouldBeChecked = numbersSet.has(taskNum);
                    if (shouldBeChecked && !isChecked) {
                        toggleQueue.push({ element: checkbox, taskNum: taskNum, action: 'check' });
                    } else if (!shouldBeChecked && isChecked) {
                        toggleQueue.push({ element: checkbox, taskNum: taskNum, action: 'uncheck' });
                    }
                }
                addLogMessage('openAndApplyStudyTasks: ' + toggleQueue.length + ' toggles needed', 'log');
                var tqi = 0;
                function processNextToggle() {
                    if (!doaState.isRunning) {
                        resolve(false);
                        return;
                    }
                    if (tqi >= toggleQueue.length) {
                        addLogMessage('openAndApplyStudyTasks: all toggles applied', 'log');
                        var closeBtn = document.querySelector(DOA_SELECTORS.tasksToggleBtn);
                        if (closeBtn) {
                            closeBtn.click();
                        }
                        var closeTid = setTimeout(function() {
                            resolve(true);
                        }, DOA_TIMEOUTS.waitAfterTasksToggleMs);
                        doaState.timeouts.push(closeTid);
                        return;
                    }
                    var toggleItem = toggleQueue[tqi];
                    toggleItem.element.click();
                    tqi++;
                    var nextTid = setTimeout(processNextToggle, DOA_TIMEOUTS.waitAfterTasksToggleMs);
                    doaState.timeouts.push(nextTid);
                }
                var startTid = setTimeout(processNextToggle, DOA_TIMEOUTS.waitAfterTasksToggleMs);
                doaState.timeouts.push(startTid);
            })
                .catch(function(err) {
                addLogMessage('openAndApplyStudyTasks: timeout opening tasks menu: ' + err, 'error');
                resolve(false);
            });
        });
    }

    function processNextDoAFromQueue() {
        addLogMessage('processNextDoAFromQueue: index=' + doaState.addQueueIndex + ' of ' + doaState.addQueue.length, 'log');
        if (!doaState.isRunning) {
            addLogMessage('processNextDoAFromQueue: stopped, marking remaining as Stopped', 'warn');
            for (var si = doaState.addQueueIndex; si < doaState.addQueue.length; si++) {
                if (doaState.addQueue[si].status === DOA_LABELS.statusPending) {
                    doaState.addQueue[si].status = DOA_LABELS.statusStopped;
                    updateDoARightPanelStatus(doaState.addQueue[si].pairKey, DOA_LABELS.statusStopped);
                    doaState.counters.pending--;
                }
            }
            updateDoARightPanelSummary(doaState.counters);
            doaState.isAddingEntries = false;
            updateDoAScanStatus('Stopped', 'stopped');
            var titleEl = document.getElementById('doa-progress-title');
            if (titleEl) {
                titleEl.textContent = 'DoA Log Staff Entries - Stopped';
            }
            return;
        }
        if (doaState.addQueueIndex >= doaState.addQueue.length) {
            addLogMessage('processNextDoAFromQueue: all candidates processed', 'log');
            doaState.isAddingEntries = false;
            updateDoAScanStatus('Complete', 'complete');
            var titleEl2 = document.getElementById('doa-progress-title');
            if (titleEl2) {
                titleEl2.textContent = 'DoA Log Staff Entries - Complete';
            }
            updateDoARightPanelSummary(doaState.counters);
            updateDoAAriaLive('Processing complete. Added: ' + doaState.counters.added + ', Duplicates: ' + doaState.counters.duplicates + ', Failures: ' + doaState.counters.failures);
            return;
        }
        var candidate = doaState.addQueue[doaState.addQueueIndex];
        addLogMessage('processNextDoAFromQueue: processing ' + candidate.display + ' role=' + candidate.roleDisplay, 'log');
        if (doaState.existingPairs.has(candidate.pairKey)) {
            addLogMessage('processNextDoAFromQueue: already exists in table', 'log');
            candidate.status = DOA_LABELS.statusAlready;
            doaState.counters.duplicates++;
            doaState.counters.pending--;
            updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusAlready);
            updateDoARightPanelSummary(doaState.counters);
            doaState.addQueueIndex++;
            var tid1 = setTimeout(processNextDoAFromQueue, 50);
            doaState.timeouts.push(tid1);
            return;
        }
        ensureDoAAddEntryFormOpen()
            .then(function() {
            addLogMessage('processNextDoAFromQueue: form open, opening dropdown', 'log');
            return ensureDoAMemberDropdownOpen();
        })
            .then(function() {
            addLogMessage('processNextDoAFromQueue: dropdown open, selecting name', 'log');
            return attemptDoASelectByScrollingForName(candidate.display, candidate.pairKey);
        })
            .then(function(selected) {
            if (!selected) {
                addLogMessage('processNextDoAFromQueue: not found in dropdown', 'warn');
                candidate.status = DOA_LABELS.statusNotInDropdown;
                doaState.counters.failures++;
                doaState.counters.pending--;
                updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusNotInDropdown);
                updateDoARightPanelSummary(doaState.counters);
                doaState.addQueueIndex++;
                var tid2 = setTimeout(processNextDoAFromQueue, 50);
                doaState.timeouts.push(tid2);
                return;
            }
            addLogMessage('processNextDoAFromQueue: name selected, selecting role', 'log');
            return doaDelay(DOA_TIMEOUTS.settleMs).then(function() {
                return selectRoleForCandidate(candidate.roleDisplay, candidate.roleKey);
            }).then(function(roleSelected) {
                if (!roleSelected) {
                    addLogMessage('processNextDoAFromQueue: role not found', 'warn');
                    candidate.status = DOA_LABELS.statusRoleNotFound;
                    doaState.counters.failures++;
                    doaState.counters.pending--;
                    updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusRoleNotFound);
                    updateDoARightPanelSummary(doaState.counters);
                    doaState.addQueueIndex++;
                    var tid3 = setTimeout(processNextDoAFromQueue, 50);
                    doaState.timeouts.push(tid3);
                    return;
                }
                addLogMessage('processNextDoAFromQueue: role selected, applying tasks', 'log');
                return doaDelay(DOA_TIMEOUTS.settleMs).then(function() {
                    return openAndApplyStudyTasks(candidate.numbersSet);
                }).then(function(tasksApplied) {
                    if (!tasksApplied) {
                        addLogMessage('processNextDoAFromQueue: tasks failed', 'warn');
                        candidate.status = DOA_LABELS.statusSelectionFailed;
                        doaState.counters.failures++;
                        doaState.counters.pending--;
                        updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusSelectionFailed);
                    } else {
                        addLogMessage('processNextDoAFromQueue: tasks applied', 'log');
                        candidate.status = DOA_LABELS.statusTasksApplied;
                        doaState.counters.added++;
                        doaState.counters.pending--;
                        updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusTasksApplied);
                    }
                    updateDoARightPanelSummary(doaState.counters);
                    doaState.addQueueIndex++;
                    var tid4 = setTimeout(processNextDoAFromQueue, 50);
                    doaState.timeouts.push(tid4);
                });
            });
        })
            .catch(function(err) {
            addLogMessage('processNextDoAFromQueue: error: ' + err, 'error');
            candidate.status = DOA_LABELS.statusSelectionFailed;
            doaState.counters.failures++;
            doaState.counters.pending--;
            updateDoARightPanelStatus(candidate.pairKey, DOA_LABELS.statusSelectionFailed);
            updateDoARightPanelSummary(doaState.counters);
            doaState.addQueueIndex++;
            var tid5 = setTimeout(processNextDoAFromQueue, 50);
            doaState.timeouts.push(tid5);
        });
    }

    function beginAddDoALogEntries() {
        addLogMessage('beginAddDoALogEntries: starting', 'log');
        doaState.existingPairs = buildExistingPairsFromScan(doaState.scannedNames);
        addLogMessage('beginAddDoALogEntries: existingPairs count=' + doaState.existingPairs.size, 'log');
        var result = buildDoAQueueSorted(doaState.parsedCandidates);
        doaState.addQueue = result.queue;
        doaState.addQueueIndex = 0;
        doaState.listScrollTop = 0;
        doaState.roleListScrollTop = 0;
        doaState.isAddingEntries = true;
        for (var di = 0; di < result.duplicateIndices.length; di++) {
            var dupIdx = result.duplicateIndices[di];
            var dupCandidate = doaState.parsedCandidates[dupIdx];
            if (dupCandidate) {
                updateDoARightPanelStatus(dupCandidate.pairKey, DOA_LABELS.statusDuplicate);
            }
        }
        doaState.counters = {
            total: doaState.addQueue.length,
            added: 0,
            duplicates: 0,
            failures: 0,
            pending: doaState.addQueue.length
        };
        updateDoARightPanelSummary(doaState.counters);
        updateDoAScanStatus('Adding Entries', 'progress');
        var titleEl = document.getElementById('doa-progress-title');
        if (titleEl) {
            titleEl.textContent = 'DoA Log Staff Entries - Adding Entries';
        }
        updateDoAAriaLive('Starting to add ' + doaState.addQueue.length + ' entries');
        addLogMessage('beginAddDoALogEntries: queue size=' + doaState.addQueue.length + ', starting processing', 'log');
        processNextDoAFromQueue();
    }

    function addDoALogStaffEntriesInit() {
        addLogMessage('addDoALogStaffEntriesInit: starting feature', 'log');
        doaState.focusReturnElement = document.getElementById('doa-staff-entries-btn');
        doaState.abortController = new AbortController();
        resetDoAState();
        var mainTable = document.querySelector(DOA_SELECTORS.mainTableContainer);
        addLogMessage('addDoALogStaffEntriesInit: checking for main table', 'log');
        if (!mainTable) {
            addLogMessage('addDoALogStaffEntriesInit: main table not found, showing warning', 'warn');
            showDoAWarning();
            return;
        }
        addLogMessage('addDoALogStaffEntriesInit: main table found, showing input panel', 'log');
        showDoAInputPanel();
    }

    function showDoAWarning() {
        addLogMessage('showDoAWarning: creating warning popup', 'log');
        var modal = document.createElement('div');
        modal.id = 'doa-warning-modal';
        modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 30000; display: flex; align-items: center; justify-content: center;';
        var container = document.createElement('div');
        container.style.cssText = 'background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); border-radius: 12px; padding: 24px; width: 450px; max-width: 90%; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); position: relative;';
        container.setAttribute('role', 'alertdialog');
        container.setAttribute('aria-modal', 'true');
        container.setAttribute('aria-labelledby', 'doa-warning-title');
        container.setAttribute('aria-describedby', 'doa-warning-message');
        var header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;';
        var title = document.createElement('h3');
        title.id = 'doa-warning-title';
        title.textContent = 'Document Log Not Found';
        title.style.cssText = 'margin: 0; color: white; font-size: 18px; font-weight: 600;';
        var closeButton = document.createElement('button');
        closeButton.innerHTML = '\u2715';
        closeButton.setAttribute('aria-label', 'Close warning');
        closeButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: none; color: white; width: 28px; height: 28px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.3s ease;';
        closeButton.onmouseover = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.3)';
        };
        closeButton.onmouseout = function() {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        var closeWarning = function() {
            addLogMessage('showDoAWarning: closing warning', 'log');
            if (modal.parentNode) {
                document.body.removeChild(modal);
            }
            stopDoA();
            if (doaState.focusReturnElement) {
                doaState.focusReturnElement.focus();
            }
        };
        closeButton.onclick = closeWarning;
        header.appendChild(title);
        header.appendChild(closeButton);
        var messageDiv = document.createElement('p');
        messageDiv.id = 'doa-warning-message';
        messageDiv.textContent = 'The current page does not contain the Document Log Entries table. Please navigate to a page with the DoA Log before using this feature.';
        messageDiv.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0; font-size: 14px; line-height: 1.5;';
        var okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = 'background: rgba(255, 255, 255, 0.2); border: 2px solid rgba(255, 255, 255, 0.3); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.3s ease; margin-top: 20px; width: 100%;';
        okButton.onmouseover = function() {
            okButton.style.background = 'rgba(255, 255, 255, 0.3)';
        };
        okButton.onmouseout = function() {
            okButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        okButton.onclick = closeWarning;
        var keyHandler = function(e) {
            if (e.key === 'Escape') {
                closeWarning();
            }
        };
        document.addEventListener('keydown', keyHandler);
        doaState.eventListeners.push({ element: document, type: 'keydown', handler: keyHandler });
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
        addLogMessage('showDoAWarning: warning displayed', 'log');
    }

    function stopDoA() {
        addLogMessage('stopDoA: stopping all DoA processes', 'log');
        doaState.isRunning = false;
        for (var i = 0; i < doaState.idleCallbackIds.length; i++) {
            try {
                if (typeof cancelIdleCallback === 'function') {
                    cancelIdleCallback(doaState.idleCallbackIds[i]);
                }
            } catch (e) {
                addLogMessage('stopDoA: error canceling idle callback: ' + e, 'error');
            }
        }
        doaState.idleCallbackIds = [];
        if (doaState.idleCallbackId && typeof cancelIdleCallback === 'function') {
            cancelIdleCallback(doaState.idleCallbackId);
            doaState.idleCallbackId = null;
        }
        for (var i2 = 0; i2 < doaState.observers.length; i2++) {
            try {
                doaState.observers[i2].disconnect();
            } catch (e2) {
                addLogMessage('stopDoA: error disconnecting observer: ' + e2, 'error');
            }
        }
        doaState.observers = [];
        for (var i3 = 0; i3 < doaState.timeouts.length; i3++) {
            try {
                clearTimeout(doaState.timeouts[i3]);
            } catch (e3) {
                addLogMessage('stopDoA: error clearing timeout: ' + e3, 'error');
            }
        }
        doaState.timeouts = [];
        for (var i4 = 0; i4 < doaState.intervals.length; i4++) {
            try {
                clearInterval(doaState.intervals[i4]);
            } catch (e4) {
                addLogMessage('stopDoA: error clearing interval: ' + e4, 'error');
            }
        }
        doaState.intervals = [];
        for (var i5 = 0; i5 < doaState.eventListeners.length; i5++) {
            try {
                var listener = doaState.eventListeners[i5];
                listener.element.removeEventListener(listener.type, listener.handler);
            } catch (e5) {
                addLogMessage('stopDoA: error removing event listener: ' + e5, 'error');
            }
        }
        doaState.eventListeners = [];
        if (doaState.abortController) {
            doaState.abortController.abort();
            doaState.abortController = null;
        }
        doaSetAriaBusyOff();
        if (doaState.scrollContainer && doaState.prevScrollTop !== undefined) {
            addLogMessage('stopDoA: restoring viewport', 'log');
            restoreViewport(doaState.scrollContainer, doaState.prevScrollTop);
        }
        doaState.scrollContainer = null;
        doaState.userScrollHandler = null;
        doaState.userScrollPaused = false;
        if (doaState.isAddingEntries) {
            addLogMessage('stopDoA: was adding entries, marking remaining as Stopped', 'log');
            for (var qi = doaState.addQueueIndex; qi < doaState.addQueue.length; qi++) {
                if (doaState.addQueue[qi].status === DOA_LABELS.statusPending) {
                    doaState.addQueue[qi].status = DOA_LABELS.statusStopped;
                    doaState.counters.pending--;
                }
            }
            doaState.isAddingEntries = false;
        }
        doaState.addQueue = [];
        doaState.addQueueIndex = 0;
        doaState.existingPairs = new Set();
        doaState.listScrollTop = 0;
        doaState.roleListScrollTop = 0;
        var inputModal = document.getElementById('doa-input-modal');
        if (inputModal && inputModal.parentNode) {
            inputModal.parentNode.removeChild(inputModal);
        }
        var progressModal = document.getElementById('doa-progress-modal');
        if (progressModal && progressModal.parentNode) {
            progressModal.parentNode.removeChild(progressModal);
        }
        var warningModal = document.getElementById('doa-warning-modal');
        if (warningModal && warningModal.parentNode) {
            warningModal.parentNode.removeChild(warningModal);
        }
        removeCollectingDataPanel('doa');
        if (doaState.focusReturnElement) {
            doaState.focusReturnElement.focus();
        }
        resetDoAState();
        addLogMessage('stopDoA: cleanup complete', 'log');
    }

    function stopResponsibilities() {
        addLogMessage('stopResponsibilities: stopping', 'log');
        respState.stopRequested = true;
        respState.isRunning = false;
        for (var i = 0; i < respState.idleCallbackIds.length; i++) {
            try {
                if (typeof cancelIdleCallback === 'function') {
                    cancelIdleCallback(respState.idleCallbackIds[i]);
                }
            } catch (e) {
                addLogMessage('stopResponsibilities: error canceling idle callback: ' + e, 'error');
            }
        }
        respState.idleCallbackIds = [];
        for (var i2 = 0; i2 < respState.observers.length; i2++) {
            try {
                respState.observers[i2].disconnect();
            } catch (e2) {
                addLogMessage('stopResponsibilities: error disconnecting observer: ' + e2, 'error');
            }
        }
        respState.observers = [];
        for (var i3 = 0; i3 < respState.timeouts.length; i3++) {
            try {
                clearTimeout(respState.timeouts[i3]);
            } catch (e3) {
                addLogMessage('stopResponsibilities: error clearing timeout: ' + e3, 'error');
            }
        }
        respState.timeouts = [];
        for (var i4 = 0; i4 < respState.intervals.length; i4++) {
            try {
                clearInterval(respState.intervals[i4]);
            } catch (e4) {
                addLogMessage('stopResponsibilities: error clearing interval: ' + e4, 'error');
            }
        }
        respState.intervals = [];
        for (var i5 = 0; i5 < respState.eventListeners.length; i5++) {
            try {
                var l = respState.eventListeners[i5];
                l.element.removeEventListener(l.type, l.handler);
            } catch (e5) {
                addLogMessage('stopResponsibilities: error removing listener: ' + e5, 'error');
            }
        }
        respState.eventListeners = [];
        respSetAriaBusyOff();
        var im = document.getElementById('resp-input-modal');
        if (im && im.parentNode) {
            im.parentNode.removeChild(im);
        }
        var pm = document.getElementById('resp-progress-modal');
        if (pm && pm.parentNode) {
            pm.parentNode.removeChild(pm);
        }
        var wm = document.getElementById('resp-warning-modal');
        if (wm && wm.parentNode) {
            wm.parentNode.removeChild(wm);
        }
        if (respState.focusReturnElement) {
            respState.focusReturnElement.focus();
        }
        resetRespState();
        addLogMessage('stopResponsibilities: cleanup complete', 'log');
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

    let logBoxPendingUpdate = false;
    function addLogMessage(message, type = 'log') {
        const timestamp = new Date().toLocaleTimeString();
        logMessages.push({ timestamp, message, type });
        if (!logBoxPendingUpdate) {
            logBoxPendingUpdate = true;
            requestAnimationFrame(function() {
                logBoxPendingUpdate = false;
                updateLogBox();
            });
        }
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

        for (let i = 1; i <= 5; i++) {
            const button = document.createElement('button');
            if (i === 1) {
                button.textContent = 'Add Signatures';
            } else if (i === 2) {
                button.textContent = 'Add Training Log Staff Entries';
                button.id = 'elog-staff-entries-btn';
            } else if (i === 5) {
                button.textContent = 'Set Role Responsibilities';
                button.id = 'resp-set-btn';
            } else if (i === 3) {
                button.textContent = CLEAN_LABELS.featureButton;
                button.id = 'clean-resp-btn';
            } else if (i === 4) {
                button.textContent = DOA_LABELS.featureButton;
                button.id = 'doa-staff-entries-btn';
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
            } else if (i === 3) {
                button.onclick = () => {
                    console.log('Set Responsibilities button clicked');
                    setResponsibilitiesInit();
                };
            } else if (i === 4) {
                button.onclick = () => {
                    console.log('Clean Responsibility button clicked');
                    cleanResponsibilityInit();
                };
            } else if (i === 5) {
                button.onclick = () => {
                    console.log('Add DoA Log Staff Entries button clicked');
                    addDoALogStaffEntriesInit();
                };
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

        const clearLogsBtn = document.createElement('button');
        clearLogsBtn.textContent = 'Clear Logs';
        clearLogsBtn.style.cssText = `
        margin-top: 8px;
        width: 100%;
        background: rgba(255, 255, 255, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.2);
        color: rgba(255, 255, 255, 0.8);
        padding: 6px 12px;
        border-radius: 6px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 500;
        transition: all 0.3s ease;
    `;
        clearLogsBtn.onmouseover = () => {
            clearLogsBtn.style.background = 'rgba(255, 67, 54, 0.4)';
            clearLogsBtn.style.borderColor = 'rgba(255, 67, 54, 0.6)';
            clearLogsBtn.style.color = 'white';
        };
        clearLogsBtn.onmouseout = () => {
            clearLogsBtn.style.background = 'rgba(255, 255, 255, 0.1)';
            clearLogsBtn.style.borderColor = 'rgba(255, 255, 255, 0.2)';
            clearLogsBtn.style.color = 'rgba(255, 255, 255, 0.8)';
        };
        clearLogsBtn.onclick = () => {
            logMessages.length = 0;
            updateLogBox();
        };
        scaleContainer.appendChild(clearLogsBtn);

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
        var maxDisplay = 200;
        var startIdx = Math.max(0, logMessages.length - maxDisplay);
        var html = '';
        for (var i = startIdx; i < logMessages.length; i++) {
            var msg = logMessages[i];
            var color = msg.type === 'error' ? '#ff6b6b' :
                msg.type === 'warn' ? '#ffd93d' : '#6bcf7f';
            html += '<div style="color: ' + color + '; margin-bottom: 4px;"><span style="opacity: 0.7;">[' + msg.timestamp + ']</span> ' + msg.message + '</div>';
        }
        logBox.innerHTML = html;
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