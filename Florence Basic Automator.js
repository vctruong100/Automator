
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
        stepPx: 600,
        idleDelayMs: 80,
        settleDelayMs: 250,
        maxDurationMs: 120000,
        maxNoProgressIterations: 8,
        userScrollPauseMs: 800,
        viewportOverscanPx: 400,
        retryScanAttempts: 3,
        retryScanDelayMs: 200
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
        listContainer: 'ul.filtered-select__list.u-z-index-1060, ul.filtered-select__list',
        virtualViewport: '.filtered-select__list, .cdk-virtual-scroll-viewport',
        optionItem: '.filtered-select__list [role="option"], .filtered-select__list li, .cdk-virtual-scroll-viewport [role="option"], .cdk-virtual-scroll-viewport li',
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
        statusAdded: 'Added',
        statusNotInDropdown: 'Not In Dropdown',
        statusSelectionFailed: 'Selection Failed',
        statusSaveFailed: 'Save Failed',
        statusStopped: 'Stopped'
    };

    const RESP_SELECTORS = {
        pageStepRoot: 'doa-log-template-study-roles-step',
        roleColumns: '.doa-log-form-step__column.roles__column',
        roleSearchInput: '#filtered-select-input.filtered-select__input[placeholder*="Search"]',
        roleListContainer: 'ul.filtered-select__list.u-z-index-1060, ul.filtered-select__list',
        virtualViewport: 'cdk-virtual-scroll-viewport.cdk-virtual-scroll-viewport, .filtered-select__list',
        roleOptionItem: '.filtered-select__list__item, [role="option"], .cdk-virtual-scroll-viewport .filtered-select__list__item',
        roleOptionText: '.filtered-select__list__item__text',
        roleOptionCheckboxTri: '.test-checkboxTristate[role="checkbox"]',
        responsibilitiesToggleBtn: 'button[dropdowntoggle][data-test="study-responsibilities-dropdown-toggle"].roles__select-options-dropdown-button',
        responsibilitiesMenu: 'ul[role="menu"][data-test="study-responsibilities-list"]',
        responsibilitiesItem: 'ul[role="menu"][data-test="study-responsibilities-list"] li',
        responsibilitiesItemCheckbox: 'ul[role="menu"][data-test="study-responsibilities-list"] li input[type="checkbox"]',
        addStudyRoleBtn: 'button[data-test="add-study-role-button"]',
        selectedRoleCheckboxInColumn: '[role="checkbox"][aria-checked="true"]',
        mainPanelButtonTarget: '.main-gui-panel',
        ariaLiveRegion: '.aria-live-region'
    };

    const RESP_TIMEOUTS = {
        waitPageMs: 10000,
        waitInputPanelMs: 10000,
        waitProgressPanelMs: 10000,
        waitListOpenMs: 6000,
        waitOptionRenderMs: 3000,
        waitAfterSelectRoleMs: 2000,
        waitResponsibilitiesMenuMs: 4000,
        waitAfterToggleResponsibilityMs: 500,
        waitAddRoleResultMs: 8000,
        settleAfterScrollMs: 250,
        idleBetweenScrollsMs: 120,
        maxSelectRoleDurationMs: 45000
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
        quoteCleanup: /[\u201c\u201d\u201e\u201f"""]+/g,
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
        'research assistant': 'Research Assistant',
        'study coordinator': 'Study Coordinator',
        'pharmacy': 'Pharmacy'
    };

    const RESP_SCROLL = {
        stepPx: 0.7,
        userPauseMs: 800
    };

    const RESP_ATTRS = {
        ariaBusyTarget: 'body',
        ariaBusyAttr: 'aria-busy'
    };

    let respState = {
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
        closeButton.innerHTML = 'âœ•';
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
        closeButton.innerHTML = 'âœ•';
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
        closeButton.innerHTML = 'âœ•';
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
                        if (data.reason !== 'stopped' && elogState.isRunning) {
                            addLogMessage('startELogScan: scan complete, starting add entries flow', 'log');
                            beginAddNonDuplicateELogEntries();
                        }
                    },
                    onError: function(error) {
                        addLogMessage('startELogScan: error - ' + error.message, 'error');
                        updateScanStatus('Error', 'error');
                        showInlineNotice('Error during auto-scroll scan: ' + error.message);
                    }
                });
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
        addLogMessage('performRescan: restarting scan with auto-scroll', 'log');
        updateAriaLiveRegion(ELOG_LABELS.progressRescanning);
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
        const leftPanel = document.getElementById('elog-left-panel');
        if (leftPanel) {
            leftPanel.innerHTML = '';
        }
        elogState.scannedNames = [];
        elogState.seenNormalizedNames = new Set();
        elogState.leftPanelRowIndex = 0;
        elogState.userScrollPaused = false;
        updateScanStatus(ELOG_LABELS.progressRescanning, 'progress');
        const title = document.getElementById('elog-progress-title');
        if (title) {
            title.textContent = 'ELog Staff Entries - ' + ELOG_LABELS.progressRescanning;
        }
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
                if (data.reason !== 'stopped' && elogState.isRunning) {
                    addLogMessage('performRescan: re-scan complete, starting add entries flow', 'log');
                    beginAddNonDuplicateELogEntries();
                }
            },
            onError: function(error) {
                addLogMessage('performRescan: error - ' + error.message, 'error');
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
        addLogMessage('getFirstAndLast: parsing ' + name, 'log');
        if (!name || !name.trim()) {
            addLogMessage('getFirstAndLast: empty name', 'warn');
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
            addLogMessage('getFirstAndLast: no tokens', 'warn');
            return ['', ''];
        }
        if (filtered.length === 1) {
            addLogMessage('getFirstAndLast: single token=' + filtered[0], 'log');
            return [filtered[0], filtered[0]];
        }
        const first = filtered[0];
        let last = filtered[filtered.length - 1];
        if (filtered.length > 2 && suffixPattern.test(last)) {
            addLogMessage('getFirstAndLast: stripping suffix=' + last, 'log');
            last = filtered[filtered.length - 2];
        }
        addLogMessage('getFirstAndLast: first=' + first + ' last=' + last, 'log');
        return [first, last];
    }

    function normalizeFirstLastPair(name) {
        addLogMessage('normalizeFirstLastPair: name=' + name, 'log');
        const pair = getFirstAndLast(name);
        const result = elogNormalizeName(pair[0] + ' ' + pair[1]);
        addLogMessage('normalizeFirstLastPair: result=' + result, 'log');
        return result;
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
            const normalized = item.getAttribute('data-normalized');
            if (!normalized) {
                continue;
            }
            const itemPairKey = normalizeFirstLastPair(item.querySelector('span') ? item.querySelector('span:not(.elog-status-badge)').textContent : '');
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

    function attemptSelectByScrollingForName(targetDisplay, targetPairKey) {
        addLogMessage('attemptSelectByScrollingForName: target=' + targetDisplay + ' pairKey=' + targetPairKey, 'log');
        return new Promise(function(resolve) {
            var viewportEl = findDropdownViewportElement();
            if (!viewportEl) {
                addLogMessage('attemptSelectByScrollingForName: no viewport element', 'error');
                resolve(false);
                return;
            }
            restoreListScrollPosition(viewportEl);
            var passCount = 0;
            var lastScrollTop = -1;
            var lastOptionSnapshot = '';
            var startTime = Date.now();

            function scanAndScroll() {
                if (!elogState.isRunning) {
                    addLogMessage('attemptSelectByScrollingForName: stopped', 'warn');
                    rememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                if (Date.now() - startTime > ELOG_FORM_TIMEOUTS.maxSelectDurationMs) {
                    addLogMessage('attemptSelectByScrollingForName: max duration exceeded', 'warn');
                    rememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                if (passCount >= ELOG_FORM_RETRY.maxScrollPasses) {
                    addLogMessage('attemptSelectByScrollingForName: max passes reached', 'warn');
                    rememberListScrollPosition(viewportEl);
                    resolve(false);
                    return;
                }
                var options = document.querySelectorAll(ELOG_FORM_SELECTORS.optionItem);
                var retryCount = 0;
                function tryScanOptions() {
                    options = document.querySelectorAll(ELOG_FORM_SELECTORS.optionItem);
                    if (options.length === 0 && retryCount < ELOG_FORM_RETRY.selectRetriesPerScroll) {
                        retryCount++;
                        addLogMessage('attemptSelectByScrollingForName: empty render, retry ' + retryCount, 'log');
                        var rtid = setTimeout(tryScanOptions, ELOG_FORM_TIMEOUTS.waitOptionRenderMs);
                        elogState.timeouts.push(rtid);
                        return;
                    }
                    var found = false;
                    var optionTexts = [];
                    for (var oi = 0; oi < options.length; oi++) {
                        var optText = (options[oi].textContent || '').trim();
                        optionTexts.push(optText);
                        var optPairKey = normalizeFirstLastPair(optText);
                        if (optPairKey === targetPairKey) {
                            addLogMessage('attemptSelectByScrollingForName: match found at index ' + oi + ' text=' + optText, 'log');
                            options[oi].click();
                            found = true;
                            rememberListScrollPosition(viewportEl);
                            var verifyTid = setTimeout(function() {
                                var inputEl = document.querySelector(ELOG_FORM_SELECTORS.memberInput);
                                if (inputEl && inputEl.value && inputEl.value.trim()) {
                                    addLogMessage('attemptSelectByScrollingForName: selection verified via input value', 'log');
                                    resolve(true);
                                } else {
                                    var listGone = !document.querySelector(ELOG_FORM_SELECTORS.listContainer);
                                    if (listGone) {
                                        addLogMessage('attemptSelectByScrollingForName: selection verified via list close', 'log');
                                        resolve(true);
                                    } else {
                                        addLogMessage('attemptSelectByScrollingForName: selection not verified, treating as success', 'log');
                                        resolve(true);
                                    }
                                }
                            }, ELOG_FORM_TIMEOUTS.settleMs);
                            elogState.timeouts.push(verifyTid);
                            break;
                        }
                    }
                    if (!found) {
                        var currentSnapshot = optionTexts.join('|');
                        var currentTop = viewportEl.scrollTop;
                        var noProgress = (currentTop === lastScrollTop && currentSnapshot === lastOptionSnapshot);
                        if (noProgress) {
                            passCount++;
                            addLogMessage('attemptSelectByScrollingForName: no progress pass ' + passCount, 'log');
                        }
                        lastScrollTop = currentTop;
                        lastOptionSnapshot = currentSnapshot;
                        var stepSize = Math.round(viewportEl.clientHeight * 0.7);
                        var maxScroll = viewportEl.scrollHeight - viewportEl.clientHeight;
                        var newTop = Math.min(currentTop + stepSize, maxScroll);
                        if (newTop <= currentTop && currentTop > 0) {
                            addLogMessage('attemptSelectByScrollingForName: at bottom, trying reverse', 'log');
                            newTop = 0;
                            lastScrollTop = -1;
                            lastOptionSnapshot = '';
                            passCount++;
                        }
                        viewportEl.scrollTop = newTop;
                        var scrollTid = setTimeout(scanAndScroll, ELOG_FORM_TIMEOUTS.scrollIdleMs);
                        elogState.timeouts.push(scrollTid);
                    }
                }
                tryScanOptions();
            }
            var initTid = setTimeout(scanAndScroll, ELOG_FORM_TIMEOUTS.scrollIdleMs);
            elogState.timeouts.push(initTid);
        });
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

    function findScrollableContainer(gridEl) {
        addLogMessage('findScrollableContainer: searching for scrollable ancestor', 'log');
        if (!gridEl) {
            addLogMessage('findScrollableContainer: gridEl is null', 'warn');
            return null;
        }
        let current = gridEl;
        while (current && current !== document.body) {
            const style = window.getComputedStyle(current);
            const overflowY = style.overflowY;
            const scrollDiff = current.scrollHeight - current.clientHeight;
            addLogMessage('findScrollableContainer: checking ' + (current.className || current.tagName) + ' scrollDiff=' + scrollDiff + ' overflowY=' + overflowY, 'log');
            if (scrollDiff >= 20 && (overflowY === 'scroll' || overflowY === 'auto')) {
                addLogMessage('findScrollableContainer: found scrollable container', 'log');
                return current;
            }
            current = current.parentElement;
        }
        addLogMessage('findScrollableContainer: no scrollable ancestor, using gridEl', 'warn');
        return gridEl;
    }

    function getRenderedRowCount() {
        addLogMessage('getRenderedRowCount: counting rows', 'log');
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
        addLogMessage('getRenderedRowCount: count=' + count, 'log');
        return count;
    }

    function getRenderedLastRowKey() {
        addLogMessage('getRenderedLastRowKey: computing key', 'log');
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
                addLogMessage('getRenderedLastRowKey: key=' + firstCellText, 'log');
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
        addLogMessage('getRenderedLastRowKey: hash key=' + key, 'log');
        return key;
    }

    function awaitSettle(container) {
        addLogMessage('awaitSettle: waiting for settle', 'log');
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
                    addLogMessage('awaitSettle: timeout after ' + ELOG_SCROLL.settleDelayMs + 'ms', 'log');
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
                        addLogMessage('awaitSettle: mutation detected', 'log');
                        resolve();
                    }
                });
                observer.observe(gridTable, { childList: true, subtree: true });
                elogState.observers.push(observer);
            }
        });
    }

    function setAriaBusyOn() {
        addLogMessage('setAriaBusyOn: setting aria-busy', 'log');
        const target = document.querySelector(ELOG_ATTRS.ariaBusyTarget);
        if (target) {
            elogState.prevAriaBusy = target.getAttribute(ELOG_ATTRS.ariaBusyAttr);
            target.setAttribute(ELOG_ATTRS.ariaBusyAttr, 'true');
            addLogMessage('setAriaBusyOn: prev=' + elogState.prevAriaBusy, 'log');
        }
    }

    function setAriaBusyOff() {
        addLogMessage('setAriaBusyOff: restoring aria-busy', 'log');
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
        addLogMessage('computeEndReached: checking end', 'log');
        if (!container) {
            return true;
        }
        const atBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 1;
        addLogMessage('computeEndReached: atBottom=' + atBottom + ' noProgress=' + noProgress, 'log');
        if (atBottom && noProgress >= 1) {
            return true;
        }
        return false;
    }

    function restoreViewport(container, prevTop) {
        addLogMessage('restoreViewport: prevTop=' + prevTop, 'log');
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
                addLogMessage('observeUserScrollPause: user scroll detected, pausing', 'log');
                elogState.userScrollPaused = true;
                const resumeTimeout = setTimeout(function() {
                    addLogMessage('observeUserScrollPause: resuming', 'log');
                    elogState.userScrollPaused = false;
                }, ELOG_SCROLL.userScrollPauseMs);
                elogState.timeouts.push(resumeTimeout);
            }
        };
        container.addEventListener('scroll', elogState.userScrollHandler);
        elogState.eventListeners.push({ element: container, type: 'scroll', handler: elogState.userScrollHandler });
    }

    function scanVisibleRowsOnce(onRow) {
        addLogMessage('scanVisibleRowsOnce: scanning', 'log');
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
            addLogMessage('scanVisibleRowsOnce: new name: ' + extractedName, 'log');
            if (onRow) {
                onRow(extractedName, normalized);
            }
        }
        addLogMessage('scanVisibleRowsOnce: newCount=' + newCount, 'log');
        return newCount;
    }

    function updateAriaLiveRegion(message) {
        addLogMessage('updateAriaLiveRegion: ' + message, 'log');
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
        addLogMessage('autoScrollScan: initial scan', 'log');
        scanVisibleRowsOnce(function(name) {
            elogState.leftPanelRowIndex++;
            updateLeftPanelList(name, elogState.leftPanelRowIndex);
            checkAndUpdateRightPanel(name);
        });
        onProgress({ scanned: elogState.scannedNames.length });
        updateAriaLiveRegion('Scanning: ' + elogState.scannedNames.length + ' names');
        function scrollLoop() {
            addLogMessage('autoScrollScan: loop iteration', 'log');
            if (!elogState.isRunning) {
                addLogMessage('autoScrollScan: stopped', 'warn');
                updateAriaLiveRegion(ELOG_LABELS.progressStopped);
                finishScan(ELOG_LABELS.progressStopped, 'stopped');
                return;
            }
            if (Date.now() - startTime > ELOG_SCROLL.maxDurationMs) {
                addLogMessage('autoScrollScan: timeout', 'warn');
                finishScan(ELOG_LABELS.progressComplete, 'timeout');
                return;
            }
            if (elogState.userScrollPaused) {
                addLogMessage('autoScrollScan: paused by user', 'log');
                const pt = setTimeout(scrollLoop, 100);
                elogState.timeouts.push(pt);
                return;
            }
            const priorKey = getRenderedLastRowKey();
            const priorCount = getRenderedRowCount();
            const currTop = container.scrollTop;
            const maxScroll = container.scrollHeight - container.clientHeight;
            const newTop = Math.min(currTop + ELOG_SCROLL.stepPx, maxScroll);
            addLogMessage('autoScrollScan: scroll ' + currTop + ' -> ' + newTop, 'log');
            elogState.lastAutoScrollTime = Date.now();
            container.scrollTo({ top: newTop, behavior: 'auto' });
            awaitSettle(container).then(function() {
                let attempts = 0;
                function attemptScan() {
                    const rc = getRenderedRowCount();
                    if (rc === 0 && attempts < ELOG_SCROLL.retryScanAttempts) {
                        attempts++;
                        addLogMessage('autoScrollScan: zero rows, retry ' + attempts, 'log');
                        const rt = setTimeout(attemptScan, ELOG_SCROLL.retryScanDelayMs);
                        elogState.timeouts.push(rt);
                        return;
                    }
                    scanVisibleRowsOnce(function(name) {
                        elogState.leftPanelRowIndex++;
                        updateLeftPanelList(name, elogState.leftPanelRowIndex);
                        checkAndUpdateRightPanel(name);
                    });
                    onProgress({ scanned: elogState.scannedNames.length });
                    updateAriaLiveRegion('Scanning: ' + elogState.scannedNames.length + ' names');
                    const currKey = getRenderedLastRowKey();
                    const currCount = getRenderedRowCount();
                    if (currKey === priorKey && currCount === priorCount) {
                        noProgress++;
                        addLogMessage('autoScrollScan: no progress ' + noProgress, 'log');
                    } else {
                        noProgress = 0;
                    }
                    if (computeEndReached(container, noProgress)) {
                        addLogMessage('autoScrollScan: end reached', 'log');
                        finishScan(ELOG_LABELS.progressNoMore, 'endReached');
                        return;
                    }
                    if (noProgress >= ELOG_SCROLL.maxNoProgressIterations) {
                        addLogMessage('autoScrollScan: max no-progress', 'log');
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
            addLogMessage('autoScrollScan: finish reason=' + reason, 'log');
            for (let i = 0; i < elogState.parsedNames.length; i++) {
                const nameObj = elogState.parsedNames[i];
                if (nameObj.status === 'Pending') {
                    nameObj.status = 'Not Found';
                    updateRightPanelItemStatus(nameObj.normalized, 'Not Found', 'notfound');
                }
            }
            if (reason === 'stopped') {
                updateScanStatus(ELOG_LABELS.progressStopped, 'stopped');
            } else {
                updateScanStatus(ELOG_LABELS.progressComplete, 'complete');
            }
            const title = document.getElementById('elog-progress-title');
            if (title) {
                title.textContent = 'ELog Staff Entries - ' + label;
            }
            setAriaBusyOff();
            addLogMessage('autoScrollScan: total=' + elogState.scannedNames.length, 'log');
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
        resetELogState();
        addLogMessage('stopELog: cleanup complete', 'log');
    }

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
                    if (idx > -1) { respState.observers.splice(idx, 1); }
                    resolve(el);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            respState.observers.push(observer);
            var timeoutId = setTimeout(function() {
                observer.disconnect();
                var idx = respState.observers.indexOf(observer);
                if (idx > -1) { respState.observers.splice(idx, 1); }
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
        if (!s) { return { display: '', key: '' }; }
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
                    for (var n = lo; n <= hi; n++) { result.add(n); }
                    i += 3;
                    continue;
                }
            }
            var num = parseInt(current, 10);
            if (!isNaN(num)) { result.add(num); }
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
                        if (buffer[ci] === '"') { quoteCount++; }
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
                    if (line[ci2] === '"') { qc++; }
                }
                if (qc % 2 !== 0) {
                    pendingQuote = true;
                    buffer = line;
                    continue;
                }
                mergedLines.push(line);
            }
            if (buffer) { mergedLines.push(buffer); }
            addLogMessage('parseResponsibilitiesInput: merged into ' + mergedLines.length + ' lines', 'log');
            var parsedMap = {};
            for (var mi = 0; mi < mergedLines.length; mi++) {
                var mline = mergedLines[mi].trim();
                if (!mline) { continue; }
                mline = mline.replace(/"+/g, '');
                var parts = mline.split(/\t+/);
                var rolePart = '';
                var numberPart = '';
                if (parts.length >= 2) {
                    rolePart = parts[0].trim();
                    numberPart = parts.slice(1).join(' ');
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
                var numTokens = numberPart.replace(/,/g, ' ').split(/\s+/).filter(function(t) { return t.length > 0; });
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
                occ.forEach(function(n) { union.add(n); });
                if (intersection === null) {
                    intersection = new Set(occ);
                } else {
                    var newInt = new Set();
                    intersection.forEach(function(n) { if (occ.has(n)) { newInt.add(n); } });
                    intersection = newInt;
                }
            }
            if (intersection === null) { intersection = new Set(); }
            var excluded = new Set();
            union.forEach(function(n) { if (!intersection.has(n)) { excluded.add(n); } });
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
                common: Array.from(intersection).sort(function(a, b) { return a - b; }),
                excluded: Array.from(excluded).sort(function(a, b) { return a - b; }),
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
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.3)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        var closeWarning = function() {
            addLogMessage('showRespWarning: closing warning', 'log');
            if (modal.parentNode) { document.body.removeChild(modal); }
            if (respState.focusReturnElement) { respState.focusReturnElement.focus(); }
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
        okButton.onmouseover = function() { okButton.style.background = 'rgba(255, 255, 255, 0.3)'; };
        okButton.onmouseout = function() { okButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        okButton.onclick = closeWarning;
        var keyHandler = function(e) { if (e.key === 'Escape') { closeWarning(); } };
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
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 67, 54, 0.8)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        closeButton.onclick = function() {
            addLogMessage('showResponsibilitiesInputPanel: closed by user', 'warn');
            if (modal.parentNode) { document.body.removeChild(modal); }
            stopResponsibilities();
        };
        header.appendChild(titleEl);
        header.appendChild(closeButton);
        var description = document.createElement('p');
        description.textContent = 'Paste role-to-responsibility assignments below. Each line: Role name followed by tab and responsibility numbers. Ranges like "1 to 8" are supported.';
        description.style.cssText = 'color: rgba(255, 255, 255, 0.9); margin: 0 0 12px 0; font-size: 14px; line-height: 1.4;';
        var textarea = document.createElement('textarea');
        textarea.id = 'resp-input-textarea';
        textarea.placeholder = 'PI  1 to 8  13  14  17  21\nStudy Coordinator  1 6 7 8 10 12 13 14 17 33';
        textarea.setAttribute('aria-label', 'Role responsibilities input');
        textarea.style.cssText = 'width: 100%; height: 200px; padding: 12px 14px; border: 2px solid rgba(255, 255, 255, 0.35); border-radius: 10px; background: rgba(255, 255, 255, 0.95); color: #1e293b; font-size: 14px; font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; resize: vertical; outline: none; transition: all 0.25s ease; box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset; box-sizing: border-box;';
        textarea.onfocus = function() { textarea.style.borderColor = '#8ea0ff'; textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)'; };
        textarea.onblur = function() { textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)'; textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset'; };
        var confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.disabled = true;
        confirmButton.style.cssText = 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; letter-spacing: 0.2px; transition: all 0.25s ease; opacity: 0.5;';
        textarea.oninput = function() {
            if (textarea.value.trim().length > 0) { confirmButton.disabled = false; confirmButton.style.opacity = '1'; confirmButton.style.cursor = 'pointer'; }
            else { confirmButton.disabled = true; confirmButton.style.opacity = '0.5'; confirmButton.style.cursor = 'not-allowed'; }
        };
        confirmButton.onmouseover = function() { if (!confirmButton.disabled) { confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)'; } };
        confirmButton.onmouseout = function() { confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)'; };
        confirmButton.onclick = function() {
            addLogMessage('showResponsibilitiesInputPanel: Confirm clicked', 'log');
            var parsedMap = parseResponsibilitiesInput(textarea.value);
            if (!parsedMap || Object.keys(parsedMap).length === 0) { addLogMessage('showResponsibilitiesInputPanel: no valid roles', 'warn'); return; }
            respState.parsedRoles = parsedMap;
            var rd = computeRoleCommonAndExcluded(parsedMap);
            respState.rolesData = rd;
            addLogMessage('showResponsibilitiesInputPanel: parsed ' + rd.length + ' roles', 'log');
            if (modal.parentNode) { document.body.removeChild(modal); }
            showResponsibilitiesProgressPanel(rd);
            processRolesWorkflow(rd);
        };
        var clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = 'background: rgba(255, 255, 255, 0.18); border: 2px solid rgba(255, 255, 255, 0.35); color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.25s ease;';
        clearButton.onmouseover = function() { clearButton.style.background = 'rgba(255, 255, 255, 0.28)'; };
        clearButton.onmouseout = function() { clearButton.style.background = 'rgba(255, 255, 255, 0.18)'; };
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
        if (status === RESP_LABELS.statusCompleted) { return { color: '#6bcf7f', bg: 'rgba(107, 207, 127, 0.2)' }; }
        if (status === RESP_LABELS.statusFailed) { return { color: '#ff6b6b', bg: 'rgba(255, 107, 107, 0.2)' }; }
        if (status === RESP_LABELS.statusStopped) { return { color: '#aaa', bg: 'rgba(170, 170, 170, 0.2)' }; }
        return { color: '#ffd93d', bg: 'rgba(255, 217, 61, 0.2)' };
    }

    function createRespRoleRow(roleData, index) {
        var item = document.createElement('div');
        item.className = 'resp-role-item';
        item.setAttribute('data-role-key', roleData.key);
        item.style.cssText = 'display: flex; flex-direction: column; padding: 10px 12px; margin: 4px 0; background: rgba(255, 255, 255, 0.08); border-radius: 6px; transition: background 0.2s ease;';
        item.onmouseover = function() { item.style.background = 'rgba(255, 255, 255, 0.12)'; };
        item.onmouseout = function() { item.style.background = 'rgba(255, 255, 255, 0.08)'; };
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
            if (rolesData[ci].status === RESP_LABELS.statusPending) { respState.counters.pending++; }
            else if (rolesData[ci].status === RESP_LABELS.statusFailed) { respState.counters.failed++; }
            else if (rolesData[ci].status === RESP_LABELS.statusCompleted) { respState.counters.completed++; }
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
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0;';
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
        closeButton.onmouseover = function() { closeButton.style.background = 'rgba(255, 67, 54, 0.8)'; };
        closeButton.onmouseout = function() { closeButton.style.background = 'rgba(255, 255, 255, 0.2)'; };
        closeButton.onclick = function() { addLogMessage('showResponsibilitiesProgressPanel: closed', 'warn'); stopResponsibilities(); };
        header.appendChild(titleContainer);
        header.appendChild(closeButton);
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
                items[si].style.display = (!term || text.indexOf(term) !== -1) ? 'flex' : 'none';
            }
        };
        var listContainer = document.createElement('div');
        listContainer.id = 'resp-roles-list';
        listContainer.style.cssText = 'flex: 1; overflow-y: auto; min-height: 150px; max-height: 400px;';
        for (var ri = 0; ri < rolesData.length; ri++) { listContainer.appendChild(createRespRoleRow(rolesData[ri], ri)); }
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
        var ariaLiveRegion = document.createElement('div');
        ariaLiveRegion.id = 'resp-aria-live';
        ariaLiveRegion.setAttribute('aria-live', 'polite');
        ariaLiveRegion.setAttribute('aria-atomic', 'true');
        ariaLiveRegion.style.cssText = 'position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;';
        container.appendChild(header);
        container.appendChild(searchInput);
        container.appendChild(listContainer);
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
        closeButton.focus();
        addLogMessage('showResponsibilitiesProgressPanel: displayed', 'log');
    }

    function updateRespRoleStatus(roleKey, newStatus, reason) {
        addLogMessage('updateRespRoleStatus: role=' + roleKey + ' status=' + newStatus, 'log');
        var items = document.querySelectorAll('.resp-role-item');
        for (var i = 0; i < items.length; i++) {
            if (items[i].getAttribute('data-role-key') === roleKey) {
                var badge = items[i].querySelector('.resp-status-badge');
                if (badge) { badge.textContent = newStatus; var colors = getRespBadgeColors(newStatus); badge.style.color = colors.color; badge.style.background = colors.bg; }
                if (reason) {
                    var dr = items[i].querySelector('.resp-detail-row');
                    if (dr) { var re = document.createElement('div'); re.style.cssText = 'color: #ff6b6b; margin-top: 2px;'; re.textContent = reason; dr.appendChild(re); }
                }
                var lr = items[i].querySelector('[aria-live]');
                if (lr) { lr.textContent = roleKey + ' ' + newStatus; }
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
        if (el1) { el1.textContent = String(respState.counters.total); }
        if (el2) { el2.textContent = String(respState.counters.completed); }
        if (el3) { el3.textContent = String(respState.counters.failed); }
        if (el4) { el4.textContent = String(respState.counters.pending); }
        if (el5) {
            var processed = respState.counters.completed + respState.counters.failed;
            var pct = respState.counters.total > 0 ? Math.round((processed / respState.counters.total) * 100) : 0;
            el5.textContent = pct + '%';
        }
        updateRespAriaLive('Completed: ' + respState.counters.completed + ', Failed: ' + respState.counters.failed + ', Pending: ' + respState.counters.pending);
    }

    function updateRespAriaLive(message) {
        var lr = document.getElementById('resp-aria-live');
        if (lr) { lr.textContent = message; }
    }

    function scanExistingCompletedRoles() {
        addLogMessage('scanExistingCompletedRoles: scanning', 'log');
        try {
            var columns = document.querySelectorAll(RESP_SELECTORS.roleColumns);
            addLogMessage('scanExistingCompletedRoles: ' + columns.length + ' columns', 'log');
            for (var colIdx = 0; colIdx < columns.length; colIdx++) {
                var checked = columns[colIdx].querySelectorAll(RESP_SELECTORS.selectedRoleCheckboxInColumn);
                for (var bi = 0; bi < checked.length; bi++) {
                    var ariaLabel = checked[bi].getAttribute('aria-label') || '';
                    var parentItem = checked[bi].closest('.filtered-select__list__item');
                    var roleText = ariaLabel || (parentItem ? parentItem.textContent || '' : '');
                    if (!roleText) { continue; }
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

    function ensureRoleListOpenForColumn() {
        addLogMessage('ensureRoleListOpenForColumn: ensuring open', 'log');
        return new Promise(function(resolve, reject) {
            var retries = 0;
            function tryOpen() {
                if (respState.stopRequested) { reject(new Error('Stopped')); return; }
                addLogMessage('ensureRoleListOpenForColumn: attempt ' + (retries + 1), 'log');
                var listEl = document.querySelector(RESP_SELECTORS.roleListContainer);
                if (listEl) { resolve(listEl); return; }
                var inputEl = document.querySelector(RESP_SELECTORS.roleSearchInput);
                if (inputEl) { inputEl.click(); inputEl.focus(); }
                respWaitForElement(RESP_SELECTORS.roleListContainer, RESP_TIMEOUTS.waitListOpenMs)
                    .then(function(el) { resolve(el); })
                    .catch(function(err) {
                        retries++;
                        if (retries < RESP_RETRY.openListRetries) {
                            var tid = setTimeout(tryOpen, 300);
                            respState.timeouts.push(tid);
                        } else { reject(err); }
                    });
            }
            tryOpen();
        });
    }

    function findRoleListViewport() {
        var el = document.querySelector(RESP_SELECTORS.virtualViewport);
        if (el) { addLogMessage('findRoleListViewport: found virtual', 'log'); }
        else {
            el = document.querySelector(RESP_SELECTORS.roleListContainer);
            if (el) { addLogMessage('findRoleListViewport: fallback', 'log'); }
            else { addLogMessage('findRoleListViewport: not found', 'warn'); }
        }
        return el;
    }

    function respRememberListScrollPosition(viewportEl) {
        if (viewportEl) { respState.listScrollTop = viewportEl.scrollTop; addLogMessage('respRememberScroll: ' + respState.listScrollTop, 'log'); }
    }

    function respRestoreListScrollPosition(viewportEl) {
        if (viewportEl && respState.listScrollTop > 0) { viewportEl.scrollTop = respState.listScrollTop; addLogMessage('respRestoreScroll: ' + respState.listScrollTop, 'log'); }
    }

    function attemptSelectByScrollingForRole(targetRoleDisplay, targetRoleKey) {
        addLogMessage('attemptSelectByScrollingForRole: target=' + targetRoleDisplay, 'log');
        return new Promise(function(resolve) {
            var viewportEl = findRoleListViewport();
            if (!viewportEl) { resolve(false); return; }
            respRestoreListScrollPosition(viewportEl);
            var passCount = 0;
            var lastScrollTop = -1;
            var lastSnapshot = '';
            var startTime = Date.now();
            function scanAndScroll() {
                if (respState.stopRequested) { respRememberListScrollPosition(viewportEl); resolve(false); return; }
                if (Date.now() - startTime > RESP_TIMEOUTS.maxSelectRoleDurationMs) { respRememberListScrollPosition(viewportEl); resolve(false); return; }
                if (passCount >= RESP_RETRY.maxScrollPasses) { respRememberListScrollPosition(viewportEl); resolve(false); return; }
                var options = document.querySelectorAll(RESP_SELECTORS.roleOptionItem);
                var retryCount = 0;
                function tryScan() {
                    options = document.querySelectorAll(RESP_SELECTORS.roleOptionItem);
                    if (options.length === 0 && retryCount < RESP_RETRY.optionScanRetries) {
                        retryCount++;
                        var rtid = setTimeout(tryScan, RESP_TIMEOUTS.waitOptionRenderMs);
                        respState.timeouts.push(rtid);
                        return;
                    }
                    var found = false;
                    var optTexts = [];
                    for (var oi = 0; oi < options.length; oi++) {
                        var textEl = options[oi].querySelector(RESP_SELECTORS.roleOptionText);
                        var optText = textEl ? textEl.textContent.trim() : (options[oi].textContent || '').trim();
                        optTexts.push(optText);
                        var norm = normalizeRoleName(optText);
                        if (norm.key === targetRoleKey) {
                            addLogMessage('attemptSelectByScrollingForRole: match at ' + oi, 'log');
                            var cb = options[oi].querySelector(RESP_SELECTORS.roleOptionCheckboxTri);
                            if (cb) { cb.click(); } else { options[oi].click(); }
                            found = true;
                            respRememberListScrollPosition(viewportEl);
                            var vt = setTimeout(function() { resolve(true); }, RESP_TIMEOUTS.waitAfterSelectRoleMs);
                            respState.timeouts.push(vt);
                            break;
                        }
                    }
                    if (!found) {
                        var snap = optTexts.join('|');
                        var curTop = viewportEl.scrollTop;
                        if (curTop === lastScrollTop && snap === lastSnapshot) { passCount++; }
                        lastScrollTop = curTop;
                        lastSnapshot = snap;
                        var step = Math.round(viewportEl.clientHeight * RESP_SCROLL.stepPx);
                        var maxS = viewportEl.scrollHeight - viewportEl.clientHeight;
                        var newTop = Math.min(curTop + step, maxS);
                        if (newTop <= curTop && curTop > 0) { newTop = 0; lastScrollTop = -1; lastSnapshot = ''; passCount++; }
                        viewportEl.scrollTop = newTop;
                        var st = setTimeout(scanAndScroll, RESP_TIMEOUTS.idleBetweenScrollsMs);
                        respState.timeouts.push(st);
                    }
                }
                tryScan();
            }
            var initTid = setTimeout(scanAndScroll, RESP_TIMEOUTS.idleBetweenScrollsMs);
            respState.timeouts.push(initTid);
        });
    }

    function openResponsibilitiesDropdown() {
        addLogMessage('openResponsibilitiesDropdown: opening', 'log');
        return new Promise(function(resolve, reject) {
            try {
                var toggleBtn = document.querySelector(RESP_SELECTORS.responsibilitiesToggleBtn);
                if (!toggleBtn) { reject(new Error('Toggle button not found')); return; }
                var existing = document.querySelector(RESP_SELECTORS.responsibilitiesMenu);
                if (existing) { resolve(existing); return; }
                toggleBtn.click();
                respWaitForElement(RESP_SELECTORS.responsibilitiesMenu, RESP_TIMEOUTS.waitResponsibilitiesMenuMs)
                    .then(function(menu) { resolve(menu); })
                    .catch(function(err) { reject(err); });
            } catch (error) { reject(error); }
        });
    }

    function selectResponsibilitiesByNumbers(numbersSet) {
        addLogMessage('selectResponsibilitiesByNumbers: selecting ' + numbersSet.size + ' numbers', 'log');
        return new Promise(function(resolve) {
            try {
                var items = document.querySelectorAll(RESP_SELECTORS.responsibilitiesItem);
                var itemIndex = 0;
                function processNext() {
                    if (respState.stopRequested) { resolve(); return; }
                    if (itemIndex >= items.length) { resolve(); return; }
                    var item = items[itemIndex];
                    var text = (item.textContent || '').trim();
                    var numMatch = text.match(/^(\d+)/);
                    if (numMatch) {
                        var num = parseInt(numMatch[1], 10);
                        if (numbersSet.has(num)) {
                            var checkbox = item.querySelector('input[type="checkbox"]');
                            if (checkbox && !checkbox.checked) {
                                checkbox.click();
                                itemIndex++;
                                var tid = setTimeout(processNext, RESP_TIMEOUTS.waitAfterToggleResponsibilityMs);
                                respState.timeouts.push(tid);
                                return;
                            } else if (!checkbox) {
                                item.click();
                                itemIndex++;
                                var tid2 = setTimeout(processNext, RESP_TIMEOUTS.waitAfterToggleResponsibilityMs);
                                respState.timeouts.push(tid2);
                                return;
                            }
                        }
                    }
                    itemIndex++;
                    var tid3 = setTimeout(processNext, 20);
                    respState.timeouts.push(tid3);
                }
                processNext();
            } catch (error) { addLogMessage('selectResponsibilitiesByNumbers: error: ' + error.message, 'error'); resolve(); }
        });
    }

    function clickAddStudyRole() {
        addLogMessage('clickAddStudyRole: clicking', 'log');
        return new Promise(function(resolve) {
            var retries = 0;
            function tryClick() {
                if (respState.stopRequested) { resolve(false); return; }
                var addBtn = document.querySelector(RESP_SELECTORS.addStudyRoleBtn);
                if (!addBtn) { resolve(false); return; }
                if (addBtn.disabled || addBtn.getAttribute('disabled') !== null) {
                    if (retries < RESP_RETRY.addRoleRetries) {
                        retries++;
                        var tid = setTimeout(tryClick, 1000);
                        respState.timeouts.push(tid);
                        return;
                    }
                    resolve(false);
                    return;
                }
                addBtn.click();
                var wt = setTimeout(function() { resolve(true); }, RESP_TIMEOUTS.waitAddRoleResultMs);
                respState.timeouts.push(wt);
            }
            tryClick();
        });
    }

    function respSetAriaBusyOn() {
        var target = document.querySelector(RESP_ATTRS.ariaBusyTarget);
        if (target) { respState.prevAriaBusy = target.getAttribute(RESP_ATTRS.ariaBusyAttr); target.setAttribute(RESP_ATTRS.ariaBusyAttr, 'true'); }
    }

    function respSetAriaBusyOff() {
        var target = document.querySelector(RESP_ATTRS.ariaBusyTarget);
        if (target) {
            if (respState.prevAriaBusy !== null) { target.setAttribute(RESP_ATTRS.ariaBusyAttr, respState.prevAriaBusy); }
            else { target.removeAttribute(RESP_ATTRS.ariaBusyAttr); }
            respState.prevAriaBusy = null;
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
                for (var si = roleIndex; si < rolesData.length; si++) {
                    if (rolesData[si].status === RESP_LABELS.statusPending) { rolesData[si].status = RESP_LABELS.statusStopped; respState.counters.pending--; updateRespRoleStatus(rolesData[si].key, RESP_LABELS.statusStopped, ''); }
                }
                updateRespSummary();
                respSetAriaBusyOff();
                var t1 = document.getElementById('resp-progress-title');
                if (t1) { t1.textContent = 'Set Responsibilities - Stopped'; }
                var b1 = document.getElementById('resp-status-badge');
                if (b1) { b1.textContent = 'Stopped'; b1.style.color = '#aaa'; }
                updateRespAriaLive('Process stopped');
                respState.isRunning = false;
                return;
            }
            if (roleIndex >= rolesData.length) {
                respSetAriaBusyOff();
                respState.isRunning = false;
                var t2 = document.getElementById('resp-progress-title');
                if (t2) { t2.textContent = 'Set Responsibilities - Complete'; }
                var b2 = document.getElementById('resp-status-badge');
                if (b2) { b2.textContent = 'Complete'; b2.style.background = 'rgba(107, 207, 127, 0.3)'; b2.style.color = '#6bcf7f'; }
                updateRespSummary();
                updateRespAriaLive('Complete. Done: ' + respState.counters.completed + ', Failed: ' + respState.counters.failed);
                return;
            }
            var role = rolesData[roleIndex];
            if (role.status !== RESP_LABELS.statusPending) {
                roleIndex++;
                var sk = setTimeout(processNextRole, 50);
                respState.timeouts.push(sk);
                return;
            }
            addLogMessage('processRolesWorkflow: processing ' + role.displayRole, 'log');
            updateRespAriaLive(RESP_LABELS.selectingRole + ': ' + role.displayRole);
            ensureRoleListOpenForColumn()
                .then(function() {
                    if (respState.stopRequested) { throw new Error('Stopped'); }
                    return attemptSelectByScrollingForRole(role.displayRole, role.key);
                })
                .then(function(selected) {
                    if (respState.stopRequested) { throw new Error('Stopped'); }
                    if (!selected) {
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
                    updateRespAriaLive(RESP_LABELS.selectingResponsibilities + ': ' + role.displayRole);
                    return openResponsibilitiesDropdown()
                        .then(function() {
                            if (respState.stopRequested) { throw new Error('Stopped'); }
                            return selectResponsibilitiesByNumbers(role.intersection);
                        })
                        .then(function() {
                            if (respState.stopRequested) { throw new Error('Stopped'); }
                            updateRespAriaLive(RESP_LABELS.addingRole + ': ' + role.displayRole);
                            return clickAddStudyRole();
                        })
                        .then(function(addOk) {
                            if (addOk) {
                                role.status = RESP_LABELS.statusCompleted;
                                respState.counters.completed++;
                                respState.counters.pending--;
                                updateRespRoleStatus(role.key, RESP_LABELS.statusCompleted, '');
                            } else {
                                role.status = RESP_LABELS.statusFailed;
                                role.reason = 'Add Study Role failed';
                                respState.counters.failed++;
                                respState.counters.pending--;
                                updateRespRoleStatus(role.key, RESP_LABELS.statusFailed, 'Add Study Role failed');
                            }
                            updateRespSummary();
                            roleIndex++;
                            var nt = setTimeout(processNextRole, 500);
                            respState.timeouts.push(nt);
                        });
                })
                .catch(function(err) {
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

    function stopResponsibilities() {
        addLogMessage('stopResponsibilities: stopping', 'log');
        respState.stopRequested = true;
        respState.isRunning = false;
        for (var i = 0; i < respState.idleCallbackIds.length; i++) {
            try { if (typeof cancelIdleCallback === 'function') { cancelIdleCallback(respState.idleCallbackIds[i]); } } catch (e) {}
        }
        respState.idleCallbackIds = [];
        for (var i2 = 0; i2 < respState.observers.length; i2++) {
            try { respState.observers[i2].disconnect(); } catch (e2) {}
        }
        respState.observers = [];
        for (var i3 = 0; i3 < respState.timeouts.length; i3++) {
            try { clearTimeout(respState.timeouts[i3]); } catch (e3) {}
        }
        respState.timeouts = [];
        for (var i4 = 0; i4 < respState.intervals.length; i4++) {
            try { clearInterval(respState.intervals[i4]); } catch (e4) {}
        }
        respState.intervals = [];
        for (var i5 = 0; i5 < respState.eventListeners.length; i5++) {
            try { var l = respState.eventListeners[i5]; l.element.removeEventListener(l.type, l.handler); } catch (e5) {}
        }
        respState.eventListeners = [];
        respSetAriaBusyOff();
        var im = document.getElementById('resp-input-modal');
        if (im && im.parentNode) { im.parentNode.removeChild(im); }
        var pm = document.getElementById('resp-progress-modal');
        if (pm && pm.parentNode) { pm.parentNode.removeChild(pm); }
        var wm = document.getElementById('resp-warning-modal');
        if (wm && wm.parentNode) { wm.parentNode.removeChild(wm); }
        if (respState.focusReturnElement) { respState.focusReturnElement.focus(); }
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
        closeButton.innerHTML = 'âœ•';
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
        closeButton.innerHTML = 'âœ•';
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
            } else if (i === 3) {
                button.textContent = 'Set Responsibilities';
                button.id = 'resp-set-btn';
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
            } else if (i === 3) {
                button.onclick = () => {
                    console.log('Set Responsibilities button clicked');
                    setResponsibilitiesInit();
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
        closeButton.innerHTML = 'âœ•';
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
        closeButton.innerHTML = 'âœ•';
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
            <strong style="display: block; margin-bottom: 4px; font-size: 14px;">âš ï¸ Important</strong>
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
        checkIcon.innerHTML = 'âœ“';
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
        closeButton.innerHTML = 'âœ•';
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