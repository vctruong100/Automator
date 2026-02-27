
// ==UserScript==
// @name        ClinSpark Automator
// @namespace   vinh.activity.plan.state
// @version     1.9.0
// @description Automate various tasks in ClinSpark platform
// @match       https://cenexel.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Automator.js
// @run-at      document-idle
// @grant       GM.openInTab
// @grant       GM_openInTab
// @grant       GM.xmlHttpRequest
// ==/UserScript==

(function () {
    var STORAGE_KEY = "activityPlanState.run";
    var STORAGE_PENDING = "activityPlanState.pendingIds";
    var STORAGE_AFTER_REFRESH = "activityPlanState.afterRefresh";
    var STORAGE_EDIT_STUDY = "activityPlanState.editStudy";
    var STORAGE_RUN_MODE = "activityPlanState.runMode";
    var STORAGE_CONTINUE_EPOCH = "activityPlanState.continueEpoch";
    var STORAGE_PANEL_TOP = "activityPlanState.panel.top";
    var STORAGE_PANEL_RIGHT = "activityPlanState.panel.right";
    var LIST_URL = "https://cenexel.clinspark.com/secure/crfdesign/activityplans/list";
    var STUDY_SHOW_URL = "https://cenexel.clinspark.com/secure/administration/studies/show";
    var PANEL_ID = "activityPlanStatePanel";
    var LOG_ID = "activityPlanStateLog";
    var STORAGE_SELECTED_IDS = "activityPlanState.selectedVolunteerIds";
    var STORAGE_IC_BARCODE = "activityPlanState.ic.barcode";
    var STORAGE_CONSENT_SCAN_INDEX = "activityPlanState.consent.scanIndex";
    var STORAGE_PAUSED = "activityPlanState.paused";
    var STORAGE_CHECK_ELIG_LOCK = "activityPlanState.checkEligLock";
    var btnRowRef = null;
    var PENDING_BUTTONS = [];
    var STUDY_METADATA_URL = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/show/studymetadata";
    var STORAGE_BARCODE_SUBJECT_TEXT = "activityPlanState.barcode.subjectText";
    var STORAGE_BARCODE_SUBJECT_ID = "activityPlanState.barcode.subjectId";
    var STORAGE_BARCODE_RESULT = "activityPlanState.barcode.result";
    var STORAGE_PANEL_COLLAPSED = "activityPlanState.panel.collapsed";
    var STORAGE_PANEL_WIDTH = "activityPlanState.panel.width";
    var STORAGE_PANEL_HEIGHT = "activityPlanState.panel.height";
    var STORAGE_LOG_VISIBLE = "activityPlanState.log.visible";

    // UI Scale Constants
    var UI_SCALE = 1.0; // Master scale factor (will be initialized after function definitions)
    var PANEL_DEFAULT_WIDTH = 340;
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 30;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 60;
    var PANEL_PADDING_PX = 12;
    var PANEL_BORDER_RADIUS_PX = 8;
    var PANEL_FONT_SIZE_PX = 14;
    var HEADER_FONT_WEIGHT = "600";
    var HEADER_TITLE_OFFSET_PX = 16;
    var HEADER_LEFT_SPACER_WIDTH_PX = 32;
    var BUTTON_BORDER_RADIUS_PX = 6;
    var BUTTON_PADDING_PX = 8;
    var BUTTON_GAP_PX = 8;
    var STATUS_MARGIN_TOP_PX = 10;
    var STATUS_PADDING_PX = 6;
    var STATUS_FONT_SIZE_PX = 13;
    var STATUS_BORDER_RADIUS_PX = 6;
    var LOG_MARGIN_TOP_PX = 8;
    var LOG_HEIGHT_PX = 220;
    var LOG_PADDING_PX = 6;
    var LOG_FONT_SIZE_PX = 12;
    var LOG_BORDER_RADIUS_PX = 6;
    var CLOSE_BTN_TEXT = "✕";
    var STORAGE_UI_SCALE = "activityPlanState.ui.scale";

    var FORM_DELAY_MS = 800;
    var DELAY_BETWEEN_ITEMS_MS = 100;
    var DELAY_AFTER_COLLECT_CLICK_MS = 500;
    var DELAY_BEFORE_SAVE_RETURN_MS = 500;
    var DELAY_BEFORE_FORM_DETAILS_CLICK_MS = 3000;
    var DELAY_AFTER_FORM_DETAILS_CLICK_MS = 1000;
    var RUN_FORM_START_TS = 0;
    var BARCODE_START_TS = 0;

    var AE_LIST_BASE_URL = "https://cenexel.clinspark.com/secure/study/data/list?search=true";
    var AE_LIST_TEST_BASE_URL = "https://cenexel.clinspark.com/secure/study/data/list?search=true";
    var AE_EVENT_KEYWORDS = ["AE/CM","Adverse Event","AE","ADVERSE EVENT","Adverse"];

    var AE_POPUP_TITLE = "Find Adverse Event";
    var AE_POPUP_LABEL = "Subject Identifier";
    var AE_POPUP_OK_TEXT = "Continue";
    var AE_POPUP_CANCEL_TEXT = "Cancel";

    var FORM_LIST_URL = "https://cenexel.clinspark.com/secure/study/data/list";
    var FORM_POPUP_TITLE = "Find Form";
    var FORM_POPUP_KEYWORD_LABEL = "Form Keyword";
    var FORM_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var FORM_POPUP_OK_TEXT = "Continue";
    var FORM_POPUP_CANCEL_TEXT = "Cancel";

    var FORM_NO_MATCH_TITLE = "Find Form";
    var FORM_NO_MATCH_MESSAGE = "No form is found.";

    var BARCODE_BG_TAB = null;

    var STORAGE_FIND_FORM_PENDING = "activityPlanState.findForm.pending";
    var STORAGE_FIND_FORM_KEYWORD = "activityPlanState.findForm.keyword";
    var STORAGE_FIND_FORM_SUBJECT = "activityPlanState.findForm.subject";
    var STORAGE_FIND_FORM_STATUS_VALUES = "activityPlanState.findForm.statusValues";

    var STORAGE_FIND_STUDY_EVENT_PENDING = "activityPlanState.findStudyEvent.pending";
    var STORAGE_FIND_STUDY_EVENT_KEYWORD = "activityPlanState.findStudyEvent.keyword";
    var STORAGE_FIND_STUDY_EVENT_SUBJECT = "activityPlanState.findStudyEvent.subject";
    var STORAGE_FIND_STUDY_EVENT_STATUS_VALUES = "activityPlanState.findStudyEvent.statusValues";

    var STUDY_EVENT_POPUP_TITLE = "Find Study Events";
    var STUDY_EVENT_POPUP_KEYWORD_LABEL = "Study Event Keyword";
    var STUDY_EVENT_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var STUDY_EVENT_POPUP_OK_TEXT = "Continue";
    var STUDY_EVENT_POPUP_CANCEL_TEXT = "Cancel";

    var STUDY_EVENT_NO_MATCH_TITLE = "Find Study Events";
    var STUDY_EVENT_NO_MATCH_MESSAGE = "No study event is found.";


    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const STORAGE_PANEL_HOTKEY = "activityPlanState.panel.hotkey";
    const PANEL_TOGGLE_KEY = "F2";
    const RUNMODE_CLEAR_MAPPING = "clearMapping";

    const ELIGIBILITY_LIST_URL = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";

    // Run Parse Method
    var STORAGE_PARSE_METHOD_RUNNING = "activityPlanState.parseMethod.running";
    var STORAGE_PARSE_METHOD_ITEM_NAME = "activityPlanState.parseMethod.itemName";
    var STORAGE_PARSE_METHOD_RESULTS = "activityPlanState.parseMethod.results";
    var STORAGE_PARSE_METHOD_COMPLETED = "activityPlanState.parseMethod.completed";
    var RUNMODE_PARSE_METHOD = "parseMethod";
    var METHOD_LIBRARY_URL = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/list/method";
    var PARSE_METHOD_CANCELED = false;
    var PARSE_METHOD_COLLECTED_METHODS = [];
    var PARSE_METHOD_COLLECTED_FORMS = [];

    // Cohort Eligibility Feature
    var STORAGE_COHORT_ELIG_DATA = "activityPlanState.cohortElig.data";
    var STORAGE_COHORT_ELIG_RUNNING = "activityPlanState.cohortElig.running";
    var STORAGE_COHORT_ELIG_AUTO_TAB = "activityPlanState.cohortElig.autoTab";
    var COHORT_ELIG_STUDY_SHOW_URL = "https://cenexel.clinspark.com/secure/administration/studies/show";
    var COHORT_ELIG_SUBJECTS_LIST_URL = "https://cenexel.clinspark.com/secure/study/subjects/list";
    var COHORT_ELIG_CANCELED = false;
    var COHORT_ELIG_POPUP = null;
    var COHORT_ELIG_POPUP_CONTENT = null;
    var COHORT_ELIG_LOG_CONTAINER = null;

    // Subject Eligibility Feature
    var STORAGE_SUBJECT_ELIG_PENDING = "activityPlanState.subjectElig.pending";
    var STORAGE_SUBJECT_ELIG_IDENTIFIER = "activityPlanState.subjectElig.identifier";
    var STORAGE_SUBJECT_ELIG_AUTO_TAB = "activityPlanState.subjectElig.autoTab";
    var SUBJECT_ELIG_SUBJECTS_LIST_URL = "https://cenexel.clinspark.com/secure/study/subjects/list";
    var SUBJECT_ELIG_POPUP = null;


    //==========================
    // PARSE DEVIATION FEATURE
    //==========================
    // This section contains all functions related to the Parse Deviation feature.
    // This feature automates extracting deviation form data from the study data list
    // and formatting it for Excel output.
    //==========================

    // Parse Deviation Feature
    var STORAGE_PARSE_DEVIATION_RUNNING = "activityPlanState.parseDeviation.running";
    var STORAGE_PARSE_DEVIATION_KEYWORDS = "activityPlanState.parseDeviation.keywords";
    var STORAGE_PARSE_DEVIATION_RESULTS = "activityPlanState.parseDeviation.results";
    var STORAGE_PARSE_DEVIATION_TIMESTAMP = "activityPlanState.parseDeviation.timestamp";
    var PARSE_DEVIATION_DATA_LIST_URL = "https://cenexel.clinspark.com/secure/study/data/list";
    var PARSE_DEVIATION_POPUP_REF = null;
    var PARSE_DEVIATION_CANCELED = false;

    function APS_ParseDeviation() {
        log("[ParseDeviation] Starting...");

        var popupContainer = document.createElement("div");
        popupContainer.style.display = "flex";
        popupContainer.style.flexDirection = "column";
        popupContainer.style.gap = "16px";

        var instructionText = document.createElement("div");
        instructionText.style.fontSize = "14px";
        instructionText.style.color = "#ccc";
        instructionText.style.marginBottom = "8px";
        instructionText.innerHTML = "Enter subject number(s) to search for deviation forms.<br>" +
            "<span style='font-size:12px;color:#888;'>Separate multiple subjects with commas or newlines.</span>";
        popupContainer.appendChild(instructionText);

        var textArea = document.createElement("textarea");
        textArea.style.width = "100%";
        textArea.style.height = "120px";
        textArea.style.background = "#1a1a1a";
        textArea.style.color = "#fff";
        textArea.style.border = "1px solid #444";
        textArea.style.borderRadius = "6px";
        textArea.style.padding = "10px";
        textArea.style.fontSize = "14px";
        textArea.style.fontFamily = "monospace";
        textArea.style.resize = "vertical";
        textArea.style.boxSizing = "border-box";
        textArea.placeholder = "Example:\n4-002, 4-005\n0043\n0000700004";
        popupContainer.appendChild(textArea);

        var buttonsRow = document.createElement("div");
        buttonsRow.style.display = "flex";
        buttonsRow.style.gap = "12px";
        buttonsRow.style.justifyContent = "flex-end";

        var clearAllBtn = document.createElement("button");
        clearAllBtn.textContent = "Clear All";
        clearAllBtn.style.background = "#6c757d";
        clearAllBtn.style.color = "#fff";
        clearAllBtn.style.border = "none";
        clearAllBtn.style.borderRadius = "6px";
        clearAllBtn.style.padding = "10px 20px";
        clearAllBtn.style.cursor = "pointer";
        clearAllBtn.style.fontSize = "14px";
        clearAllBtn.style.fontWeight = "500";
        clearAllBtn.style.transition = "background 0.2s";
        clearAllBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        clearAllBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#dc3545";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "none";
        confirmBtn.style.borderRadius = "6px";
        confirmBtn.style.padding = "10px 20px";
        confirmBtn.style.cursor = "pointer";
        confirmBtn.style.fontSize = "14px";
        confirmBtn.style.fontWeight = "500";
        confirmBtn.style.transition = "background 0.2s";
        confirmBtn.disabled = true;
        confirmBtn.style.opacity = "0.5";

        function updateConfirmState() {
            var val = textArea.value.trim();
            if (val.length > 0) {
                confirmBtn.disabled = false;
                confirmBtn.style.opacity = "1";
                confirmBtn.style.cursor = "pointer";
            } else {
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = "0.5";
                confirmBtn.style.cursor = "not-allowed";
            }
        }

        textArea.addEventListener("input", updateConfirmState);

        confirmBtn.onmouseenter = function() {
            if (!confirmBtn.disabled) this.style.background = "#c82333";
        };
        confirmBtn.onmouseleave = function() {
            if (!confirmBtn.disabled) this.style.background = "#dc3545";
        };

        buttonsRow.appendChild(clearAllBtn);
        buttonsRow.appendChild(confirmBtn);
        popupContainer.appendChild(buttonsRow);

        var popup = createPopup({
            title: "Parse Deviation",
            content: popupContainer,
            width: "450px",
            height: "auto",
            onClose: function() {
                log("[ParseDeviation] Cancelled by user (close button)");
                PARSE_DEVIATION_CANCELED = true;
            }
        });

        PARSE_DEVIATION_POPUP_REF = popup;

        clearAllBtn.addEventListener("click", function() {
            var confirmClear = confirm("Are you sure you want to clear all input?");
            if (confirmClear) {
                textArea.value = "";
                updateConfirmState();
                log("[ParseDeviation] Input cleared");
            }
        });

        confirmBtn.addEventListener("click", async function() {
            var rawInput = textArea.value;
            var keywords = parseDeviationParseKeywords(rawInput);

            if (keywords.length === 0) {
                log("[ParseDeviation] No valid keywords found");
                return;
            }

            log("[ParseDeviation] Keywords: " + keywords.join(", "));

            var loadingContainer = document.createElement("div");
            loadingContainer.style.display = "flex";
            loadingContainer.style.flexDirection = "column";
            loadingContainer.style.alignItems = "center";
            loadingContainer.style.gap = "20px";
            loadingContainer.style.padding = "40px 20px";

            var spinner = document.createElement("div");
            spinner.style.width = "48px";
            spinner.style.height = "48px";
            spinner.style.border = "4px solid #333";
            spinner.style.borderTop = "4px solid #dc3545";
            spinner.style.borderRadius = "50%";
            spinner.style.animation = "parseDeviationSpin 1s linear infinite";

            var style = document.createElement("style");
            style.textContent = "@keyframes parseDeviationSpin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }";
            document.head.appendChild(style);

            var statusText = document.createElement("div");
            statusText.style.fontSize = "16px";
            statusText.style.fontWeight = "500";
            statusText.style.color = "#f99";
            statusText.textContent = "Navigating to data list...";

            var progressText = document.createElement("div");
            progressText.style.fontSize = "13px";
            progressText.style.color = "#888";
            progressText.textContent = "Please wait...";

            loadingContainer.appendChild(spinner);
            loadingContainer.appendChild(statusText);
            loadingContainer.appendChild(progressText);

            popup.setContent(loadingContainer);

            try {
                localStorage.setItem(STORAGE_PARSE_DEVIATION_KEYWORDS, JSON.stringify(keywords));
                localStorage.setItem(STORAGE_PARSE_DEVIATION_RUNNING, "true");
                localStorage.setItem(STORAGE_PARSE_DEVIATION_TIMESTAMP, String(Date.now()));
            } catch (e) {}

            await sleep(500);

            if (location.href.indexOf("/secure/study/data/list") === -1) {
                statusText.textContent = "Navigating to data list page...";
                progressText.textContent = "Will continue processing after page loads";
                await sleep(500);
                location.href = PARSE_DEVIATION_DATA_LIST_URL;
                return;
            }

            await parseDeviationProcessPage(popup, statusText, progressText, spinner, keywords);
        });
    }

    function parseDeviationParseKeywords(rawInput) {
        var keywords = [];
        var lines = rawInput.split(/[\n\r]+/);
        for (var i = 0; i < lines.length; i++) {
            var parts = lines[i].split(",");
            for (var j = 0; j < parts.length; j++) {
                var kw = parts[j].trim();
                if (kw.length > 0) {
                    keywords.push(kw);
                }
            }
        }
        return keywords;
    }

    async function parseDeviationProcessPage(popup, statusText, progressText, spinner, keywords) {
        log("[ParseDeviation] Processing data list page");

        statusText.textContent = "Selecting deviation forms...";
        progressText.textContent = "Resetting form selections";

        await sleep(500);

        var formSelect = document.querySelector("select[name='formIds']");
        if (!formSelect) {
            log("[ParseDeviation] Form select not found");
            statusText.textContent = "Error: Form select not found";
            statusText.style.color = "#f66";
            spinner.style.display = "none";
            return;
        }

        var allFormOpts = formSelect.querySelectorAll("option");
        for (var fi = 0; fi < allFormOpts.length; fi++) {
            allFormOpts[fi].selected = false;
        }
        var formClearEvt = new Event("change", { bubbles: true });
        formSelect.dispatchEvent(formClearEvt);

        var formSelect2Container = document.querySelector("#s2id_formIds");
        if (formSelect2Container) {
            var formSelect2Choices = formSelect2Container.querySelectorAll(".select2-search-choice");
            for (var fci = 0; fci < formSelect2Choices.length; fci++) {
                var formCloseBtn = formSelect2Choices[fci].querySelector(".select2-search-choice-close");
                if (formCloseBtn) formCloseBtn.click();
            }
        }

        await sleep(300);
        log("[ParseDeviation] Cleared existing form selections");
        progressText.textContent = "Scanning form options";

        var formOptions = formSelect.querySelectorAll("option");
        var deviationKeywords = ["dev", "dpn", "deviation"];
        var selectedFormCount = 0;

        for (var i = 0; i < formOptions.length; i++) {
            var opt = formOptions[i];
            var optText = (opt.textContent || "").toLowerCase();
            for (var k = 0; k < deviationKeywords.length; k++) {
                if (optText.indexOf(deviationKeywords[k]) !== -1) {
                    opt.selected = true;
                    selectedFormCount++;
                    log("[ParseDeviation] Selected form: " + opt.textContent);
                    break;
                }
            }
        }

        if (selectedFormCount === 0) {
            log("[ParseDeviation] No deviation forms found");
            statusText.textContent = "No deviation forms found";
            statusText.style.color = "#f66";
            spinner.style.display = "none";
            progressText.textContent = "Could not find forms matching: dev, DPN, Deviation, PD";
            return;
        }

        var evt = new Event("change", { bubbles: true });
        formSelect.dispatchEvent(evt);

        log("[ParseDeviation] Selected " + selectedFormCount + " deviation form(s)");
        progressText.textContent = "Selected " + selectedFormCount + " form(s)";

        await sleep(500);

        statusText.textContent = "Selecting subjects...";
        progressText.textContent = "Resetting subject selections";

        var subjectSelect = document.querySelector("select[name='subjectIds']");
        if (!subjectSelect) {
            log("[ParseDeviation] Subject select not found");
            statusText.textContent = "Error: Subject select not found";
            statusText.style.color = "#f66";
            spinner.style.display = "none";
            return;
        }

        var allSubjectOpts = subjectSelect.querySelectorAll("option");
        for (var ci = 0; ci < allSubjectOpts.length; ci++) {
            allSubjectOpts[ci].selected = false;
        }
        var clearEvt = new Event("change", { bubbles: true });
        subjectSelect.dispatchEvent(clearEvt);

        var select2Container = document.querySelector("#s2id_subjectIds");
        if (select2Container) {
            var select2Choices = select2Container.querySelectorAll(".select2-search-choice");
            for (var sci = 0; sci < select2Choices.length; sci++) {
                var closeBtn = select2Choices[sci].querySelector(".select2-search-choice-close");
                if (closeBtn) closeBtn.click();
            }
        }

        await sleep(300);
        log("[ParseDeviation] Cleared existing subject selections");
        progressText.textContent = "Matching subjects to keywords";

        var subjectOptions = subjectSelect.querySelectorAll("option");
        var selectedSubjectCount = 0;

        for (var si = 0; si < subjectOptions.length; si++) {
            var subOpt = subjectOptions[si];
            var subText = (subOpt.textContent || "").trim();
            for (var ki = 0; ki < keywords.length; ki++) {
                if (subText.toLowerCase().indexOf(keywords[ki].toLowerCase()) !== -1) {
                    subOpt.selected = true;
                    selectedSubjectCount++;
                    log("[ParseDeviation] Selected subject: " + subText);
                    break;
                }
            }
        }

        if (selectedSubjectCount === 0) {
            log("[ParseDeviation] No matching subjects found");
            statusText.textContent = "No matching subjects found";
            statusText.style.color = "#f66";
            spinner.style.display = "none";
            progressText.textContent = "Keywords: " + keywords.join(", ");
            return;
        }

        var subEvt = new Event("change", { bubbles: true });
        subjectSelect.dispatchEvent(subEvt);

        log("[ParseDeviation] Selected " + selectedSubjectCount + " subject(s)");
        progressText.textContent = "Selected " + selectedSubjectCount + " subject(s)";

        await sleep(500);

        statusText.textContent = "Searching...";
        progressText.textContent = "Clicking search button";

        var searchBtn = document.querySelector("button#dataSearchButton");
        if (!searchBtn) {
            searchBtn = document.querySelector("button[id='dataSearchButton']");
        }
        if (!searchBtn) {
            log("[ParseDeviation] Search button not found");
            statusText.textContent = "Error: Search button not found";
            statusText.style.color = "#f66";
            spinner.style.display = "none";
            return;
        }

        searchBtn.click();
        log("[ParseDeviation] Search button clicked");

        statusText.textContent = "Waiting for results...";
        progressText.textContent = "Loading data...";

        await parseDeviationWaitForResults(5000);

        await sleep(3000);

        statusText.textContent = "Parsing deviation table...";
        progressText.textContent = "Extracting data from rows";

        var allResults = [];
        var pageNum = 1;
        var hasMorePages = true;

        while (hasMorePages) {
            if (PARSE_DEVIATION_CANCELED) {
                log("[ParseDeviation] Process canceled by user");
                return;
            }

            log("[ParseDeviation] Parsing page " + pageNum);
            progressText.textContent = "Parsing page " + pageNum + "...";

            var pageResults = await parseDeviationParseTable();
            allResults = allResults.concat(pageResults);

            log("[ParseDeviation] Found " + pageResults.length + " items on page " + pageNum);

            if (PARSE_DEVIATION_CANCELED) {
                log("[ParseDeviation] Process canceled by user");
                return;
            }

            var nextPage = parseDeviationGetNextPage();
            if (nextPage) {
                nextPage.click();
                await sleep(1500);
                pageNum++;
            } else {
                hasMorePages = false;
            }
        }

        log("[ParseDeviation] Total results: " + allResults.length);

        statusText.textContent = "Fetching subject initials...";
        progressText.textContent = "Processing " + allResults.length + " deviation(s)";

        for (var ri = 0; ri < allResults.length; ri++) {
            if (PARSE_DEVIATION_CANCELED) {
                log("[ParseDeviation] Process canceled by user");
                return;
            }

            if (!allResults[ri].initials || allResults[ri].initials.length === 0) {
                var initials = await parseDeviationFetchInitials(allResults[ri].subjectLink);
                allResults[ri].initials = initials;
            }
        }

        try {
            localStorage.removeItem(STORAGE_PARSE_DEVIATION_RUNNING);
            localStorage.removeItem(STORAGE_PARSE_DEVIATION_KEYWORDS);
        } catch (e) {}

        parseDeviationShowResults(popup, allResults);
    }

    async function parseDeviationWaitForResults(timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var tbody = document.querySelector("tbody#studyDataTableBody");
            if (tbody) {
                var rows = tbody.querySelectorAll("tr");
                if (rows.length > 0) {
                    log("[ParseDeviation] Results loaded: " + rows.length + " rows");
                    return true;
                }
            }
            await sleep(300);
        }
        log("[ParseDeviation] Timeout waiting for results");
        return false;
    }

    async function parseDeviationParseTable() {
        var results = [];
        var tbody = document.querySelector("tbody#studyDataTableBody");
        if (!tbody) {
            log("[ParseDeviation] Table body not found");
            return results;
        }

        var rows = tbody.querySelectorAll("tr");
        log("[ParseDeviation] Parsing " + rows.length + " rows");

        var currentForm = null;
        var currentSubject = "";
        var seenLabelsForForm = {};

        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var cells = row.querySelectorAll("td");
            if (cells.length < 5) continue;

            var statusCell = cells[2];
            var canceledIcon = statusCell.querySelector('span.tooltips[data-original-title="Form Canceled"]');
            if (canceledIcon) {
                log("[ParseDeviation] Row " + i + ": Skipping canceled form");
                continue;
            }

            var subjectCell = cells[1];
            var subjectLink = subjectCell.querySelector("a");
            var subjectNumber = "";
            var subjectHref = "";

            if (subjectLink) {
                subjectNumber = (subjectLink.textContent || "").trim();
                subjectHref = subjectLink.getAttribute("href") || "";
            }

            log("[ParseDeviation] Row " + i + ": Subject = '" + subjectNumber + "'");

            var dataCell = cells[4];
            var miniTables = dataCell.querySelectorAll("tbody");
            log("[ParseDeviation] Row " + i + ": Found " + miniTables.length + " mini-tables (item groups)");

            for (var ti = 0; ti < miniTables.length; ti++) {
                var miniRows = miniTables[ti].querySelectorAll("tr");
                var itemGroupData = {};

                for (var mi = 0; mi < miniRows.length; mi++) {
                    var miniCells = miniRows[mi].querySelectorAll("td");
                    if (miniCells.length < 2) continue;

                    var labelCell = miniCells[0];
                    var valueCell = miniCells[1];

                    var label = (labelCell.textContent || "").trim().toLowerCase();
                    var valueEl = valueCell.querySelector("span.novalue");
                    var value = "";
                    if (valueEl) {
                        value = "";
                    } else {
                        value = (valueCell.textContent || "").trim();
                    }

                    itemGroupData[label] = value;
                }

                log("[ParseDeviation] Item group " + ti + " data: " + JSON.stringify(itemGroupData));

                var isNewForm = false;
                if (subjectNumber !== currentSubject) {
                    log("[ParseDeviation] NEW FORM: Subject changed from '" + currentSubject + "' to '" + subjectNumber + "'");
                    isNewForm = true;
                } else {
                    for (var key in itemGroupData) {
                        if (key.indexOf("date of deviation") !== -1 && seenLabelsForForm[key]) {
                            var newDateVal = itemGroupData[key];
                            log("[ParseDeviation] Checking date of deviation: '" + newDateVal + "' (seen before: " + seenLabelsForForm[key] + ")");
                            if (newDateVal && newDateVal.length > 0 && newDateVal !== "-") {
                                log("[ParseDeviation] NEW FORM: Date of deviation has non-empty value");
                                isNewForm = true;
                                break;
                            }
                        }
                    }
                }

                if (isNewForm && currentForm) {
                    log("[ParseDeviation] Pushing completed form to results");
                    currentForm.actionTakenCombined = currentForm.actionTaken.join("; ");
                    delete currentForm.actionTaken;
                    results.push(currentForm);
                    currentForm = null;
                    seenLabelsForForm = {};
                }

                if (!currentForm) {
                    log("[ParseDeviation] Creating new form for subject '" + subjectNumber + "'");
                    currentForm = {
                        subjectNumber: subjectNumber,
                        subjectLink: subjectHref,
                        dateOfDeviation: "",
                        studyVisitDay: "",
                        dateDiscovered: "",
                        deviationExplanation: "",
                        severity: "",
                        reportable: "",
                        initials: "",
                        category: "",
                        actionTaken: []
                    };
                    currentSubject = subjectNumber;
                }

                for (var key in itemGroupData) {
                    var val = itemGroupData[key];
                    if (val && val.length > 0 && val !== "-") {
                        seenLabelsForForm[key] = true;

                        if (key.indexOf("date of deviation") !== -1) {
                            if (!currentForm.dateOfDeviation || currentForm.dateOfDeviation === "-") currentForm.dateOfDeviation = val;
                        } else if (key.indexOf("study visit day") !== -1) {
                            if (!currentForm.studyVisitDay || currentForm.studyVisitDay === "-") currentForm.studyVisitDay = val;
                        } else if (key.indexOf("date discovered") !== -1) {
                            if (!currentForm.dateDiscovered || currentForm.dateDiscovered === "-") currentForm.dateDiscovered = val;
                        } else if (key.indexOf("description") !== -1 || key.indexOf("explanation") !== -1) {
                            if (!currentForm.deviationExplanation || currentForm.deviationExplanation === "-") currentForm.deviationExplanation = val;
                        } else if (key.indexOf("severity") !== -1) {
                            if (!currentForm.severity || currentForm.severity === "-") currentForm.severity = val;
                        } else if (key.indexOf("reportable") !== -1) {
                            if (!currentForm.reportable || currentForm.reportable === "-") currentForm.reportable = val;
                        } else if (key.indexOf("initial") !== -1) {
                            if (!currentForm.initials || currentForm.initials === "-") currentForm.initials = val;
                        } else if (key.indexOf("category") !== -1) {
                            if (!currentForm.category || currentForm.category === "-") currentForm.category = val;
                        } else if (key.indexOf("corrective") !== -1 || key.indexOf("preventative") !== -1 || key.indexOf("preventive") !== -1) {
                            currentForm.actionTaken.push(val);
                        }
                    }
                }
                log("[ParseDeviation] Current form state: " + JSON.stringify(currentForm));
            }
        }

        if (currentForm) {
            log("[ParseDeviation] Pushing final form to results");
            currentForm.actionTakenCombined = currentForm.actionTaken.join("; ");
            delete currentForm.actionTaken;
            results.push(currentForm);
        }

        log("[ParseDeviation] Total forms created: " + results.length);
        return results;
    }

    function parseDeviationGetNextPage() {
        var pager = document.querySelector("p#pagerButtons");
        if (!pager) {
            var pagerAlt = document.querySelector("#pagerButtons");
            if (pagerAlt) pager = pagerAlt;
        }
        if (!pager) return null;

        var currentDisabled = pager.querySelector("li.disabled a, li.active a");
        if (!currentDisabled) return null;

        var currentLi = currentDisabled.closest("li");
        if (!currentLi) return null;

        var nextLi = currentLi.nextElementSibling;
        while (nextLi) {
            var dataLp = nextLi.getAttribute("data-lp");
            if (dataLp && !nextLi.classList.contains("disabled")) {
                var nextLink = nextLi.querySelector("a");
                if (nextLink) {
                    return nextLink;
                }
            }
            nextLi = nextLi.nextElementSibling;
        }

        return null;
    }

    async function parseDeviationFetchInitials(subjectHref) {
        if (!subjectHref || subjectHref.length === 0) {
            return "";
        }

        try {
            var fullUrl = location.origin + subjectHref;
            log("[ParseDeviation] Fetching initials from: " + fullUrl);

            return new Promise(function(resolve) {
                var iframe = document.createElement("iframe");
                iframe.style.position = "fixed";
                iframe.style.top = "-9999px";
                iframe.style.left = "-9999px";
                iframe.style.width = "1px";
                iframe.style.height = "1px";
                iframe.style.visibility = "hidden";

                var timeoutId = setTimeout(function() {
                    if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                    resolve("");
                }, 10000);

                iframe.onload = function() {
                    try {
                        var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
                        var volunteerLink = iframeDoc.querySelector("a[href*='/secure/volunteers/manage/show/']");
                        if (volunteerLink) {
                            var linkText = (volunteerLink.textContent || "").trim();
                            var parts = linkText.split(",");
                            if (parts.length >= 3) {
                                var initials = parts[2].trim();
                                log("[ParseDeviation] Found initials: " + initials);
                                clearTimeout(timeoutId);
                                if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                                resolve(initials);
                                return;
                            }
                        }
                        clearTimeout(timeoutId);
                        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                        resolve("");
                    } catch (e) {
                        clearTimeout(timeoutId);
                        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                        resolve("");
                    }
                };

                iframe.onerror = function() {
                    clearTimeout(timeoutId);
                    if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                    resolve("");
                };

                document.body.appendChild(iframe);
                iframe.src = fullUrl;
            });
        } catch (e) {
            log("[ParseDeviation] Error fetching initials: " + e);
            return "";
        }
    }

    function parseDeviationFormatDate(dateStr) {
        if (!dateStr || dateStr.length === 0) return "";

        var months = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
            "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
        };

        var match = dateStr.match(/(\d{1,2})\s*([A-Za-z]{3})\s*(\d{4})/);
        if (match) {
            var day = parseInt(match[1], 10);
            var monthStr = match[2].toLowerCase();
            var year = parseInt(match[3], 10);
            var month = months[monthStr] || 1;
            return month + "/" + day + "/" + year;
        }

        return dateStr;
    }

    function parseDeviationFormatVisitDay(visitDayStr) {
        if (!visitDayStr || visitDayStr.length === 0) return "";

        var match = visitDayStr.match(/[Dd]ay\s*(\d+)/i);
        if (match) {
            return "D" + match[1];
        }

        var numMatch = visitDayStr.match(/^(\d+)$/);
        if (numMatch) {
            return "D" + numMatch[1];
        }

        return visitDayStr;
    }

    function parseDeviationSanitizeValue(val) {
        if (!val) return "";
        return String(val).replace(/[\t\r\n]+/g, " ").trim();
    }

    function parseDeviationFormatSubjectNumber(subjectStr) {
        if (!subjectStr) return "";
        var parts = subjectStr.split("/");
        if (parts.length >= 2) {
            var middle = parts[1].trim();
            if (/[a-zA-Z]/.test(middle)) {
                var allNumbers = middle.match(/\d+(?:-\d+)?/g);
                if (allNumbers) {
                    for (var i = 0; i < allNumbers.length; i++) {
                        var numWithoutDash = allNumbers[i].replace(/-/g, "");
                        if (numWithoutDash.length > 3) {
                            return allNumbers[i];
                        }
                    }
                }
                return parts[0].trim();
            }
            return middle;
        }
        return subjectStr.trim();
    }

    function parseDeviationFormatActionTaken(actionStr) {
        if (!actionStr) return "";
        return actionStr.replace(/;\s*/g, " ");
    }

    function parseDeviationFormatSeverity(severityStr) {
        if (!severityStr) return "";
        if (severityStr.toLowerCase().indexOf("minor") !== -1) {
            return "Minor";
        }
        return severityStr;
    }

    function parseDeviationFormatReportable(reportableStr) {
        if (!reportableStr) return "";
        var lower = reportableStr.toLowerCase();
        if (lower.indexOf("not") !== -1 || lower.indexOf("no") !== -1) {
            return "No";
        }
        return "Yes";
    }

    function parseDeviationFormatRow(item) {
        var TAB = "\t";

        var colA = "=IFERROR(INDEX(A:A,ROW()-1)+1,1)";
        var colB = "ACT";
        var colC = "";
        var colD = "";
        var colE = parseDeviationSanitizeValue(parseDeviationFormatSubjectNumber(item.subjectNumber));
        var colF = parseDeviationSanitizeValue(parseDeviationFormatVisitDay(item.studyVisitDay));
        var colG = "";
        var colH = parseDeviationSanitizeValue(parseDeviationFormatDate(item.dateOfDeviation));
        var colI = parseDeviationSanitizeValue(parseDeviationFormatDate(item.dateDiscovered));
        var colJ = parseDeviationSanitizeValue(item.initials);
        var colK = parseDeviationSanitizeValue(item.deviationExplanation);
        var colL = parseDeviationSanitizeValue(parseDeviationFormatActionTaken(item.actionTakenCombined));
        var colM = "";
        var colN = parseDeviationSanitizeValue(item.category);
        var colO = "";
        var colP = "";
        var colQ = "";
        var colR = "";
        var colS = "";
        var colT = parseDeviationSanitizeValue(parseDeviationFormatSeverity(item.severity));
        var colU = parseDeviationSanitizeValue(parseDeviationFormatReportable(item.reportable));

        var row = colA + TAB + colB + TAB + colC + TAB + colD + TAB +
            colE + TAB + colF + TAB + colG + TAB + colH + TAB +
            colI + TAB + colJ + TAB + colK + TAB + colL + TAB +
            colM + TAB + colN + TAB + colO + TAB + colP + TAB +
            colQ + TAB + colR + TAB + colS + TAB + colT + TAB + colU;

        return row;
    }

    function parseDeviationShowResults(popup, results) {
        log("[ParseDeviation] Displaying " + results.length + " results");

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";
        container.style.maxHeight = "500px";
        container.style.overflowY = "auto";

        var headerDiv = document.createElement("div");
        headerDiv.style.display = "flex";
        headerDiv.style.justifyContent = "space-between";
        headerDiv.style.alignItems = "center";
        headerDiv.style.marginBottom = "8px";
        headerDiv.style.padding = "8px";
        headerDiv.style.background = "#1a1a1a";
        headerDiv.style.borderRadius = "6px";

        var countText = document.createElement("div");
        countText.style.fontSize = "14px";
        countText.style.fontWeight = "600";
        countText.textContent = "Found " + results.length + " deviation(s)";
        headerDiv.appendChild(countText);

        var copyAllBtn = document.createElement("button");
        copyAllBtn.textContent = "Copy All";
        copyAllBtn.style.background = "#dc3545";
        copyAllBtn.style.color = "#fff";
        copyAllBtn.style.border = "none";
        copyAllBtn.style.borderRadius = "4px";
        copyAllBtn.style.padding = "6px 12px";
        copyAllBtn.style.cursor = "pointer";
        copyAllBtn.style.fontSize = "12px";
        copyAllBtn.style.fontWeight = "500";
        copyAllBtn.onmouseenter = function() { this.style.background = "#c82333"; };
        copyAllBtn.onmouseleave = function() { this.style.background = "#dc3545"; };
        copyAllBtn.addEventListener("click", function() {
            var allRows = [];
            for (var i = 0; i < results.length; i++) {
                allRows.push(parseDeviationFormatRow(results[i]));
            }
            var allText = allRows.join("\n");
            navigator.clipboard.writeText(allText).then(function() {
                copyAllBtn.textContent = "Copied!";
                setTimeout(function() { copyAllBtn.textContent = "Copy All"; }, 1500);
            });
        });
        headerDiv.appendChild(copyAllBtn);

        container.appendChild(headerDiv);

        if (results.length === 0) {
            var noResults = document.createElement("div");
            noResults.style.textAlign = "center";
            noResults.style.padding = "40px";
            noResults.style.color = "#888";
            noResults.textContent = "No deviation data found.";
            container.appendChild(noResults);
        } else {
            for (var i = 0; i < results.length; i++) {
                var item = results[i];
                var rowDiv = document.createElement("div");
                rowDiv.style.display = "flex";
                rowDiv.style.flexDirection = "column";
                rowDiv.style.gap = "8px";
                rowDiv.style.padding = "12px";
                rowDiv.style.background = "#1a1a1a";
                rowDiv.style.borderRadius = "6px";
                rowDiv.style.border = "1px solid #333";

                var rowHeader = document.createElement("div");
                rowHeader.style.display = "flex";
                rowHeader.style.justifyContent = "space-between";
                rowHeader.style.alignItems = "center";

                var subjectInfo = document.createElement("div");
                subjectInfo.style.fontWeight = "600";
                subjectInfo.style.fontSize = "14px";
                subjectInfo.textContent = item.subjectNumber || "Unknown Subject";
                rowHeader.appendChild(subjectInfo);

                var copyBtn = document.createElement("button");
                copyBtn.textContent = "Copy";
                copyBtn.style.background = "#28a745";
                copyBtn.style.color = "#fff";
                copyBtn.style.border = "none";
                copyBtn.style.borderRadius = "4px";
                copyBtn.style.padding = "4px 10px";
                copyBtn.style.cursor = "pointer";
                copyBtn.style.fontSize = "12px";
                copyBtn.style.fontWeight = "500";
                copyBtn.onmouseenter = function() { this.style.background = "#218838"; };
                copyBtn.onmouseleave = function() { this.style.background = "#28a745"; };

                (function(btn, itemData) {
                    btn.addEventListener("click", function() {
                        var formattedRow = parseDeviationFormatRow(itemData);
                        navigator.clipboard.writeText(formattedRow).then(function() {
                            btn.textContent = "Copied!";
                            setTimeout(function() { btn.textContent = "Copy"; }, 1500);
                        });
                    });
                })(copyBtn, item);

                rowHeader.appendChild(copyBtn);
                rowDiv.appendChild(rowHeader);

                var detailsGrid = document.createElement("div");
                detailsGrid.style.display = "grid";
                detailsGrid.style.gridTemplateColumns = "1fr 1fr";
                detailsGrid.style.gap = "6px";
                detailsGrid.style.fontSize = "12px";

                function addDetail(label, value) {
                    var detailItem = document.createElement("div");
                    detailItem.innerHTML = "<span style='color:#888;'>" + label + ":</span> " +
                        "<span style='color:#fff;'>" + (value || "-") + "</span>";
                    detailsGrid.appendChild(detailItem);
                }

                addDetail("Visit Day", parseDeviationFormatVisitDay(item.studyVisitDay));
                addDetail("Date of Deviation", parseDeviationFormatDate(item.dateOfDeviation));
                addDetail("Date Discovered", parseDeviationFormatDate(item.dateDiscovered));
                addDetail("Initials", item.initials);
                addDetail("Severity", item.severity);
                addDetail("Reportable", item.reportable);

                rowDiv.appendChild(detailsGrid);

                if (item.deviationExplanation) {
                    var explanationDiv = document.createElement("div");
                    explanationDiv.style.fontSize = "12px";
                    explanationDiv.style.color = "#ccc";
                    explanationDiv.style.marginTop = "4px";
                    explanationDiv.style.padding = "8px";
                    explanationDiv.style.background = "#222";
                    explanationDiv.style.borderRadius = "4px";
                    explanationDiv.style.maxHeight = "60px";
                    explanationDiv.style.overflowY = "auto";
                    explanationDiv.textContent = item.deviationExplanation;
                    rowDiv.appendChild(explanationDiv);
                }

                container.appendChild(rowDiv);
            }
        }

        popup.setContent(container);
        popup.element.style.width = "600px";
        popup.element.style.height = "auto";
        popup.element.style.maxHeight = "80%";

        var rect = popup.element.getBoundingClientRect();
        popup.element.style.left = (window.innerWidth / 2) + "px";
        popup.element.style.top = (window.innerHeight / 2) + "px";
        popup.element.style.transform = "translate(-50%, -50%)";
    }

    function parseDeviationCheckOnPageLoad() {
        var isRunning = false;
        try {
            isRunning = localStorage.getItem(STORAGE_PARSE_DEVIATION_RUNNING) === "true";
        } catch (e) {}

        if (!isRunning) return;

        if (location.href.indexOf("/secure/study/data/list") === -1) return;

        // Check timestamp to prevent auto-run on casual navigation
        var timestamp = 0;
        try {
            var tsRaw = localStorage.getItem(STORAGE_PARSE_DEVIATION_TIMESTAMP);
            if (tsRaw) timestamp = parseInt(tsRaw, 10);
        } catch (e) {}

        var now = Date.now();
        var elapsed = now - timestamp;
        var maxAge = 10000; // 10 seconds

        if (elapsed > maxAge || timestamp === 0) {
            log("[ParseDeviation] Timestamp expired or missing (elapsed=" + elapsed + "ms), skipping auto-resume");
            try {
                localStorage.removeItem(STORAGE_PARSE_DEVIATION_RUNNING);
                localStorage.removeItem(STORAGE_PARSE_DEVIATION_KEYWORDS);
                localStorage.removeItem(STORAGE_PARSE_DEVIATION_TIMESTAMP);
            } catch (e) {}
            return;
        }

        var keywords = [];
        try {
            var raw = localStorage.getItem(STORAGE_PARSE_DEVIATION_KEYWORDS);
            if (raw) keywords = JSON.parse(raw);
        } catch (e) {}

        if (keywords.length === 0) {
            try {
                localStorage.removeItem(STORAGE_PARSE_DEVIATION_RUNNING);
                localStorage.removeItem(STORAGE_PARSE_DEVIATION_TIMESTAMP);
            } catch (e) {}
            return;
        }
        log("[ParseDeviation] Resuming after page load with keywords: " + keywords.join(", "));

        var popupContainer = document.createElement("div");
        popupContainer.style.display = "flex";
        popupContainer.style.flexDirection = "column";
        popupContainer.style.alignItems = "center";
        popupContainer.style.gap = "20px";
        popupContainer.style.padding = "40px 20px";

        var spinner = document.createElement("div");
        spinner.style.width = "48px";
        spinner.style.height = "48px";
        spinner.style.border = "4px solid #333";
        spinner.style.borderTop = "4px solid #dc3545";
        spinner.style.borderRadius = "50%";
        spinner.style.animation = "parseDeviationSpin 1s linear infinite";

        var style = document.createElement("style");
        style.textContent = "@keyframes parseDeviationSpin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }";
        document.head.appendChild(style);

        var statusText = document.createElement("div");
        statusText.style.fontSize = "16px";
        statusText.style.fontWeight = "500";
        statusText.style.color = "#f99";
        statusText.textContent = "Resuming Parse Deviation...";

        var progressText = document.createElement("div");
        progressText.style.fontSize = "13px";
        progressText.style.color = "#888";
        progressText.textContent = "Please wait...";

        popupContainer.appendChild(spinner);
        popupContainer.appendChild(statusText);
        popupContainer.appendChild(progressText);

        var popup = createPopup({
            title: "Parse Deviation",
            content: popupContainer,
            width: "450px",
            height: "auto"
        });

        PARSE_DEVIATION_POPUP_REF = popup;

        setTimeout(async function() {
            await parseDeviationProcessPage(popup, statusText, progressText, spinner, keywords);
        }, 1000);
    }

    //=========================
    // SETTING FEATURE
    //=========================
    // This section contains functions for setting up the extension.
    // It includes functions for loading and saving settings, as well as
    // functions for updating the UI based on the current settings.
    //=========================
    var STORAGE_BUTTON_VISIBILITY = "activityPlanState.buttonVisibility";
    var SETTINGS_POPUP_REF = null;

    function getPanelHotkey() {
        try {
            var saved = localStorage.getItem(STORAGE_PANEL_HOTKEY);
            if (saved) {
                return saved;
            }
        } catch (err) {
            log("Error reading hotkey from localStorage: " + String(err));
        }
        return "F2";
    }

    function setPanelHotkey(hotkey) {
        try {
            localStorage.setItem(STORAGE_PANEL_HOTKEY, hotkey);
            return true;
        } catch (err) {
            log("Error saving hotkey to localStorage: " + String(err));
            return false;
        }
    }

    function getButtonVisibility() {
        try {
            var raw = localStorage.getItem(STORAGE_BUTTON_VISIBILITY);
            if (raw) {
                return JSON.parse(raw);
            }
        } catch (e) {}
        return null;
    }

    function setButtonVisibility(visibilityMap) {
        try {
            localStorage.setItem(STORAGE_BUTTON_VISIBILITY, JSON.stringify(visibilityMap));
        } catch (e) {}
    }

    function isButtonVisible(label) {
        var visibility = getButtonVisibility();
        if (!visibility) {
            return true;
        }
        if (visibility.hasOwnProperty(label)) {
            return visibility[label];
        }
        return true;
    }

    function openSettingsPopup() {
        var buttonLabels = [
            "Run Barcode",
            "Scheduled Activities Builder",
            "Search Methods",
            "Parse Deviation",
            "Import I/E",
            "Find Adverse Event",
            "Find Form",
            "Find Study Events",
            "Item Method Forms",
            "Cohort Eligibility",
            "Subject Eligibility",
            "Parse Study Event",
            "Pause",
            "Clear Logs",
            "Hide Logs"
        ];

        var currentVisibility = getButtonVisibility() || {};

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "8px";
        container.style.minWidth = "280px";

        var description = document.createElement("div");
        description.textContent = "Select which buttons to display in the panel:";
        description.style.fontSize = "13px";
        description.style.color = "#aaa";
        description.style.marginBottom = "12px";
        container.appendChild(description);

        var checkboxContainer = document.createElement("div");
        checkboxContainer.style.display = "flex";
        checkboxContainer.style.flexDirection = "column";
        checkboxContainer.style.gap = "6px";
        checkboxContainer.style.maxHeight = "320px";
        checkboxContainer.style.overflowY = "auto";
        checkboxContainer.style.paddingRight = "8px";

        var checkboxes = [];

        for (var i = 0; i < buttonLabels.length; i++) {
            var label = buttonLabels[i];
            var isChecked = currentVisibility.hasOwnProperty(label) ? currentVisibility[label] : true;

            var row = document.createElement("label");
            row.style.display = "flex";
            row.style.alignItems = "center";
            row.style.gap = "10px";
            row.style.padding = "8px 12px";
            row.style.background = "#1a1a1a";
            row.style.borderRadius = "6px";
            row.style.cursor = "pointer";
            row.style.transition = "background 0.15s";
            row.onmouseenter = function() { this.style.background = "#252525"; };
            row.onmouseleave = function() { this.style.background = "#1a1a1a"; };

            var checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = isChecked;
            checkbox.style.width = "18px";
            checkbox.style.height = "18px";
            checkbox.style.cursor = "pointer";
            checkbox.style.accentColor = "#5b43c7";
            checkbox.dataset.label = label;

            var labelText = document.createElement("span");
            labelText.textContent = label;
            labelText.style.fontSize = "14px";
            labelText.style.color = "#fff";

            row.appendChild(checkbox);
            row.appendChild(labelText);
            checkboxContainer.appendChild(row);
            checkboxes.push(checkbox);
        }

        container.appendChild(checkboxContainer);

        // Hotkey Configuration Section
        var hotkeySection = document.createElement("div");
        hotkeySection.style.marginTop = "20px";
        hotkeySection.style.paddingTop = "20px";
        hotkeySection.style.borderTop = "1px solid #333";

        var hotkeyLabel = document.createElement("div");
        hotkeyLabel.textContent = "Panel Toggle Hotkey:";
        hotkeyLabel.style.fontSize = "13px";
        hotkeyLabel.style.color = "#aaa";
        hotkeyLabel.style.marginBottom = "8px";
        hotkeySection.appendChild(hotkeyLabel);

        var hotkeyInputRow = document.createElement("div");
        hotkeyInputRow.style.display = "flex";
        hotkeyInputRow.style.gap = "10px";
        hotkeyInputRow.style.alignItems = "center";

        var hotkeyInput = document.createElement("input");
        hotkeyInput.type = "text";
        hotkeyInput.value = getPanelHotkey();
        hotkeyInput.placeholder = "Press a key...";
        hotkeyInput.style.flex = "1";
        hotkeyInput.style.padding = "10px 12px";
        hotkeyInput.style.background = "#2a2a2a";
        hotkeyInput.style.border = "1px solid #444";
        hotkeyInput.style.borderRadius = "6px";
        hotkeyInput.style.color = "#fff";
        hotkeyInput.style.fontSize = "14px";
        hotkeyInput.style.outline = "none";
        hotkeyInput.style.fontFamily = "monospace";
        hotkeyInput.style.cursor = "pointer";

        hotkeyInput.addEventListener("focus", function() {
            this.style.borderColor = "#5b43c7";
        });

        hotkeyInput.addEventListener("blur", function() {
            this.style.borderColor = "#333";
        });

        hotkeyInput.addEventListener("keydown", function(e) {
            e.preventDefault();
            e.stopPropagation();

            var key = e.key;
            var code = e.code;
            var displayKey = key;

            // Handle special keys
            if (key.length === 1 && key.match(/[a-z]/i)) {
                displayKey = key.toUpperCase();
            } else if (code && code.startsWith("Key")) {
                displayKey = code.substring(3);
            } else if (code && code.startsWith("Digit")) {
                displayKey = code.substring(5);
            } else if (key === " ") {
                displayKey = "Space";
            } else if (key.startsWith("F") && key.length <= 3) {
                displayKey = key.toUpperCase();
            } else if (key === "Escape") {
                displayKey = "Escape";
            } else if (key === "Enter") {
                displayKey = "Enter";
            } else if (key === "Tab") {
                displayKey = "Tab";
            } else if (key === "Backspace") {
                displayKey = "Backspace";
            } else if (code) {
                displayKey = code;
            }

            this.value = displayKey;
            return false;
        });

        // Prevent typing but allow key capture for hotkey capture
        hotkeyInput.addEventListener("keypress", function(e) {
            e.preventDefault();
            return false;
        });

        var hotkeyResetBtn = document.createElement("button");
        hotkeyResetBtn.textContent = "Reset";
        hotkeyResetBtn.style.background = "#333";
        hotkeyResetBtn.style.color = "#fff";
        hotkeyResetBtn.style.border = "none";
        hotkeyResetBtn.style.borderRadius = "6px";
        hotkeyResetBtn.style.padding = "10px 16px";
        hotkeyResetBtn.style.cursor = "pointer";
        hotkeyResetBtn.style.fontSize = "13px";
        hotkeyResetBtn.onmouseenter = function() { this.style.background = "#444"; };
        hotkeyResetBtn.onmouseleave = function() { this.style.background = "#333"; };
        hotkeyResetBtn.addEventListener("click", function() {
            hotkeyInput.value = "F2";
        });

        hotkeyInputRow.appendChild(hotkeyInput);
        hotkeyInputRow.appendChild(hotkeyResetBtn);
        hotkeySection.appendChild(hotkeyInputRow);

        var hotkeyHint = document.createElement("div");
        hotkeyHint.textContent = "Click the input field and press any key to set a new hotkey";
        hotkeyHint.style.fontSize = "11px";
        hotkeyHint.style.color = "#666";
        hotkeyHint.style.marginTop = "6px";
        hotkeyHint.style.fontStyle = "italic";
        hotkeySection.appendChild(hotkeyHint);

        container.appendChild(hotkeySection);

        var buttonRow = document.createElement("div");
        buttonRow.style.display = "flex";
        buttonRow.style.gap = "10px";
        buttonRow.style.marginTop = "16px";
        buttonRow.style.justifyContent = "flex-end";

        var selectAllBtn = document.createElement("button");
        selectAllBtn.textContent = "Select All";
        selectAllBtn.style.background = "#333";
        selectAllBtn.style.color = "#fff";
        selectAllBtn.style.border = "none";
        selectAllBtn.style.borderRadius = "6px";
        selectAllBtn.style.padding = "8px 16px";
        selectAllBtn.style.cursor = "pointer";
        selectAllBtn.style.fontSize = "13px";
        selectAllBtn.onmouseenter = function() { this.style.background = "#444"; };
        selectAllBtn.onmouseleave = function() { this.style.background = "#333"; };
        selectAllBtn.addEventListener("click", function() {
            for (var j = 0; j < checkboxes.length; j++) {
                checkboxes[j].checked = true;
            }
        });

        var deselectAllBtn = document.createElement("button");
        deselectAllBtn.textContent = "Deselect All";
        deselectAllBtn.style.background = "#333";
        deselectAllBtn.style.color = "#fff";
        deselectAllBtn.style.border = "none";
        deselectAllBtn.style.borderRadius = "6px";
        deselectAllBtn.style.padding = "8px 16px";
        deselectAllBtn.style.cursor = "pointer";
        deselectAllBtn.style.fontSize = "13px";
        deselectAllBtn.onmouseenter = function() { this.style.background = "#444"; };
        deselectAllBtn.onmouseleave = function() { this.style.background = "#333"; };
        deselectAllBtn.addEventListener("click", function() {
            for (var j = 0; j < checkboxes.length; j++) {
                checkboxes[j].checked = false;
            }
        });

        var saveBtn = document.createElement("button");
        saveBtn.textContent = "Save & Refresh";
        saveBtn.style.background = "#5b43c7";
        saveBtn.style.color = "#fff";
        saveBtn.style.border = "none";
        saveBtn.style.borderRadius = "6px";
        saveBtn.style.padding = "10px 20px";
        saveBtn.style.cursor = "pointer";
        saveBtn.style.fontSize = "14px";
        saveBtn.style.fontWeight = "600";
        saveBtn.onmouseenter = function() { this.style.background = "#4a35a6"; };
        saveBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };

        buttonRow.appendChild(selectAllBtn);
        buttonRow.appendChild(deselectAllBtn);
        buttonRow.appendChild(saveBtn);
        container.appendChild(buttonRow);

        var settingsPopup = createPopup({
            title: "Settings",
            content: container,
            width: "360px",
            height: "auto",
            maxHeight: "80%"
        });
        var originalClose = settingsPopup.close;
        settingsPopup.close = function() {
            SETTINGS_POPUP_REF = null;
            originalClose();
        };
        saveBtn.addEventListener("click", function() {
            var newVisibility = {};
            for (var j = 0; j < checkboxes.length; j++) {
                var cb = checkboxes[j];
                newVisibility[cb.dataset.label] = cb.checked;
            }
            setButtonVisibility(newVisibility);

            var newHotkey = hotkeyInput.value.trim();
            if (newHotkey) {
                setPanelHotkey(newHotkey);
                log("Settings: Hotkey saved as " + newHotkey);
            }

            log("Settings: Button visibility saved");
            settingsPopup.close();
            location.reload();
        });
        return settingsPopup;
    }

    //==========================
    // PARSE STUDY EVENT FEATURE
    //==========================
    // This section contains functions for parsing study events from the study library.
    // It extracts study event names and generates formatted output for coding use.
    //==========================

    var PARSE_STUDY_EVENT_IGNORE_KEYWORDS = [
        "subject", "screening", "ae/cm", "ae", "cm", "unscheduled", "early term",
        "early termination", "medication", "concomitant", "concommitant", "protocol",
        "adverse", "event"
    ];

    function isStudyEventPage() {
        var url = location.href.toLowerCase();
        var match1 = url.indexOf("cenexel.clinspark.com/secure/crfdesign/studylibrary/list/studyevent") !== -1;
        var match2 = url.indexOf("cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/studyevent") !== -1;
        return match1 || match2;
    }

    function parseStudyEventShouldIgnore(name) {
        var lower = name.toLowerCase();
        for (var i = 0; i < PARSE_STUDY_EVENT_IGNORE_KEYWORDS.length; i++) {
            if (lower.indexOf(PARSE_STUDY_EVENT_IGNORE_KEYWORDS[i]) !== -1) {
                return true;
            }
        }
        return false;
    }

    function parseStudyEventExtractDayNumber(name) {
        var normalized = name.replace(/\s+/g, " ").trim();
        var rangeMatch = normalized.match(/\b(?:day|d)\s*(-?\d+)\s*(?:-|to)\s*(?:day|d)?\s*(-?\d+)/i);
        if (rangeMatch) {
            var n1 = parseInt(rangeMatch[1], 10);
            var n2 = parseInt(rangeMatch[2], 10);
            if (n1 > n2) { var tmp = n1; n1 = n2; n2 = tmp; }
            return { type: "range", min: n1, max: n2, mid: Math.round((n1 + n2) / 2) };
        }
        var singleMatch = normalized.match(/\b(?:day|d)\s*(-?\d+)\b/i);
        if (singleMatch) {
            var num = parseInt(singleMatch[1], 10);
            return { type: "single", value: num };
        }
        return null;
    }

    function parseStudyEventExtractGroupLabel(name) {
        var match = name.match(/^\s*(\([^)]+\)|\[[^\]]+\])\s*/);
        if (match) {
            return match[1].trim();
        }
        return null;
    }

    function parseStudyEventExtractAffix(name) {
        var normalized = name.replace(/\s+/g, " ").trim();
        var dayMatch = normalized.match(/\b(?:day|d)\s*-?\d+(?:\s*(?:-|to)\s*(?:day|d)?\s*-?\d+)?\s*/gi);
        if (!dayMatch || dayMatch.length === 0) {
            return { baseName: name, affix: null };
        }
        var lastDayMatch = dayMatch[dayMatch.length - 1];
        var lastIdx = normalized.toLowerCase().lastIndexOf(lastDayMatch.toLowerCase());
        var afterDay = normalized.substring(lastIdx + lastDayMatch.length).trim();
        afterDay = afterDay.replace(/^[\)\]]+/, "").trim();
        if (afterDay.length > 0 && !/^\d/.test(afterDay)) {
            var baseName = normalized.substring(0, lastIdx + lastDayMatch.length).trim();
            return { baseName: baseName, affix: afterDay };
        }
        return { baseName: name, affix: null };
    }

    function parseStudyEventGetBaseName(name, affixChecked) {
        if (!affixChecked) {
            return name;
        }
        var result = parseStudyEventExtractAffix(name);
        return result.baseName;
    }

    function parseStudyEventCollectFromTable() {
        var tbody = document.getElementById("sortableTable");
        if (!tbody) {
            log("[ParseStudyEvent] sortableTable not found");
            return [];
        }
        var rows = tbody.querySelectorAll("tr[id^='se_']");
        var events = [];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var rowId = row.id;
            if (rowId === "listTableNoRecordId") continue;
            var anchor = row.querySelector("td.dragHandle a");
            if (anchor) {
                var rawName = anchor.textContent || "";
                var name = rawName.replace(/\s+/g, " ").trim();
                if (name.length > 0) {
                    events.push({ name: name, originalName: rawName.trim() });
                    log("[ParseStudyEvent] Collected: \"" + name + "\"");
                }
            }
        }
        return events;
    }

    function parseStudyEventGenerateColumn1(events) {
        var lines = [];
        for (var i = 0; i < events.length; i++) {
            lines.push("\"" + events[i].name + "\"");
        }
        return lines.join(",\n");
    }

    function parseStudyEventFindScreeningEvent(events) {
        for (var i = 0; i < events.length; i++) {
            var lower = events[i].name.toLowerCase();
            if (lower.indexOf("screen") !== -1 || lower.indexOf("scrn") !== -1) {
                return events[i].name;
            }
        }
        return null;
    }

    function parseStudyEventFindSubjectEvent(events) {
        for (var i = 0; i < events.length; i++) {
            var lower = events[i].name.toLowerCase();
            if (lower.indexOf("subject") !== -1) {
                return events[i].name;
            }
        }
        return null;
    }

    function parseStudyEventGenerateColumn2(events, affixChecked, groupChecked) {
        var lines = [];
        var screeningEvent = parseStudyEventFindScreeningEvent(events);
        var subjectEvent = parseStudyEventFindSubjectEvent(events);
        var fallbackForNeg1 = screeningEvent || subjectEvent || null;

        var processedEvents = [];
        for (var i = 0; i < events.length; i++) {
            var ev = events[i];
            var dayInfo = parseStudyEventExtractDayNumber(ev.name);
            var groupLabel = groupChecked ? parseStudyEventExtractGroupLabel(ev.name) : null;
            var affixInfo = parseStudyEventExtractAffix(ev.name);
            var baseName = affixChecked ? affixInfo.baseName : ev.name;
            var hasAffix = affixChecked && affixInfo.affix !== null && affixInfo.affix.length > 0;
            processedEvents.push({
                name: ev.name,
                baseName: baseName,
                dayInfo: dayInfo,
                dayValue: dayInfo ? (dayInfo.type === "single" ? dayInfo.value : dayInfo.max) : null,
                groupLabel: groupLabel,
                hasAffix: hasAffix,
                index: i
            });
        }

        for (var i = 0; i < processedEvents.length; i++) {
            var current = processedEvents[i];

            if (parseStudyEventShouldIgnore(current.name)) {
                log("[ParseStudyEvent] Column2 skip (ignore keyword): \"" + current.name + "\"");
                continue;
            }

            if (current.dayInfo === null) {
                log("[ParseStudyEvent] Column2 skip (no day number): \"" + current.name + "\"");
                continue;
            }

            var mappedTo = null;
            var currentDay = current.dayValue;

            if (currentDay === -1) {
                mappedTo = fallbackForNeg1;
                log("[ParseStudyEvent] Column2 \"" + current.name + "\" mapped to Screening/Subject: \"" + mappedTo + "\"");
            } else {
                if (affixChecked && current.hasAffix) {
                    for (var j = i - 1; j >= 0; j--) {
                        var prev = processedEvents[j];
                        if (parseStudyEventShouldIgnore(prev.name)) continue;
                        if (groupChecked && prev.groupLabel !== current.groupLabel) continue;
                        if (!prev.hasAffix && prev.dayValue !== null) {
                            mappedTo = prev.name;
                            log("[ParseStudyEvent] Column2 \"" + current.name + "\" (affixed) mapped to non-affixed: \"" + mappedTo + "\"");
                            break;
                        }
                    }
                }

                if (!mappedTo) {
                    var targetDay = currentDay - 1;
                    var closestPrev = null;
                    var closestDiff = Infinity;

                    for (var j = i - 1; j >= 0; j--) {
                        var prev = processedEvents[j];
                        if (parseStudyEventShouldIgnore(prev.name)) continue;
                        if (groupChecked && prev.groupLabel !== current.groupLabel) continue;
                        if (prev.dayValue === null) continue;
                        if (affixChecked && prev.hasAffix) continue;

                        if (prev.dayValue === targetDay) {
                            mappedTo = prev.name;
                            log("[ParseStudyEvent] Column2 \"" + current.name + "\" mapped to N-1: \"" + mappedTo + "\"");
                            break;
                        }

                        if (prev.dayValue < currentDay) {
                            var diff = currentDay - prev.dayValue;
                            if (diff < closestDiff) {
                                closestDiff = diff;
                                closestPrev = prev;
                            }
                        }
                    }

                    if (!mappedTo && closestPrev) {
                        mappedTo = closestPrev.name;
                        log("[ParseStudyEvent] Column2 \"" + current.name + "\" mapped to closest previous: \"" + mappedTo + "\"");
                    }

                    if (!mappedTo && groupChecked && current.groupLabel) {
                        for (var j = i - 1; j >= 0; j--) {
                            var prev = processedEvents[j];
                            if (parseStudyEventShouldIgnore(prev.name)) continue;
                            if (prev.dayValue === null) continue;
                            if (affixChecked && prev.hasAffix) continue;
                            if (prev.dayValue < currentDay) {
                                mappedTo = prev.name;
                                log("[ParseStudyEvent] Column2 \"" + current.name + "\" (group fallback) mapped to: \"" + mappedTo + "\"");
                                break;
                            }
                        }
                    }
                }
            }

            if (mappedTo) {
                lines.push("\"" + current.name + "\" : \"" + mappedTo + "\"");
            }
        }

        return lines.join(",\n");
    }

    function parseStudyEventGenerateColumn3(events) {
        var lines = [];
        for (var i = 0; i < events.length; i++) {
            var ev = events[i];
            var dayInfo = parseStudyEventExtractDayNumber(ev.name);

            if (!dayInfo) {
                log("[ParseStudyEvent] Column3 skip (no day): \"" + ev.name + "\"");
                continue;
            }

            var outputValue;
            if (dayInfo.type === "range") {
                outputValue = dayInfo.mid;
                log("[ParseStudyEvent] Column3 \"" + ev.name + "\" range mid: " + outputValue);
            } else {
                outputValue = dayInfo.value - 1;
                log("[ParseStudyEvent] Column3 \"" + ev.name + "\" day-1: " + outputValue);
            }

            lines.push("\"" + ev.name + "\" : " + outputValue);
        }
        return lines.join(",\n");
    }

    function parseStudyEventCopyToClipboard(text) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(function() {
                log("[ParseStudyEvent] Copied to clipboard");
            }).catch(function(err) {
                log("[ParseStudyEvent] Clipboard error: " + err);
            });
        } else {
            var ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            try {
                document.execCommand("copy");
                log("[ParseStudyEvent] Copied to clipboard (fallback)");
            } catch (err) {
                log("[ParseStudyEvent] Copy failed: " + err);
            }
            document.body.removeChild(ta);
        }
    }

    function parseStudyEventCreateResultsUI(events, affixChecked, groupChecked) {
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "16px";
        container.style.height = "100%";

        var countDiv = document.createElement("div");
        countDiv.style.textAlign = "center";
        countDiv.style.fontSize = "16px";
        countDiv.style.fontWeight = "600";
        countDiv.style.color = "#9df";
        countDiv.style.padding = "8px";
        countDiv.style.background = "#1a1a1a";
        countDiv.style.borderRadius = "6px";
        countDiv.textContent = "Collected " + events.length + " events";
        container.appendChild(countDiv);

        var optionsDiv = document.createElement("div");
        optionsDiv.style.textAlign = "center";
        optionsDiv.style.fontSize = "12px";
        optionsDiv.style.color = "#888";
        optionsDiv.textContent = "Options: Affix=" + (affixChecked ? "Yes" : "No") + ", Group=" + (groupChecked ? "Yes" : "No");
        container.appendChild(optionsDiv);

        var columnsContainer = document.createElement("div");
        columnsContainer.style.display = "grid";
        columnsContainer.style.gridTemplateColumns = "1fr 1fr 1fr";
        columnsContainer.style.gap = "12px";
        columnsContainer.style.flex = "1";
        columnsContainer.style.minHeight = "0";
        columnsContainer.style.overflow = "hidden";

        var col1Data = parseStudyEventGenerateColumn1(events);
        var col2Data = parseStudyEventGenerateColumn2(events, affixChecked, groupChecked);
        var col3Data = parseStudyEventGenerateColumn3(events);

        function createColumn(title, data) {
            var col = document.createElement("div");
            col.style.display = "flex";
            col.style.flexDirection = "column";
            col.style.gap = "8px";
            col.style.minHeight = "0";
            col.style.overflow = "hidden";

            var header = document.createElement("div");
            header.style.display = "flex";
            header.style.justifyContent = "space-between";
            header.style.alignItems = "center";
            header.style.padding = "6px 8px";
            header.style.background = "#2a2a2a";
            header.style.borderRadius = "4px";

            var titleEl = document.createElement("span");
            titleEl.style.fontWeight = "600";
            titleEl.style.fontSize = "13px";
            titleEl.textContent = title;

            var copyBtn = document.createElement("button");
            copyBtn.textContent = "Copy";
            copyBtn.style.background = "#28a745";
            copyBtn.style.color = "#fff";
            copyBtn.style.border = "none";
            copyBtn.style.borderRadius = "4px";
            copyBtn.style.padding = "4px 10px";
            copyBtn.style.cursor = "pointer";
            copyBtn.style.fontSize = "12px";
            copyBtn.style.fontWeight = "500";
            copyBtn.addEventListener("mouseenter", function() { this.style.background = "#218838"; });
            copyBtn.addEventListener("mouseleave", function() { this.style.background = "#28a745"; });
            copyBtn.addEventListener("click", function() {
                parseStudyEventCopyToClipboard(data);
                var origText = copyBtn.textContent;
                copyBtn.textContent = "Copied!";
                copyBtn.style.background = "#17a2b8";
                setTimeout(function() {
                    copyBtn.textContent = origText;
                    copyBtn.style.background = "#28a745";
                }, 1500);
            });

            header.appendChild(titleEl);
            header.appendChild(copyBtn);
            col.appendChild(header);

            var content = document.createElement("div");
            content.style.flex = "1";
            content.style.overflow = "auto";
            content.style.background = "#1a1a1a";
            content.style.border = "1px solid #333";
            content.style.borderRadius = "4px";
            content.style.padding = "8px";
            content.style.fontSize = "12px";
            content.style.fontFamily = "monospace";
            content.style.whiteSpace = "pre-wrap";
            content.style.wordBreak = "break-word";
            content.style.color = "#e0e0e0";
            content.textContent = data;

            col.appendChild(content);
            return col;
        }

        columnsContainer.appendChild(createColumn("Study Event List", col1Data));
        columnsContainer.appendChild(createColumn("Previous Event Mapping", col2Data));
        columnsContainer.appendChild(createColumn("Day-1 Mapping", col3Data));

        container.appendChild(columnsContainer);

        return container;
    }

    function APS_ParseStudyEvent() {
        log("[ParseStudyEvent] Starting...");
        log("[ParseStudyEvent] URL check: " + location.href);

        if (!isStudyEventPage()) {
            log("[ParseStudyEvent] Not on study event page");
            showWrongPagePopup("Parse Study Event", "cenexel.clinspark.com/.../studyevent or cenexeltest.clinspark.com/.../studyevent", location.pathname);
            return;
        }

        log("[ParseStudyEvent] On valid study event page, showing options");

        var optionsContainer = document.createElement("div");
        optionsContainer.style.display = "flex";
        optionsContainer.style.flexDirection = "column";
        optionsContainer.style.gap = "20px";
        optionsContainer.style.padding = "12px";

        var checkboxesDiv = document.createElement("div");
        checkboxesDiv.style.display = "flex";
        checkboxesDiv.style.flexDirection = "column";
        checkboxesDiv.style.gap = "16px";

        function createCheckboxRow(labelText, description) {
            var row = document.createElement("label");
            row.style.display = "flex";
            row.style.alignItems = "flex-start";
            row.style.gap = "12px";
            row.style.cursor = "pointer";
            row.style.padding = "12px";
            row.style.background = "#1a1a1a";
            row.style.borderRadius = "8px";
            row.style.border = "1px solid #333";
            row.style.transition = "border-color 0.2s";
            row.addEventListener("mouseenter", function() { row.style.borderColor = "#555"; });
            row.addEventListener("mouseleave", function() { row.style.borderColor = "#333"; });

            var checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.style.width = "18px";
            checkbox.style.height = "18px";
            checkbox.style.marginTop = "2px";
            checkbox.style.cursor = "pointer";
            checkbox.style.accentColor = "#17a2b8";

            var textDiv = document.createElement("div");
            textDiv.style.display = "flex";
            textDiv.style.flexDirection = "column";
            textDiv.style.gap = "4px";

            var labelEl = document.createElement("span");
            labelEl.style.fontWeight = "600";
            labelEl.style.fontSize = "14px";
            labelEl.textContent = labelText;

            var descEl = document.createElement("span");
            descEl.style.fontSize = "12px";
            descEl.style.color = "#888";
            descEl.textContent = description;

            textDiv.appendChild(labelEl);
            textDiv.appendChild(descEl);
            row.appendChild(checkbox);
            row.appendChild(textDiv);

            return { row: row, checkbox: checkbox };
        }

        var affixCheckbox = createCheckboxRow("Affix", "Map affixed events (Predose, 1.5Hr, etc.) to previous non-affixed event");
        var groupCheckbox = createCheckboxRow("Group", "Map events within the same group label (e.g., (Sched B)) separately");

        checkboxesDiv.appendChild(affixCheckbox.row);
        checkboxesDiv.appendChild(groupCheckbox.row);
        optionsContainer.appendChild(checkboxesDiv);

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#17a2b8";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "none";
        confirmBtn.style.borderRadius = "8px";
        confirmBtn.style.padding = "14px 24px";
        confirmBtn.style.cursor = "pointer";
        confirmBtn.style.fontSize = "16px";
        confirmBtn.style.fontWeight = "600";
        confirmBtn.style.transition = "background 0.2s";
        confirmBtn.addEventListener("mouseenter", function() { confirmBtn.style.background = "#138496"; });
        confirmBtn.addEventListener("mouseleave", function() { confirmBtn.style.background = "#17a2b8"; });

        optionsContainer.appendChild(confirmBtn);

        var popup = createPopup({
            title: "Parse Study Event - Options",
            content: optionsContainer,
            width: "480px",
            height: "auto"
        });

        confirmBtn.addEventListener("click", async function() {
            var affixChecked = affixCheckbox.checkbox.checked;
            var groupChecked = groupCheckbox.checkbox.checked;

            log("[ParseStudyEvent] Options chosen: Affix=" + affixChecked + ", Group=" + groupChecked);

            var loadingContainer = document.createElement("div");
            loadingContainer.style.display = "flex";
            loadingContainer.style.flexDirection = "column";
            loadingContainer.style.alignItems = "center";
            loadingContainer.style.gap = "20px";
            loadingContainer.style.padding = "40px 20px";

            var spinner = document.createElement("div");
            spinner.style.width = "48px";
            spinner.style.height = "48px";
            spinner.style.border = "4px solid #333";
            spinner.style.borderTop = "4px solid #17a2b8";
            spinner.style.borderRadius = "50%";
            spinner.style.animation = "parseStudyEventSpin 1s linear infinite";

            var style = document.createElement("style");
            style.textContent = "@keyframes parseStudyEventSpin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }";
            document.head.appendChild(style);

            var statusText = document.createElement("div");
            statusText.style.fontSize = "16px";
            statusText.style.fontWeight = "500";
            statusText.style.color = "#9df";
            statusText.textContent = "Scanning table...";

            var progressText = document.createElement("div");
            progressText.style.fontSize = "13px";
            progressText.style.color = "#888";
            progressText.textContent = "Collecting study events from sortableTable";

            loadingContainer.appendChild(spinner);
            loadingContainer.appendChild(statusText);
            loadingContainer.appendChild(progressText);

            popup.setContent(loadingContainer);

            await sleep(300);

            log("[ParseStudyEvent] Searching for sortableTable");
            var events = parseStudyEventCollectFromTable();

            if (events.length === 0) {
                log("[ParseStudyEvent] No events found");
                statusText.textContent = "No study events found";
                statusText.style.color = "#f66";
                spinner.style.display = "none";
                progressText.textContent = "Make sure the table has data loaded";
                return;
            }

            log("[ParseStudyEvent] Found " + events.length + " events, generating output");
            statusText.textContent = "Processing " + events.length + " events...";

            await sleep(200);

            var resultsUI = parseStudyEventCreateResultsUI(events, affixChecked, groupChecked);
            popup.setContent(resultsUI);

            popup.element.style.width = "1000px";
            popup.element.style.height = "600px";
            popup.element.style.maxWidth = "95%";
            popup.element.style.maxHeight = "85%";

            var rect = popup.element.getBoundingClientRect();
            popup.element.style.left = (window.innerWidth / 2) + "px";
            popup.element.style.top = (window.innerHeight / 2) + "px";
            popup.element.style.transform = "translate(-50%, -50%)";

            log("[ParseStudyEvent] Results displayed, collected " + events.length + " events");
        });
    }


    //==========================
    // METHODS LIBRARY FEATURE
    //==========================

    // Methods Library Configuration
    var METHODS_INDEX_URL = "https://raw.githubusercontent.com/vctruong100/Automator/refs/heads/main/index.json";
    var METHODS_CACHE_KEY = "activityPlanState.methodsLibrary.cache";
    var METHODS_CACHE_EXPIRY_MS = 1000 * 60 * 60; // 1 hour
    var METHODS_PRELOAD_BODIES = false;
    var METHODS_BODY_CACHE = {};
    var METHODS_LIBRARY_MODAL_REF = null;
    var STORAGE_METHODS_FAVORITES = "activityPlanState.methodsLibrary.favorites";
    var STORAGE_METHODS_RECENTS = "activityPlanState.methodsLibrary.recents";
    var STORAGE_METHODS_PINS = "activityPlanState.methodsLibrary.pins";
    var STORAGE_METHODS_LAST_SEARCH = "activityPlanState.methodsLibrary.lastSearch";
    var STORAGE_METHODS_LAST_TAG = "activityPlanState.methodsLibrary.lastTag";
    var STORAGE_METHODS_LAST_METHOD = "activityPlanState.methodsLibrary.lastMethod";
    var STORAGE_METHODS_SORT_ORDER = "activityPlanState.methodsLibrary.sortOrder";
    var MAX_RECENTS = 10;
    var MAX_PINS = 5;
    var BOOTSTRAP_FOCUS_NS = "focusin.bs.modal";

    function disableBootstrapModalFocusTrap() {
        try {
            if (window.jQuery && window.jQuery(document)) {
                log("Methods Library: disabling Bootstrap focus trap");
                window.jQuery(document).off(BOOTSTRAP_FOCUS_NS);
            }
        } catch (e) {
            log("Methods Library: error disabling focus trap - " + String(e));
        }
    }

    function enableBootstrapModalFocusTrap() {
        try {
            // No direct re-bind because Bootstrap binds internally on show.
            // We just log here; when the next Bootstrap modal opens, it will re-bind its handler.
            log("Methods Library: focus trap will be restored by Bootstrap on next modal show");
        } catch (e) {
            log("Methods Library: error enabling focus trap - " + String(e));
        }
    }

    function neutralizeAjaxModal() {
        try {
            var ajaxModal = document.getElementById("ajaxModal");
            if (!ajaxModal) {
                log("Methods Library: #ajaxModal not found to neutralize");
                return;
            }

            ajaxModal.setAttribute("data-aps-neutralized", "1");
            ajaxModal.style.pointerEvents = "none";
            ajaxModal.style.visibility = "hidden";
            ajaxModal.style.opacity = "0";
            ajaxModal.setAttribute("aria-hidden", "true");

            var modalOpenOnBody = document.body.classList.contains("modal-open");
            if (modalOpenOnBody) {
                document.body.classList.remove("modal-open");
                log("Methods Library: removed modal-open from body");
            }

            log("Methods Library: #ajaxModal neutralized");
        } catch (e) {
            log("Methods Library: error neutralizing #ajaxModal - " + String(e));
        }
    }

    function restoreAjaxModal() {
        try {
            var ajaxModal = document.getElementById("ajaxModal");
            if (!ajaxModal) {
                log("Methods Library: #ajaxModal not found to restore");
                return;
            }

            var wasNeutralized = ajaxModal.getAttribute("data-aps-neutralized") === "1";
            if (!wasNeutralized) {
                log("Methods Library: #ajaxModal not previously neutralized; nothing to restore");
                return;
            }

            ajaxModal.removeAttribute("data-aps-neutralized");
            ajaxModal.style.pointerEvents = "";
            ajaxModal.style.visibility = "";
            ajaxModal.style.opacity = "";
            ajaxModal.removeAttribute("aria-hidden");

            // If the Bootstrap modal is still shown, body may need modal-open again
            var stillShown = ajaxModal.classList.contains("in") || ajaxModal.classList.contains("show");
            if (stillShown && !document.body.classList.contains("modal-open")) {
                document.body.classList.add("modal-open");
                log("Methods Library: restored modal-open on body");
            }

            log("Methods Library: #ajaxModal restored");
        } catch (e) {
            log("Methods Library: error restoring #ajaxModal - " + String(e));
        }
    }

    function getMethodsCache() {
        try {
            var raw = localStorage.getItem(METHODS_CACHE_KEY);
            if (!raw) return null;
            var cache = JSON.parse(raw);
            if (!cache || !cache.data || !cache.timestamp) return null;
            var age = Date.now() - cache.timestamp;
            if (age > METHODS_CACHE_EXPIRY_MS) {
                localStorage.removeItem(METHODS_CACHE_KEY);
                return null;
            }
            return cache.data;
        } catch (e) {
            return null;
        }
    }

    function setMethodsCache(data) {
        try {
            var cache = { data: data, timestamp: Date.now() };
            localStorage.setItem(METHODS_CACHE_KEY, JSON.stringify(cache));
        } catch (e) {
            log("Methods Library: cache write error - " + String(e));
        }
    }

    async function fetchMethodsIndex(forceRefresh) {
        if (!forceRefresh) {
            var cached = getMethodsCache();
            if (cached) {
                log("Methods Library: using cached index (" + cached.length + " methods)");
                return { success: true, data: cached };
            }
        }
        log("Methods Library: fetching index from " + METHODS_INDEX_URL);
        try {
            var response = await fetch(METHODS_INDEX_URL, { cache: forceRefresh ? "no-store" : "default" });
            if (!response.ok) {
                return { success: false, error: "HTTP " + response.status + ": " + response.statusText };
            }
            var data = await response.json();
            if (!Array.isArray(data)) {
                return { success: false, error: "Invalid index format (expected array)" };
            }
            setMethodsCache(data);
            log("Methods Library: fetched " + data.length + " methods");
            return { success: true, data: data };
        } catch (e) {
            log("Methods Library: fetch error - " + String(e));
            return { success: false, error: String(e) };
        }
    }

    async function fetchMethodBody(url) {
        if (METHODS_BODY_CACHE[url]) {
            return { success: true, body: METHODS_BODY_CACHE[url] };
        }

        // Convert relative or GitHub blob URLs to raw.githubusercontent.com format
        var fetchUrl = url;
        if (url.indexOf("github.com") !== -1 && url.indexOf("raw.githubusercontent.com") === -1) {
            // Convert github.com/user/repo/blob/branch/path to raw.githubusercontent.com/user/repo/branch/path
            fetchUrl = url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/");
        } else if (url.indexOf("http") !== 0) {
            // Handle relative paths - prepend the base GitHub raw URL
            var baseUrl = "https://raw.githubusercontent.com/vctruong100/Automator/main/";
            fetchUrl = baseUrl + url.replace(/^\/+/, "");
        }

        try {
            var response = await fetch(fetchUrl);
            if (!response.ok) {
                return { success: false, error: "HTTP " + response.status + " for " + fetchUrl };
            }
            var text = await response.text();
            METHODS_BODY_CACHE[url] = text;
            return { success: true, body: text };
        } catch (e) {
            return { success: false, error: String(e) };
        }
    }

    function buildTagList(methods) {
        var tagSet = {};
        for (var i = 0; i < methods.length; i++) {
            var tags = methods[i].tags || [];
            for (var j = 0; j < tags.length; j++) {
                tagSet[tags[j]] = true;
            }
        }
        return Object.keys(tagSet).sort();
    }

    function scoreMethod(method, query, searchBody, bodyText) {
        if (!query || query.trim().length === 0) return 100;
        var q = query.toLowerCase().trim();
        var score = 0;
        var title = (method.title || "").toLowerCase();
        var id = (method.id || "").toLowerCase();
        var tags = (method.tags || []).map(function(t) { return t.toLowerCase(); });

        if (title.indexOf(q) !== -1) {
            score += 50;
            if (title.indexOf(q) === 0) score += 20;
        }
        if (id.indexOf(q) !== -1) {
            score += 30;
            if (id === q) score += 20;
        }
        for (var i = 0; i < tags.length; i++) {
            if (tags[i].indexOf(q) !== -1) {
                score += 20;
                break;
            }
        }
        if (searchBody && bodyText) {
            var body = bodyText.toLowerCase();
            if (body.indexOf(q) !== -1) {
                score += 10;
            }
        }
        return score;
    }

    function filterAndSortMethods(methods, query, tagFilter, searchBody, bodiesMap, sortOrder, favorites, pins) {
        var results = [];
        for (var i = 0; i < methods.length; i++) {
            var m = methods[i];
            if (tagFilter && tagFilter !== "all") {
                var tags = m.tags || [];
                var hasTag = false;
                for (var j = 0; j < tags.length; j++) {
                    if (tags[j] === tagFilter) {
                        hasTag = true;
                        break;
                    }
                }
                if (!hasTag) continue;
            }
            var bodyText = bodiesMap[m.url] || "";
            var score = scoreMethod(m, query, searchBody, bodyText);
            if (query && query.trim().length > 0 && score === 0) continue;
            results.push({ method: m, score: score });
        }

        results.sort(function(a, b) {
            // Pinned items always come first
            var isPinA = pins.indexOf(a.method.id) !== -1;
            var isPinB = pins.indexOf(b.method.id) !== -1;
            if (isPinA && !isPinB) return -1;
            if (!isPinA && isPinB) return 1;

            // Then sort by selected order
            if (sortOrder === "title") {
                var titleA = (a.method.title || "").toLowerCase();
                var titleB = (b.method.title || "").toLowerCase();
                if (titleA < titleB) return -1;
                if (titleA > titleB) return 1;
                return 0;
            } else if (sortOrder === "updated") {
                var dateA = a.method.updated || "";
                var dateB = b.method.updated || "";
                if (dateB < dateA) return -1;
                if (dateB > dateA) return 1;
                return 0;
            } else {
                // Relevance (default)
                if (b.score !== a.score) return b.score - a.score;
                var titleA = (a.method.title || "").toLowerCase();
                var titleB = (b.method.title || "").toLowerCase();
                if (titleA < titleB) return -1;
                if (titleA > titleB) return 1;
                return 0;
            }
        });
        return results.map(function(r) { return r.method; });
    }

    function openMethodsLibraryModal() {
        var bootstrapBackdrops = document.querySelectorAll('.modal-backdrop');
        for (var i = 0; i < bootstrapBackdrops.length; i++) {
            bootstrapBackdrops[i].style.display = 'none';
        }

        var openBootstrapModals = document.querySelectorAll('.modal.in, .modal.show');

        // >>> NEW: disable Bootstrap focus trap and neutralize #ajaxModal if present
        disableBootstrapModalFocusTrap();
        neutralizeAjaxModal();
        // <<< NEW

        if (METHODS_LIBRARY_MODAL_REF && document.body.contains(METHODS_LIBRARY_MODAL_REF)) {
            var existingSearchInput = METHODS_LIBRARY_MODAL_REF.querySelector('input[type="text"]');
            if (existingSearchInput) {
                existingSearchInput.disabled = false;
                existingSearchInput.readOnly = false;
                existingSearchInput.focus();
            }
            METHODS_LIBRARY_MODAL_REF.focus();
            return;
        }

        var allMethods = [];
        var allTags = [];
        var selectedMethod = null;
        var currentBodyText = "";

        var overlay = document.createElement("div");
        overlay.style.position = "fixed";
        overlay.style.top = "0";
        overlay.style.left = "0";
        overlay.style.width = "100%";
        overlay.style.height = "100%";
        overlay.style.background = "rgba(0,0,0,0.6)";
        overlay.style.zIndex = "999997";

        var modal = document.createElement("div");
        modal.setAttribute("role", "dialog");
        modal.setAttribute("aria-modal", "true");
        modal.setAttribute("aria-label", "ClinSpark Methods Library");
        modal.tabIndex = -1;
        modal.style.position = "fixed";
        modal.style.top = "50%";
        modal.style.left = "50%";
        modal.style.transform = "translate(-50%, -50%)";
        modal.style.width = "900px";
        modal.style.maxWidth = "95vw";
        modal.style.height = "600px";
        modal.style.maxHeight = "90vh";
        modal.style.background = "#1a1a1a";
        modal.style.color = "#fff";
        modal.style.border = "1px solid #444";
        modal.style.borderRadius = "10px";
        modal.style.boxShadow = "0 8px 32px rgba(0,0,0,0.5)";
        modal.style.display = "flex";
        modal.style.flexDirection = "column";
        modal.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        modal.style.fontSize = "14px";
        modal.style.zIndex = "999998";
        modal.style.outline = "none";
        METHODS_LIBRARY_MODAL_REF = modal;

        var header = document.createElement("div");
        header.style.display = "flex";
        header.style.alignItems = "center";
        header.style.justifyContent = "space-between";
        header.style.padding = "12px 16px";
        header.style.borderBottom = "1px solid #333";
        header.style.cursor = "move";
        header.style.userSelect = "none";

        var titleEl = document.createElement("div");
        titleEl.textContent = "ClinSpark Methods Library";
        titleEl.style.fontWeight = "600";
        titleEl.style.fontSize = "16px";
        header.appendChild(titleEl);

        var closeBtn = document.createElement("button");
        closeBtn.textContent = "✕";
        closeBtn.setAttribute("aria-label", "Close");
        closeBtn.style.background = "transparent";
        closeBtn.style.color = "#fff";
        closeBtn.style.border = "none";
        closeBtn.style.fontSize = "18px";
        closeBtn.style.cursor = "pointer";
        closeBtn.style.padding = "4px 8px";
        closeBtn.style.borderRadius = "4px";
        closeBtn.onmouseenter = function() { closeBtn.style.background = "#333"; };
        closeBtn.onmouseleave = function() { closeBtn.style.background = "transparent"; };
        header.appendChild(closeBtn);
        modal.appendChild(header);

        var toolbar = document.createElement("div");
        toolbar.style.display = "flex";
        toolbar.style.alignItems = "center";
        toolbar.style.gap = "10px";
        toolbar.style.padding = "10px 16px";
        toolbar.style.borderBottom = "1px solid #333";
        toolbar.style.flexWrap = "wrap";

        var searchInput = document.createElement("input");
        searchInput.type = "text";
        searchInput.placeholder = "Search methods...";
        searchInput.setAttribute("aria-label", "Search methods");
        searchInput.style.flex = "1";
        searchInput.style.minWidth = "150px";
        searchInput.style.padding = "8px 12px";
        searchInput.style.background = "#2a2a2a";
        searchInput.style.border = "1px solid #444";
        searchInput.style.borderRadius = "6px";
        searchInput.style.color = "#fff";
        searchInput.style.fontSize = "14px";
        searchInput.style.outline = "none";
        searchInput.onfocus = function() { searchInput.style.borderColor = "#5b43c7"; };
        searchInput.onblur = function() { searchInput.style.borderColor = "#444"; };
        toolbar.appendChild(searchInput);

        var tagSelect = document.createElement("select");
        tagSelect.setAttribute("aria-label", "Filter by tag");
        tagSelect.style.padding = "8px 12px";
        tagSelect.style.background = "#2a2a2a";
        tagSelect.style.border = "1px solid #444";
        tagSelect.style.borderRadius = "6px";
        tagSelect.style.color = "#fff";
        tagSelect.style.fontSize = "14px";
        tagSelect.style.cursor = "pointer";
        toolbar.appendChild(tagSelect);

        var sortSelect = document.createElement("select");
        sortSelect.setAttribute("aria-label", "Sort order");
        sortSelect.style.padding = "8px 12px";
        sortSelect.style.background = "#2a2a2a";
        sortSelect.style.border = "1px solid #444";
        sortSelect.style.borderRadius = "6px";
        sortSelect.style.color = "#fff";
        sortSelect.style.fontSize = "14px";
        sortSelect.style.cursor = "pointer";

        var sortOpts = [
            { value: "relevance", text: "Sort: Relevance" },
            { value: "title", text: "Sort: Title" },
            { value: "updated", text: "Sort: Updated" }
        ];
        for (var i = 0; i < sortOpts.length; i++) {
            var opt = document.createElement("option");
            opt.value = sortOpts[i].value;
            opt.textContent = sortOpts[i].text;
            sortSelect.appendChild(opt);
        }
        toolbar.appendChild(sortSelect);

        var searchBodyLabel = document.createElement("label");
        searchBodyLabel.style.display = "flex";
        searchBodyLabel.style.alignItems = "center";
        searchBodyLabel.style.gap = "6px";
        searchBodyLabel.style.cursor = "pointer";
        searchBodyLabel.style.fontSize = "13px";
        searchBodyLabel.style.color = "#aaa";
        var searchBodyCheckbox = document.createElement("input");
        searchBodyCheckbox.type = "checkbox";
        searchBodyCheckbox.setAttribute("aria-label", "Search in body content");
        searchBodyLabel.appendChild(searchBodyCheckbox);
        searchBodyLabel.appendChild(document.createTextNode("Search body"));
        toolbar.appendChild(searchBodyLabel);

        var favoritesLabel = document.createElement("label");
        favoritesLabel.style.display = "flex";
        favoritesLabel.style.alignItems = "center";
        favoritesLabel.style.gap = "6px";
        favoritesLabel.style.cursor = "pointer";
        favoritesLabel.style.fontSize = "13px";
        favoritesLabel.style.color = "#aaa";
        var favoritesCheckbox = document.createElement("input");
        favoritesCheckbox.type = "checkbox";
        favoritesCheckbox.setAttribute("aria-label", "Show only favorites");
        favoritesLabel.appendChild(favoritesCheckbox);
        favoritesLabel.appendChild(document.createTextNode("★ Favorites"));
        toolbar.appendChild(favoritesLabel);

        var toolbarBtns = document.createElement("div");
        toolbarBtns.style.display = "flex";
        toolbarBtns.style.gap = "8px";

        function createToolbarBtn(text, ariaLabel) {
            var btn = document.createElement("button");
            btn.textContent = text;
            btn.setAttribute("aria-label", ariaLabel);
            btn.style.padding = "8px 14px";
            btn.style.background = "#333";
            btn.style.color = "#fff";
            btn.style.border = "1px solid #444";
            btn.style.borderRadius = "6px";
            btn.style.cursor = "pointer";
            btn.style.fontSize = "13px";
            btn.style.fontWeight = "500";
            btn.style.transition = "background 0.15s";
            btn.onmouseenter = function() { btn.style.background = "#444"; };
            btn.onmouseleave = function() { btn.style.background = "#333"; };
            return btn;
        }

        var copyBtn = createToolbarBtn("Copy", "Copy method code to clipboard");
        var refreshBtn = createToolbarBtn("Refresh", "Refresh methods index");
        toolbarBtns.appendChild(copyBtn);
        toolbarBtns.appendChild(refreshBtn);
        toolbar.appendChild(toolbarBtns);

        modal.appendChild(toolbar);

        var mainContent = document.createElement("div");
        mainContent.style.display = "flex";
        mainContent.style.flex = "1";
        mainContent.style.overflow = "hidden";

        var listPane = document.createElement("div");
        listPane.style.width = "280px";
        listPane.style.minWidth = "200px";
        listPane.style.borderRight = "1px solid #333";
        listPane.style.overflowY = "auto";
        listPane.style.background = "#1a1a1a";
        listPane.setAttribute("role", "listbox");
        listPane.setAttribute("aria-label", "Methods list");

        var previewPane = document.createElement("div");
        previewPane.style.flex = "1";
        previewPane.style.display = "flex";
        previewPane.style.flexDirection = "column";
        previewPane.style.overflow = "hidden";
        previewPane.style.background = "#111";

        var previewHeader = document.createElement("div");
        previewHeader.style.padding = "12px 16px";
        previewHeader.style.borderBottom = "1px solid #333";
        previewHeader.style.background = "#1a1a1a";

        var previewTitle = document.createElement("div");
        previewTitle.style.fontWeight = "600";
        previewTitle.style.fontSize = "15px";
        previewTitle.style.marginBottom = "4px";
        previewTitle.textContent = "Select a method";
        previewHeader.appendChild(previewTitle);

        var previewMeta = document.createElement("div");
        previewMeta.style.fontSize = "12px";
        previewMeta.style.color = "#888";
        previewHeader.appendChild(previewMeta);

        var previewBody = document.createElement("pre");
        previewBody.style.flex = "1";
        previewBody.style.margin = "0";
        previewBody.style.padding = "16px";
        previewBody.style.overflowY = "auto";
        previewBody.style.background = "#0d0d0d";
        previewBody.style.fontSize = "13px";
        previewBody.style.fontFamily = "Consolas, Monaco, 'Courier New', monospace";
        previewBody.style.whiteSpace = "pre-wrap";
        previewBody.style.wordBreak = "break-word";
        previewBody.style.color = "#e0e0e0";
        previewBody.style.lineHeight = "1.5";

        previewPane.appendChild(previewHeader);
        previewPane.appendChild(previewBody);
        mainContent.appendChild(listPane);
        mainContent.appendChild(previewPane);
        modal.appendChild(mainContent);

        var statusBar = document.createElement("div");
        statusBar.style.padding = "8px 16px";
        statusBar.style.borderTop = "1px solid #333";
        statusBar.style.fontSize = "12px";
        statusBar.style.color = "#888";
        statusBar.style.background = "#1a1a1a";
        statusBar.textContent = "Loading...";
        modal.appendChild(statusBar);

        function updateStatus(msg, isError) {
            statusBar.textContent = msg;
            statusBar.style.color = isError ? "#e74c3c" : "#888";
        }

        function populateTagSelect() {
            tagSelect.innerHTML = "";
            var allOpt = document.createElement("option");
            allOpt.value = "all";
            allOpt.textContent = "All tags";
            tagSelect.appendChild(allOpt);
            for (var i = 0; i < allTags.length; i++) {
                var opt = document.createElement("option");
                opt.value = allTags[i];
                opt.textContent = allTags[i];
                tagSelect.appendChild(opt);
            }
        }

        function renderList(methods) {
            listPane.innerHTML = "";
            if (methods.length === 0) {
                var empty = document.createElement("div");
                empty.style.padding = "20px";
                empty.style.color = "#666";
                empty.style.textAlign = "center";
                empty.textContent = "No methods found";
                listPane.appendChild(empty);
                return;
            }

            var favorites = getFavorites();
            var recents = getRecents();
            var pins = getPins();

            for (var i = 0; i < methods.length; i++) {
                (function(m, idx) {
                    var item = document.createElement("div");
                    item.setAttribute("role", "option");
                    item.setAttribute("aria-selected", "false");
                    item.setAttribute("data-method-id", m.id);
                    item.tabIndex = 0;
                    item.style.padding = "10px 14px";
                    item.style.borderBottom = "1px solid #2a2a2a";
                    item.style.cursor = "pointer";
                    item.style.transition = "background 0.1s";
                    item.style.position = "relative";

                    var itemId = document.createElement("div");
                    itemId.style.fontSize = "11px";
                    itemId.style.color = "#5b43c7";
                    itemId.style.marginBottom = "2px";
                    itemId.textContent = m.id || "";
                    item.appendChild(itemId);

                    var itemTitle = document.createElement("div");
                    itemTitle.style.fontWeight = "500";
                    itemTitle.style.fontSize = "13px";
                    itemTitle.textContent = m.title || "(Untitled)";
                    item.appendChild(itemTitle);

                    if (m.tags && m.tags.length > 0) {
                        var itemTags = document.createElement("div");
                        itemTags.style.fontSize = "11px";
                        itemTags.style.color = "#666";
                        itemTags.style.marginTop = "4px";
                        itemTags.textContent = m.tags.join(", ");
                        item.appendChild(itemTags);
                    }

                    var badges = document.createElement("div");
                    badges.style.display = "flex";
                    badges.style.gap = "4px";
                    badges.style.marginTop = "4px";
                    badges.style.fontSize = "10px";

                    var isFav = favorites.indexOf(m.id) !== -1;
                    var isRecent = recents.indexOf(m.id) !== -1;
                    var isPinned = pins.indexOf(m.id) !== -1;

                    if (isPinned) {
                        var pinBadge = document.createElement("span");
                        pinBadge.textContent = "📌 Pinned";
                        pinBadge.style.color = "#f39c12";
                        badges.appendChild(pinBadge);
                    }
                    if (isFav) {
                        var favBadge = document.createElement("span");
                        favBadge.textContent = "★ Favorite";
                        favBadge.style.color = "#ffd700";
                        badges.appendChild(favBadge);
                    }
                    if (isRecent && !isPinned) {
                        var recentBadge = document.createElement("span");
                        recentBadge.textContent = "🕒 Recent";
                        recentBadge.style.color = "#888";
                        badges.appendChild(recentBadge);
                    }
                    if (badges.children.length > 0) {
                        item.appendChild(badges);
                    }

                    var actions = document.createElement("div");
                    actions.style.position = "absolute";
                    actions.style.top = "8px";
                    actions.style.right = "8px";
                    actions.style.display = "none";
                    actions.style.gap = "4px";

                    var favBtn = document.createElement("button");
                    favBtn.textContent = isFav ? "★" : "☆";
                    favBtn.title = isFav ? "Remove from favorites" : "Add to favorites";
                    favBtn.style.background = "transparent";
                    favBtn.style.border = "none";
                    favBtn.style.color = isFav ? "#ffd700" : "#888";
                    favBtn.style.fontSize = "16px";
                    favBtn.style.cursor = "pointer";
                    favBtn.style.padding = "2px 6px";
                    favBtn.onclick = function(e) {
                        e.stopPropagation();
                        toggleFavorite(m.id);
                        doSearch();
                    };
                    actions.appendChild(favBtn);

                    var pinBtn = document.createElement("button");
                    pinBtn.textContent = isPinned ? "📌" : "📍";
                    pinBtn.title = isPinned ? "Unpin" : "Pin to top";
                    pinBtn.style.background = "transparent";
                    pinBtn.style.border = "none";
                    pinBtn.style.color = isPinned ? "#f39c12" : "#888";
                    pinBtn.style.fontSize = "14px";
                    pinBtn.style.cursor = "pointer";
                    pinBtn.style.padding = "2px 6px";
                    pinBtn.onclick = function(e) {
                        e.stopPropagation();
                        var result = togglePin(m.id);
                        if (!result.success) {
                            updateStatus(result.message, true);
                            setTimeout(function() { doSearch(); }, 2000);
                        } else {
                            doSearch();
                        }
                    };
                    actions.appendChild(pinBtn);

                    item.appendChild(actions);

                    item.onmouseenter = function() {
                        if (selectedMethod !== m) {
                            item.style.background = "#252525";
                        }
                        actions.style.display = "flex";
                    };
                    item.onmouseleave = function() {
                        if (selectedMethod !== m) {
                            item.style.background = "transparent";
                        }
                        actions.style.display = "none";
                    };

                    item.onclick = function() { selectMethod(m, item); };
                    item.onkeydown = function(e) {
                        if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            selectMethod(m, item);
                        }
                    };

                    listPane.appendChild(item);
                })(methods[i], i);
            }
        }

        async function selectMethod(m, itemEl) {
            selectedMethod = m;
            currentBodyText = "";
            addRecent(m.id);
            saveLastMethod(m.id);

            var items = listPane.querySelectorAll("[role='option']");
            for (var i = 0; i < items.length; i++) {
                items[i].setAttribute("aria-selected", "false");
                items[i].style.background = "transparent";
            }

            if (itemEl) {
                itemEl.setAttribute("aria-selected", "true");
                itemEl.style.background = "#2a2a2a";
                itemEl.scrollIntoView({ block: "nearest", behavior: "smooth" });
            }

            previewTitle.textContent = (m.id || "") + " — " + (m.title || "(Untitled)");

            var metaParts = [];
            if (m.tags && m.tags.length > 0) {
                metaParts.push("Tags: " + m.tags.join(", "));
            }
            if (m.updated) {
                metaParts.push("Updated: " + m.updated);
            }
            previewMeta.textContent = metaParts.join(" • ");
            previewBody.textContent = "Loading...";

            if (m.url) {
                var result = await fetchMethodBody(m.url);
                if (result.success) {
                    currentBodyText = result.body;
                    previewBody.textContent = result.body;
                } else {
                    previewBody.textContent = "Error loading method: " + result.error;
                }
            } else {
                previewBody.textContent = "(No URL specified)";
            }
        }

        function doSearch() {
            var query = searchInput.value;
            var tagFilter = tagSelect.value;
            var searchBody = searchBodyCheckbox.checked;
            var sortOrder = sortSelect.value;
            var showOnlyFavorites = favoritesCheckbox.checked;

            saveLastSearch(query);
            saveLastTag(tagFilter);
            saveSortOrder(sortOrder);

            var favorites = getFavorites();
            var pins = getPins();
            var methodsToFilter = allMethods;

            if (showOnlyFavorites) {
                methodsToFilter = allMethods.filter(function(m) {
                    return favorites.indexOf(m.id) !== -1;
                });
            }

            var filtered = filterAndSortMethods(methodsToFilter, query, tagFilter, searchBody, METHODS_BODY_CACHE, sortOrder, favorites, pins);
            renderList(filtered);
            updateStatus(filtered.length + " of " + allMethods.length + " methods");
        }

        async function loadIndex(forceRefresh) {
            updateStatus("Loading methods index...");
            var result = await fetchMethodsIndex(forceRefresh);
            if (!result.success) {
                updateStatus("Error: " + result.error, true);
                renderList([]);
                return;
            }

            allMethods = result.data;
            allTags = buildTagList(allMethods);
            populateTagSelect();

            if (METHODS_PRELOAD_BODIES) {
                updateStatus("Preloading method bodies...");
                for (var i = 0; i < allMethods.length; i++) {
                    if (allMethods[i].url) {
                        await fetchMethodBody(allMethods[i].url);
                    }
                }
            }

            searchInput.value = getLastSearch();
            tagSelect.value = getLastTag();
            sortSelect.value = getSortOrder();

            searchInput.disabled = false;
            searchInput.readOnly = false;
            searchInput.style.pointerEvents = 'auto';
            searchInput.tabIndex = 0;

            setTimeout(function() {
                var bootstrapBackdrops = document.querySelectorAll('.modal-backdrop');
                for (var i = 0; i < bootstrapBackdrops.length; i++) {
                    bootstrapBackdrops[i].style.display = 'none';
                }
                searchInput.focus();
                searchInput.select();
                searchInput.click();
            }, 150);

            doSearch();
        }

        searchInput.oninput = doSearch;
        tagSelect.onchange = doSearch;
        searchBodyCheckbox.onchange = doSearch;
        sortSelect.onchange = doSearch;
        favoritesCheckbox.onchange = doSearch;

        refreshBtn.onclick = function() {
            METHODS_BODY_CACHE = {};
            loadIndex(true);
        };

        copyBtn.onclick = async function() {
            if (!currentBodyText) {
                updateStatus("No method selected to copy", true);
                return;
            }
            try {
                await navigator.clipboard.writeText(currentBodyText);
                updateStatus("Copied to clipboard!");
                setTimeout(function() {
                    doSearch();
                }, 1500);
            } catch (e) {
                updateStatus("Copy failed: " + String(e), true);
            }
        };

        function closeModal() {
            overlay.remove();
            modal.remove();
            METHODS_LIBRARY_MODAL_REF = null;
            document.removeEventListener("keydown", keyHandler);

            // >>> NEW: restore neutralized modal and focus trap, then show backdrops again
            restoreAjaxModal();
            enableBootstrapModalFocusTrap();
            // <<< NEW

            var bootstrapBackdrops = document.querySelectorAll('.modal-backdrop');
            for (var i = 0; i < bootstrapBackdrops.length; i++) {
                bootstrapBackdrops[i].style.display = 'block';
            }
        }

        closeBtn.onclick = closeModal;
        overlay.onclick = closeModal;

        function keyHandler(e) {
            if (e.key === "Escape") {
                closeModal();
                return;
            }

            if (e.key === "Enter" && document.activeElement === searchInput) {
                var firstItem = listPane.querySelector("[role='option']");
                if (firstItem) {
                    var methodId = firstItem.getAttribute("data-method-id");
                    var method = allMethods.find(function(m) { return m.id === methodId; });
                    if (method) {
                        selectMethod(method, firstItem);
                    }
                }
                return;
            }

            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
                var items = Array.from(listPane.querySelectorAll("[role='option']"));
                if (items.length === 0) {
                    return;
                }

                var currentIndex = -1;
                for (var i = 0; i < items.length; i++) {
                    if (items[i].getAttribute("aria-selected") === "true") {
                        currentIndex = i;
                        break;
                    }
                }

                var nextIndex;
                if (e.key === "ArrowDown") {
                    if (currentIndex < items.length - 1) {
                        nextIndex = currentIndex + 1;
                    } else {
                        nextIndex = 0;
                    }
                } else {
                    if (currentIndex > 0) {
                        nextIndex = currentIndex - 1;
                    } else {
                        nextIndex = items.length - 1;
                    }
                }

                var nextItem = items[nextIndex];
                var methodId = nextItem.getAttribute("data-method-id");
                var method = allMethods.find(function(m) { return m.id === methodId; });
                if (method) {
                    selectMethod(method, nextItem);
                }
                e.preventDefault();
                return;
            }

            if (e.key === "Tab") {
                var focusable = modal.querySelectorAll('button, input, select, [tabindex]:not([tabindex="-1"])');
                if (focusable.length === 0) {
                    return;
                }
                var first = focusable[0];
                var last = focusable[focusable.length - 1];
                if (e.shiftKey && document.activeElement === first) {
                    e.preventDefault();
                    last.focus();
                } else if (!e.shiftKey && document.activeElement === last) {
                    e.preventDefault();
                    first.focus();
                }
            }
        }

        document.addEventListener("keydown", keyHandler);

        var isDragging = false;
        var dragStartX = 0, dragStartY = 0, modalStartX = 0, modalStartY = 0;

        header.onmousedown = function(e) {
            if (e.target === closeBtn) {
                return;
            }
            isDragging = true;
            var rect = modal.getBoundingClientRect();
            dragStartX = e.clientX;
            dragStartY = e.clientY;
            modalStartX = rect.left;
            modalStartY = rect.top;
            modal.style.transform = "none";
            e.preventDefault();
        };

        document.addEventListener("mousemove", function(e) {
            if (!isDragging) {
                return;
            }
            var dx = e.clientX - dragStartX;
            var dy = e.clientY - dragStartY;
            modal.style.left = (modalStartX + dx) + "px";
            modal.style.top = (modalStartY + dy) + "px";
        });

        document.addEventListener("mouseup", function() {
            isDragging = false;
        });

        document.body.appendChild(overlay);
        document.body.appendChild(modal);
        modal.focus();
        searchInput.focus();

        loadIndex(false);
        log("Methods Library: modal opened");
    }

    function getFavorites() {
        try {
            var raw = localStorage.getItem(STORAGE_METHODS_FAVORITES);
            if (!raw) return [];
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function saveFavorites(favs) {
        try {
            localStorage.setItem(STORAGE_METHODS_FAVORITES, JSON.stringify(favs));
        } catch (e) {}
    }

    function toggleFavorite(methodId) {
        var favs = getFavorites();
        var idx = favs.indexOf(methodId);
        if (idx !== -1) {
            favs.splice(idx, 1);
        } else {
            favs.push(methodId);
        }
        saveFavorites(favs);
        return favs;
    }

    function getRecents() {
        try {
            var raw = localStorage.getItem(STORAGE_METHODS_RECENTS);
            if (!raw) return [];
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function addRecent(methodId) {
        var recents = getRecents();
        var idx = recents.indexOf(methodId);
        if (idx !== -1) {
            recents.splice(idx, 1);
        }
        recents.unshift(methodId);
        if (recents.length > MAX_RECENTS) {
            recents = recents.slice(0, MAX_RECENTS);
        }
        try {
            localStorage.setItem(STORAGE_METHODS_RECENTS, JSON.stringify(recents));
        } catch (e) {}
    }

    function getPins() {
        try {
            var raw = localStorage.getItem(STORAGE_METHODS_PINS);
            if (!raw) return [];
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function savePins(pins) {
        try {
            localStorage.setItem(STORAGE_METHODS_PINS, JSON.stringify(pins));
        } catch (e) {}
    }

    function togglePin(methodId) {
        var pins = getPins();
        var idx = pins.indexOf(methodId);
        if (idx !== -1) {
            pins.splice(idx, 1);
        } else {
            if (pins.length >= MAX_PINS) {
                return { success: false, message: "Maximum " + MAX_PINS + " pins allowed" };
            }
            pins.push(methodId);
        }
        savePins(pins);
        return { success: true, pins: pins };
    }

    function getLastSearch() {
        try {
            return localStorage.getItem(STORAGE_METHODS_LAST_SEARCH) || "";
        } catch (e) {
            return "";
        }
    }

    function saveLastSearch(query) {
        try {
            localStorage.setItem(STORAGE_METHODS_LAST_SEARCH, query);
        } catch (e) {}
    }

    function getLastTag() {
        try {
            return localStorage.getItem(STORAGE_METHODS_LAST_TAG) || "all";
        } catch (e) {
            return "all";
        }
    }

    function saveLastTag(tag) {
        try {
            localStorage.setItem(STORAGE_METHODS_LAST_TAG, tag);
        } catch (e) {}
    }

    function getLastMethod() {
        try {
            return localStorage.getItem(STORAGE_METHODS_LAST_METHOD) || "";
        } catch (e) {
            return "";
        }
    }

    function saveLastMethod(methodId) {
        try {
            localStorage.setItem(STORAGE_METHODS_LAST_METHOD, methodId);
        } catch (e) {}
    }

    function getSortOrder() {
        try {
            return localStorage.getItem(STORAGE_METHODS_SORT_ORDER) || "relevance";
        } catch (e) {
            return "relevance";
        }
    }

    function saveSortOrder(order) {
        try {
            localStorage.setItem(STORAGE_METHODS_SORT_ORDER, order);
        } catch (e) {}
    }

    //========================================
    // UI Scaling
    //========================================
    function clearLogs() {
        try {
            localStorage.removeItem("activityPlanState.logs");
            log("ClearLogs: localStorage logs removed");
        } catch (e) {
            log("ClearLogs: error removing logs from storage " + String(e));
        }

        var box = document.getElementById(LOG_ID);
        var hadBox = !!box;
        log("ClearLogs: log box exists=" + String(hadBox));
        if (box) {
            var removed = 0;
            while (box.firstChild) {
                box.removeChild(box.firstChild);
                removed = removed + 1;
            }
        }
    }

    function scale(value) {
        if (typeof value === 'string' && value.endsWith('px')) {
            return (parseFloat(value) * UI_SCALE) + 'px';
        }
        return (value * UI_SCALE) + 'px';
    }

    function getStoredUIScale() {
        try {
            var stored = localStorage.getItem(STORAGE_UI_SCALE);
            if (stored) {
                var scale = parseFloat(stored);
                if (!isNaN(scale) && scale >= 0.5 && scale <= 1.0) {
                    return scale;
                }
            }
        } catch (e) {}
        return 1.0;
    }

    /**
     * @param {number} scale
     */
    function setStoredUIScale(scale) {
        try {
            localStorage.setItem(STORAGE_UI_SCALE, String(scale));
        } catch (e) {}
    }

    /**
     * @param {number} newScale
     */
    function updateUIScale(newScale) {
        UI_SCALE = Math.max(0.5, Math.min(1.0, newScale));
        setStoredUIScale(UI_SCALE);
    }

    // Initialize UI scale from storage
    UI_SCALE = getStoredUIScale();

    //==========================
    // SCHEDULED ACTIVITIES BUILDER FEATURE
    //==========================
    // This section contains all functions related to the Scheduled Activities Builder.
    // This feature automates adding scheduled activities by allowing users to select
    // segments, study events, and forms, then automatically populating and submitting them.
    //==========================

    var STORAGE_SA_BUILDER_EXISTING = "activityPlanState.saBuilder.existing";
    var STORAGE_SA_BUILDER_SEGMENTS = "activityPlanState.saBuilder.segments";
    var STORAGE_SA_BUILDER_STUDY_EVENTS = "activityPlanState.saBuilder.studyEvents";
    var STORAGE_SA_BUILDER_FORMS = "activityPlanState.saBuilder.forms";
    var STORAGE_SA_BUILDER_USER_SELECTION = "activityPlanState.saBuilder.userSelection";
    var STORAGE_SA_BUILDER_MAPPED_ITEMS = "activityPlanState.saBuilder.mappedItems";
    var STORAGE_SA_BUILDER_CURRENT_INDEX = "activityPlanState.saBuilder.currentIndex";
    var STORAGE_SA_BUILDER_RUNNING = "activityPlanState.saBuilder.running";
    var STORAGE_SA_BUILDER_TIME_OFFSET = "activityPlanState.saBuilder.timeOffset";
    var STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES = "activityPlanState.saBuilder.segmentCheckboxes";
    var SA_BUILDER_POPUP_REF = null;
    /** @type {any} */
    var SA_BUILDER_PROGRESS_POPUP_REF = null;
    var SA_BUILDER_CANCELLED = false;
    var SA_BUILDER_PAUSE = false;
    var SA_BUILDER_TARGET_URL = "https://cenexel.clinspark.com/secure/crfdesign/activityplans/show/";

    function SABuilderFunctions() {}

    // Normalize text for comparison: trim, collapse whitespace
    function normalizeSAText(t) {
        if (typeof t !== "string") return "";
        return t.trim().replace(/\s+/g, " ");
    }

    // Clear all SA Builder storage
    function clearSABuilderStorage() {
        try {
            localStorage.removeItem(STORAGE_SA_BUILDER_EXISTING);
            localStorage.removeItem(STORAGE_SA_BUILDER_SEGMENTS);
            localStorage.removeItem(STORAGE_SA_BUILDER_STUDY_EVENTS);
            localStorage.removeItem(STORAGE_SA_BUILDER_FORMS);
            localStorage.removeItem(STORAGE_SA_BUILDER_USER_SELECTION);
            localStorage.removeItem(STORAGE_SA_BUILDER_MAPPED_ITEMS);
            localStorage.removeItem(STORAGE_SA_BUILDER_CURRENT_INDEX);
            localStorage.removeItem(STORAGE_SA_BUILDER_RUNNING);
            localStorage.removeItem(STORAGE_SA_BUILDER_TIME_OFFSET);
            localStorage.removeItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES);
        } catch (e) {}
        log("SA Builder: storage cleared");
    }

    // Check if user is on the correct page
    function isOnSABuilderPage() {
        var currentUrl = location.href;
        return currentUrl.indexOf(SA_BUILDER_TARGET_URL) !== -1;
    }

    // Scan existing table and collect items
    function scanExistingSATable() {
        var existing = [];
        var tbody = document.getElementById("saTableBody");
        if (!tbody) {
            log("SA Builder: saTableBody not found");
            return existing;
        }
        var rows = tbody.querySelectorAll("tr");
        log("SA Builder: found " + rows.length + " rows in saTableBody");
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var cells = row.querySelectorAll("td");
            if (cells.length >= 4) {
                var segment = normalizeSAText(cells[1].textContent);
                var studyEvent = normalizeSAText(cells[2].textContent);
                var form = normalizeSAText(cells[3].textContent);
                if (segment && studyEvent && form) {
                    var key = segment + " - " + studyEvent + " - " + form;
                    existing.push(key);
                }
            }
        }
        log("SA Builder: collected " + existing.length + " existing items");
        return existing;
    }

    // Check if Add button is disabled
    function isAddSaButtonDisabled() {
        var btn = document.getElementById("addSaButton");
        if (!btn) {
            log("SA Builder: addSaButton not found");
            return true;
        }
        return btn.hasAttribute("disabled");
    }

    // Click the Add button
    function clickAddSaButton() {
        var btn = document.getElementById("addSaButton");
        if (!btn) {
            log("SA Builder: addSaButton not found for click");
            return false;
        }
        btn.click();
        log("SA Builder: Add button clicked");
        return true;
    }

    // Wait for modal to appear and be ready
    async function waitForSAModal(timeoutMs) {
        var start = Date.now();
        var maxTime = timeoutMs || 10000;
        while (Date.now() - start < maxTime) {
            var modal = document.getElementById("ajaxModal");
            if (modal && modal.classList.contains("in")) {
                var modalBody = modal.querySelector("#modalbody, .modal-body");
                if (modalBody) {
                    await sleep(500);
                    return modal;
                }
            }
            await sleep(300);
        }
        return null;
    }

    // Wait for modal to close
    async function waitForSAModalClose(timeoutMs) {
        var start = Date.now();
        var maxTime = timeoutMs || 10000;
        while (Date.now() - start < maxTime) {
            var modal = document.getElementById("ajaxModal");
            if (!modal || !modal.classList.contains("in")) {
                return true;
            }
            await sleep(300);
        }
        return false;
    }

    // Click Select2 field to open dropdown and force options to load
    async function openSelect2Dropdown(containerId) {
        var container = document.getElementById(containerId);
        if (!container) return false;
        var choice = container.querySelector(".select2-choice");
        if (choice) {
            choice.click();
            await sleep(300);
            return true;
        }
        return false;
    }

    // Close Select2 dropdown
    async function closeSelect2Dropdown() {
        var active = document.querySelector(".select2-drop-active");
        if (active) {
            var body = document.body;
            body.click();
            await sleep(200);
        }
    }

    // Collect options from a select element
    function collectSelectOptions(selectId) {
        var select = document.getElementById(selectId);
        if (!select) return [];
        var options = [];
        var opts = select.querySelectorAll("option");
        for (var i = 0; i < opts.length; i++) {
            var opt = opts[i];
            var val = (opt.value || "").trim();
            var txt = normalizeSAText(opt.textContent);
            if (val && txt) {
                options.push({ value: val, text: txt });
            }
        }
        return options;
    }

    // Collect all dropdown data from the Add SA modal
    async function collectModalDropdownData() {
        log("SA Builder: collecting dropdown data from modal");

        // Collect segments
        await openSelect2Dropdown("s2id_segment");
        await sleep(300);
        await closeSelect2Dropdown();
        var segments = collectSelectOptions("segment");
        log("SA Builder: collected " + segments.length + " segments");

        // Collect study events
        await openSelect2Dropdown("s2id_studyEvent");
        await sleep(300);
        await closeSelect2Dropdown();
        var studyEvents = collectSelectOptions("studyEvent");
        log("SA Builder: collected " + studyEvents.length + " study events");

        // Collect forms
        await openSelect2Dropdown("s2id_form");
        await sleep(300);
        await closeSelect2Dropdown();
        var forms = collectSelectOptions("form");
        log("SA Builder: collected " + forms.length + " forms");

        return {
            segments: segments,
            studyEvents: studyEvents,
            forms: forms
        };
    }

    async function selectSelect2Value(selectId, value) {
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled during selectSelect2Value");
            return false;
        }

        var select = document.getElementById(selectId);
        if (!select) {
            log("SA Builder: select element " + selectId + " not found");
            return false;
        }

        // Set the value directly
        select.value = value;

        // Trigger change event
        var evt = new Event("change", { bubbles: true });
        select.dispatchEvent(evt);

        // Also try to update Select2
        try {
            if (window.jQuery && window.jQuery.fn.select2) {
                window.jQuery("#" + selectId).trigger("change");
            }
        } catch (e) {}

        await sleep(300);

        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled after selectSelect2Value sleep");
            return false;
        }

        log("SA Builder: selected value " + value + " in " + selectId);
        return true;
    }

    // Create the selection GUI popup
    function createSABuilderSelectionGUI(segments, studyEvents, forms) {
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "16px";
        container.style.height = "100%";
        container.style.minHeight = "500px";

        // Search bars row
        var searchRow = document.createElement("div");
        searchRow.style.display = "grid";
        searchRow.style.gridTemplateColumns = "1fr 1fr 1fr 200px";
        searchRow.style.gap = "8px";

        var segmentSearch = document.createElement("input");
        segmentSearch.type = "text";
        segmentSearch.placeholder = "Search Segments...";
        segmentSearch.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;";

        var eventSearch = document.createElement("input");
        eventSearch.type = "text";
        eventSearch.placeholder = "Search Study Events...";
        eventSearch.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;";

        var formSearch = document.createElement("input");
        formSearch.type = "text";
        formSearch.placeholder = "Search Forms...";
        formSearch.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;";

        var timeLabel = document.createElement("div");
        timeLabel.textContent = "Time Relative to Segment";
        timeLabel.style.cssText = "padding:8px;font-weight:600;text-align:center;";

        searchRow.appendChild(segmentSearch);
        searchRow.appendChild(eventSearch);
        searchRow.appendChild(formSearch);
        searchRow.appendChild(timeLabel);
        container.appendChild(searchRow);

        // Main content row
        var contentRow = document.createElement("div");
        contentRow.style.display = "grid";
        contentRow.style.gridTemplateColumns = "1fr 1fr 1fr 200px";
        contentRow.style.gap = "8px";
        contentRow.style.flex = "1";
        contentRow.style.overflow = "hidden";

        // Segments column (with checkboxes and dropzones)
        var segmentColumnWrapper = document.createElement("div");
        segmentColumnWrapper.style.cssText = "display:flex;flex-direction:column;border:1px solid #333;border-radius:4px;background:#1a1a1a;height:100%;overflow:hidden;";

        // Select All header for segments
        var segmentHeader = document.createElement("div");
        segmentHeader.style.cssText = "display:flex;align-items:center;gap:8px;padding:8px;border-bottom:1px solid #333;background:#222;";

        var selectAllCheckbox = document.createElement("input");
        selectAllCheckbox.type = "checkbox";
        selectAllCheckbox.id = "saBuilderSelectAllSegments";
        selectAllCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;";

        var selectAllLabel = document.createElement("span");
        selectAllLabel.textContent = "Select All";
        selectAllLabel.style.cssText = "font-weight:600;font-size:13px;";

        segmentHeader.appendChild(selectAllCheckbox);
        segmentHeader.appendChild(selectAllLabel);

        var segmentColumn = document.createElement("div");
        segmentColumn.style.cssText = "overflow-y:auto;padding:8px;flex:1;";
        segmentColumn.id = "saBuilderSegments";

        segmentColumnWrapper.appendChild(segmentHeader);
        segmentColumnWrapper.appendChild(segmentColumn);

        // Study Events column (draggable items)
        var eventColumn = document.createElement("div");
        eventColumn.style.cssText = "overflow-y:auto;border:1px solid #444;border-radius:6px;padding:10px;background:#1e1e1e;box-shadow:inset 0 1px 3px rgba(0,0,0,0.3);";
        eventColumn.id = "saBuilderEvents";

        // Make Study Events column a drop zone for restoring events
        eventColumn.addEventListener("dragover", function(e) {
            e.preventDefault();
            this.style.background = "#252525";
            this.style.border = "2px dashed #007bff";
            this.style.transition = "all 0.2s ease";
        });

        eventColumn.addEventListener("dragleave", function(e) {
            e.preventDefault();
            this.style.background = "#1e1e1e";
            this.style.border = "1px solid #444";
            this.style.transition = "all 0.2s ease";
        });

        eventColumn.addEventListener("drop", function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.style.background = "#1e1e1e";
            this.style.border = "1px solid #444";
            this.style.transition = "all 0.2s ease";

            var eventData = e.dataTransfer.getData("text/plain");
            if (eventData) {
                var data = JSON.parse(eventData);
                if (data.fromSegment) {
                    // Remove from segment
                    if (data.segmentValue && segmentEventMap[data.segmentValue]) {
                        var events = segmentEventMap[data.segmentValue];
                        var index = events.findIndex(function(ev) {
                            return ev.value === data.value;
                        });
                        if (index > -1) {
                            events.splice(index, 1);
                            segmentEventMap[data.segmentValue] = events;
                            // Refresh the segment
                            renderSegments(segmentSearch.value);
                        }
                    }
                    // Restore to column
                    restoreEventToStudyEventsColumn(data);
                    log("SA Builder: Event '" + data.text + "' moved back to Study Events column");
                }
            }
        });

        // Forms column (checkboxes)
        var formColumn = document.createElement("div");
        formColumn.style.cssText = "overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;";
        formColumn.id = "saBuilderForms";

        // Time inputs column
        var timeColumn = document.createElement("div");
        timeColumn.style.cssText = "display:flex;flex-direction:column;gap:12px;padding:8px;border:1px solid #333;border-radius:4px;background:#1a1a1a;height:100%;overflow-y:auto;overflow-x:hidden;";
        timeColumn.id = "saBuilderTime";

        // Track segment-event assignments
        var segmentEventMap = {};

        // Track segment checkbox states separately to prevent resets
        var segmentCheckboxStates = {};

        // Load saved checkbox states
        try {
            var savedStates = localStorage.getItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES);
            if (savedStates) {
                segmentCheckboxStates = JSON.parse(savedStates);
            }
        } catch (e) {}

        // Populate segments
        function renderSegments(filter) {
            segmentColumn.innerHTML = "";
            var filterLower = (filter || "").toLowerCase();
            for (var i = 0; i < segments.length; i++) {
                var seg = segments[i];
                if (filterLower && seg.text.toLowerCase().indexOf(filterLower) === -1) continue;

                var segItem = document.createElement("div");
                segItem.style.cssText = "margin-bottom:8px;padding:8px;border:1px solid #444;border-radius:4px;background:#222;";
                segItem.dataset.segmentValue = seg.value;
                segItem.dataset.segmentText = seg.text;

                var checkRow = document.createElement("div");
                checkRow.style.cssText = "display:flex;align-items:center;gap:8px;";

                var checkbox = document.createElement("input");
                checkbox.type = "checkbox";
                checkbox.dataset.segmentValue = seg.value;
                checkbox.style.cssText = "width:18px;height:18px;cursor:pointer;";

                // Restore checkbox state from saved states
                if (segmentCheckboxStates[seg.value]) {
                    checkbox.checked = true;
                }

                // Save checkbox state when changed
                checkbox.addEventListener("change", function() {
                    segmentCheckboxStates[this.dataset.segmentValue] = this.checked;
                    try {
                        localStorage.setItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES, JSON.stringify(segmentCheckboxStates));
                    } catch (e) {}
                    updateSelectAllCheckbox();
                });

                var label = document.createElement("span");
                label.textContent = seg.text;
                label.style.cssText = "font-weight:500;";

                checkRow.appendChild(checkbox);
                checkRow.appendChild(label);
                segItem.appendChild(checkRow);

                // Dropzone for study events
                var dropzone = document.createElement("div");
                dropzone.style.cssText = "min-height:30px;margin-top:8px;padding:4px;border:1px dashed #555;border-radius:4px;background:#1a1a1a;";
                dropzone.dataset.segmentValue = seg.value;
                dropzone.className = "sa-segment-dropzone";

                // Initialize the map
                if (!segmentEventMap[seg.value]) {
                    segmentEventMap[seg.value] = [];
                }

                // Render existing attached events
                function renderAttachedEvents(dz, segVal) {
                    dz.innerHTML = "";
                    var events = segmentEventMap[segVal] || [];
                    if (events.length === 0) {
                        var placeholder = document.createElement("div");
                        placeholder.textContent = "Drop study events here";
                        placeholder.style.cssText = "color:#666;font-size:12px;text-align:center;padding:8px;";
                        dz.appendChild(placeholder);
                    } else {
                        for (var j = 0; j < events.length; j++) {
                            var ev = events[j];
                            var evItem = document.createElement("div");
                            evItem.style.cssText = "display:flex;justify-content:space-between;align-items:center;padding:4px 8px;margin:2px 0;background:#333;border-radius:3px;font-size:12px;";
                            evItem.dataset.eventValue = ev.value;
                            evItem.dataset.eventText = ev.text;
                            evItem.draggable = true;

                            var attachedItem = document.createElement("div");
                            attachedItem.style.cssText = "display:flex;align-items:center;justify-content:space-between;padding:8px 10px;margin-bottom:4px;border:1px solid #666;border-radius:5px;background:#3a3a3a;color:#fff;font-size:13px;font-weight:500;cursor:move;transition:all 0.2s ease;box-shadow:0 1px 3px rgba(0,0,0,0.2);";
                            attachedItem.draggable = true;
                            attachedItem.dataset.eventValue = ev.value;
                            attachedItem.dataset.eventText = ev.text;

                            var eventLabel = document.createElement("span");
                            eventLabel.textContent = ev.text;
                            eventLabel.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";

                            // Make attached events draggable back to Study Events column
                            attachedItem.addEventListener("dragstart", function(e) {
                                e.dataTransfer.setData("text/plain", JSON.stringify({
                                    value: this.dataset.eventValue,
                                    text: this.dataset.eventText,
                                    fromSegment: true,
                                    segmentValue: segVal
                                }));
                                e.dataTransfer.effectAllowed = "move";
                                this.style.opacity = "0.7";
                                this.style.transform = "scale(0.95)";
                                this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
                            });

                            attachedItem.addEventListener("dragend", function(e) {
                                this.style.opacity = "1";
                                this.style.transform = "scale(1)";
                                this.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
                            });

                            attachedItem.addEventListener("mouseenter", function() {
                                this.style.background = "#4a4a4a";
                                this.style.border = "1px solid #777";
                                this.style.transform = "translateY(-1px)";
                                this.style.boxShadow = "0 2px 5px rgba(0,0,0,0.3)";
                            });

                            attachedItem.addEventListener("mouseleave", function() {
                                this.style.background = "#3a3a3a";
                                this.style.border = "1px solid #666";
                                this.style.transform = "translateY(0)";
                                this.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
                            });

                            var removeBtn = document.createElement("button");
                            removeBtn.textContent = "×";
                            removeBtn.style.cssText = "background:transparent;border:none;color:#ff6b6b;cursor:pointer;font-size:18px;font-weight:bold;padding:0;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;transition:all 0.2s ease;flex-shrink:0;margin-left:8px;";

                            removeBtn.addEventListener("mouseenter", function() {
                                this.style.background = "rgba(255,107,107,0.2)";
                                this.style.color = "#ff5252";
                                this.style.transform = "scale(1.1)";
                            });

                            removeBtn.addEventListener("mouseleave", function() {
                                this.style.background = "transparent";
                                this.style.color = "#ff6b6b";
                                this.style.transform = "scale(1)";
                            });

                            removeBtn.addEventListener("click", function(e) {
                                e.stopPropagation();
                                var index = events.indexOf(ev);
                                if (index > -1) {
                                    events.splice(index, 1);
                                    segmentEventMap[segVal] = events;
                                    renderAttachedEvents(dz, segVal);
                                    log("SA Builder: Event '" + ev.text + "' removed from segment");

                                    // Restore event to Study Events column
                                    restoreEventToStudyEventsColumn({ value: ev.value, text: ev.text });
                                }
                            });

                            attachedItem.appendChild(eventLabel);
                            attachedItem.appendChild(removeBtn);
                            dz.appendChild(attachedItem);

                            // Make attached events re-draggable
                            evItem.addEventListener("dragstart", (function(evData, segV) {
                                return function(e) {
                                    e.dataTransfer.setData("text/plain", JSON.stringify({ value: evData.value, text: evData.text, fromSegment: segV }));
                                    e.dataTransfer.effectAllowed = "move";
                                };
                            })(ev, segVal));

                            dz.appendChild(evItem);
                        }
                    }
                }

                renderAttachedEvents(dropzone, seg.value);

                // Handle dragover
                dropzone.addEventListener("dragover", function(e) {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = "copy";
                    this.style.background = "#2a2a4a";
                });

                dropzone.addEventListener("dragleave", function(e) {
                    this.style.background = "#1a1a1a";
                });

                // Handle drop
                dropzone.addEventListener("drop", (function(segVal) {
                    return function(e) {
                        e.preventDefault();
                        this.style.background = "#1a1a1a";
                        try {
                            var data = JSON.parse(e.dataTransfer.getData("text/plain"));
                            if (data && data.value && data.text) {
                                // Remove from old segment if exists
                                if (data.fromSegment && segmentEventMap[data.fromSegment]) {
                                    segmentEventMap[data.fromSegment] = segmentEventMap[data.fromSegment].filter(function(x) {
                                        return x.value !== data.value;
                                    });
                                }
                                // Check for duplicate in this segment
                                var existing = (segmentEventMap[segVal] || []).find(function(x) { return x.value === data.value; });
                                if (!existing) {
                                    if (!segmentEventMap[segVal]) segmentEventMap[segVal] = [];
                                    segmentEventMap[segVal].push({ value: data.value, text: data.text });

                                    removeEventFromStudyEventsColumn(data.value);
                                }
                                renderSegments(segmentSearch.value);
                            }
                        } catch (err) {
                            log("SA Builder: drop error " + err);
                        }
                    };
                })(seg.value));

                segItem.appendChild(dropzone);
                segmentColumn.appendChild(segItem);
            }
        }

        // Populate study events (draggable)
        function renderStudyEvents(filter) {
            eventColumn.innerHTML = "";
            var filterLower = (filter || "").toLowerCase();
            for (var i = 0; i < studyEvents.length; i++) {
                var ev = studyEvents[i];
                if (filterLower && ev.text.toLowerCase().indexOf(filterLower) === -1) continue;

                var evItem = document.createElement("div");
                evItem.style.cssText = "padding:10px 12px;margin-bottom:6px;border:1px solid #555;border-radius:6px;background:#2a2a2a;cursor:grab;color:#fff;font-size:14px;font-weight:500;transition:all 0.2s ease;box-shadow:0 2px 4px rgba(0,0,0,0.2);";
                evItem.textContent = ev.text;
                evItem.draggable = true;
                evItem.dataset.eventValue = ev.value;
                evItem.dataset.eventText = ev.text;

                evItem.addEventListener("mouseenter", function() {
                    this.style.background = "#333";
                    this.style.border = "1px solid #666";
                    this.style.transform = "translateY(-1px)";
                    this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
                });

                evItem.addEventListener("mouseleave", function() {
                    this.style.background = "#2a2a2a";
                    this.style.border = "1px solid #555";
                    this.style.transform = "translateY(0)";
                    this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
                });

                evItem.addEventListener("dragstart", function(e) {
                    e.dataTransfer.setData("text/plain", JSON.stringify({
                        value: this.dataset.eventValue,
                        text: this.dataset.eventText
                    }));
                    e.dataTransfer.effectAllowed = "copy";
                    this.style.opacity = "0.6";
                    this.style.transform = "scale(0.95)";
                    this.style.boxShadow = "0 6px 12px rgba(0,0,0,0.4)";
                });

                evItem.addEventListener("dragend", function(e) {
                    this.style.opacity = "1";
                    this.style.transform = "scale(1)";
                    this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
                });

                eventColumn.appendChild(evItem);
                sortStudyEventsColumn();
            }
        }

        // Remove event from Study Events column
        function removeEventFromStudyEventsColumn(eventValue) {
            var eventItems = eventColumn.querySelectorAll("[data-event-value]");
            for (var i = 0; i < eventItems.length; i++) {
                var item = eventItems[i];
                if (item.dataset.eventValue === eventValue) {
                    item.remove();
                    log("SA Builder: Event removed from Study Events column");
                    break;
                }
            }
        }
        // Restore event to Study Events column
        function restoreEventToStudyEventsColumn(eventData) {
            // Check if event already exists in column
            var existingItems = eventColumn.querySelectorAll("[data-event-value]");
            for (var i = 0; i < existingItems.length; i++) {
                if (existingItems[i].dataset.eventValue === eventData.value) {
                    log("SA Builder: Event already exists in Study Events column");
                    return;
                }
            }

            // Create the event item
            var evItem = document.createElement("div");
            evItem.textContent = eventData.text;
            evItem.style.cssText = "padding:10px 12px;margin-bottom:6px;border:1px solid #555;border-radius:6px;background:#2a2a2a;cursor:grab;color:#fff;font-size:14px;font-weight:500;transition:all 0.2s ease;box-shadow:0 2px 4px rgba(0,0,0,0.2);animation:slideIn 0.3s ease;";
            evItem.draggable = true;
            evItem.dataset.eventValue = eventData.value;
            evItem.dataset.eventText = eventData.text;

            evItem.addEventListener("mouseenter", function() {
                this.style.background = "#333";
                this.style.border = "1px solid #666";
                this.style.transform = "translateY(-1px)";
                this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
            });

            evItem.addEventListener("mouseleave", function() {
                this.style.background = "#2a2a2a";
                this.style.border = "1px solid #555";
                this.style.transform = "translateY(0)";
                this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
            });

            evItem.addEventListener("dragstart", function(e) {
                e.dataTransfer.setData("text/plain", JSON.stringify({
                    value: this.dataset.eventValue,
                    text: this.dataset.eventText
                }));
                e.dataTransfer.effectAllowed = "copy";
                this.style.opacity = "0.6";
                this.style.transform = "scale(0.95)";
                this.style.boxShadow = "0 6px 12px rgba(0,0,0,0.4)";
            });

            evItem.addEventListener("dragend", function(e) {
                this.style.opacity = "1";
                this.style.transform = "scale(1)";
                this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
            });

            // Add the event to the column and sort
            eventColumn.appendChild(evItem);
            sortStudyEventsColumn();
            log("SA Builder: Event '" + eventData.text + "' restored to Study Events column");
        }

        // Sort the Study Events column alphabetically
        function sortStudyEventsColumn() {
            var allItems = eventColumn.querySelectorAll("[data-event-value]");
            var itemsArray = Array.prototype.slice.call(allItems);

            // Sort by text content (case-insensitive)
            itemsArray.sort(function(a, b) {
                var textA = (a.dataset.eventText || a.textContent || "").toLowerCase();
                var textB = (b.dataset.eventText || b.textContent || "").toLowerCase();
                return textA.localeCompare(textB);
            });

            // Clear and re-append in sorted order
            eventColumn.innerHTML = "";
            itemsArray.forEach(function(item) {
                eventColumn.appendChild(item);
            });

            log("SA Builder: Study Events column sorted alphabetically");
        }

        // Populate forms (checkboxes)
        function renderForms(filter) {
            formColumn.innerHTML = "";
            var filterLower = (filter || "").toLowerCase();
            for (var i = 0; i < forms.length; i++) {
                var frm = forms[i];
                if (filterLower && frm.text.toLowerCase().indexOf(filterLower) === -1) continue;

                var frmItem = document.createElement("div");
                frmItem.style.cssText = "display:flex;align-items:flex-start;gap:8px;padding:6px;margin-bottom:4px;border:1px solid #444;border-radius:4px;background:#222;";

                var checkbox = document.createElement("input");
                checkbox.type = "checkbox";
                checkbox.dataset.formValue = frm.value;
                checkbox.dataset.formText = frm.text;
                checkbox.style.cssText = "width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-top:2px;accent-color:#007bff;";

                var label = document.createElement("span");
                label.textContent = frm.text;
                label.style.cssText = "font-size:13px;word-break:break-word;";

                frmItem.appendChild(checkbox);
                frmItem.appendChild(label);
                formColumn.appendChild(frmItem);
            }
        }

        // Create time inputs
        function createTimeInput(labelText, id, max) {
            var row = document.createElement("div");
            row.style.cssText = "display:flex;flex-direction:column;gap:4px;";

            var lbl = document.createElement("label");
            lbl.textContent = labelText;
            lbl.style.cssText = "font-size:12px;color:#aaa;";

            var input = document.createElement("input");
            input.type = "number";
            input.id = id;
            input.min = "0";
            input.max = String(max);
            input.value = "0";
            input.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;width:100%;";

            row.appendChild(lbl);
            row.appendChild(input);
            return row;
        }

        // Hidden checkbox row
        var hiddenRow = document.createElement("div");
        hiddenRow.style.cssText = "display:flex;align-items:center;gap:8px;padding-top:12px;margin-top:12px;border-top:1px solid #444;";

        var hiddenCheckbox = document.createElement("input");
        hiddenCheckbox.type = "checkbox";
        hiddenCheckbox.id = "saBuilderHidden";
        hiddenCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var hiddenLabel = document.createElement("label");
        hiddenLabel.textContent = "Hidden?";
        hiddenLabel.setAttribute("for", "saBuilderHidden");
        hiddenLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        hiddenRow.appendChild(hiddenCheckbox);
        hiddenRow.appendChild(hiddenLabel);
        timeColumn.appendChild(hiddenRow);

        // Mandatory checkbox row
        var mandatoryRow = document.createElement("div");
        mandatoryRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;";

        var mandatoryCheckbox = document.createElement("input");
        mandatoryCheckbox.type = "checkbox";
        mandatoryCheckbox.id = "saBuilderMandatory";
        mandatoryCheckbox.checked = true;
        mandatoryCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var mandatoryLabel = document.createElement("label");
        mandatoryLabel.textContent = "Mandatory";
        mandatoryLabel.setAttribute("for", "saBuilderMandatory");
        mandatoryLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        mandatoryRow.appendChild(mandatoryCheckbox);
        mandatoryRow.appendChild(mandatoryLabel);
        timeColumn.appendChild(mandatoryRow);

        // Enforce Data Collection Order checkbox row
        var enforceRow = document.createElement("div");
        enforceRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;";

        var enforceCheckbox = document.createElement("input");
        enforceCheckbox.type = "checkbox";
        enforceCheckbox.id = "saBuilderEnforce";
        enforceCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var enforceLabel = document.createElement("label");
        enforceLabel.textContent = "Enforce Data Collection Order";
        enforceLabel.setAttribute("for", "saBuilderEnforce");
        enforceLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        enforceRow.appendChild(enforceCheckbox);
        enforceRow.appendChild(enforceLabel);
        timeColumn.appendChild(enforceRow);

        // Pre-Window input row
        var preWindowRow = document.createElement("div");
        preWindowRow.style.cssText = "display:flex;flex-direction:column;gap:4px;margin-top:8px;";

        var preWindowLabel = document.createElement("label");
        preWindowLabel.textContent = "Pre-Window";
        preWindowLabel.style.cssText = "font-size:12px;color:#aaa;";

        var preWindowInput = document.createElement("input");
        preWindowInput.type = "text";
        preWindowInput.id = "saBuilderPreWindow";
        preWindowInput.placeholder = "";
        preWindowInput.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;width:100%;";

        preWindowRow.appendChild(preWindowLabel);
        preWindowRow.appendChild(preWindowInput);
        timeColumn.appendChild(preWindowRow);

        // Post-Window input row
        var postWindowRow = document.createElement("div");
        postWindowRow.style.cssText = "display:flex;flex-direction:column;gap:4px;margin-top:8px;";

        var postWindowLabel = document.createElement("label");
        postWindowLabel.textContent = "Post-Window";
        postWindowLabel.style.cssText = "font-size:12px;color:#aaa;";

        var postWindowInput = document.createElement("input");
        postWindowInput.type = "text";
        postWindowInput.id = "saBuilderPostWindow";
        postWindowInput.placeholder = "";
        postWindowInput.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;width:100%;";

        postWindowRow.appendChild(postWindowLabel);
        postWindowRow.appendChild(postWindowInput);
        timeColumn.appendChild(postWindowRow);

        // Reference Activity checkbox row
        var refActivityRow = document.createElement("div");
        refActivityRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;padding-top:8px;border-top:1px solid #444;";

        var refActivityCheckbox = document.createElement("input");
        refActivityCheckbox.type = "checkbox";
        refActivityCheckbox.id = "saBuilderRefActivity";
        refActivityCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var refActivityLabel = document.createElement("label");
        refActivityLabel.textContent = "Reference Activity";
        refActivityLabel.setAttribute("for", "saBuilderRefActivity");
        refActivityLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        refActivityRow.appendChild(refActivityCheckbox);
        refActivityRow.appendChild(refActivityLabel);
        timeColumn.appendChild(refActivityRow);

        // Reference Activity disables time inputs and Pre-Reference checkbox
        refActivityCheckbox.addEventListener("change", function() {
            var isChecked = this.checked;
            var daysEl = document.getElementById("saBuilderDays");
            var hoursEl = document.getElementById("saBuilderHours");
            var minutesEl = document.getElementById("saBuilderMinutes");
            var secondsEl = document.getElementById("saBuilderSeconds");
            var preRefEl = document.getElementById("saBuilderPreReference");

            if (daysEl) {
                daysEl.disabled = isChecked;
                daysEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (hoursEl) {
                hoursEl.disabled = isChecked;
                hoursEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (minutesEl) {
                minutesEl.disabled = isChecked;
                minutesEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (secondsEl) {
                secondsEl.disabled = isChecked;
                secondsEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (preRefEl) {
                preRefEl.disabled = isChecked;
                preRefEl.style.opacity = isChecked ? "0.5" : "1";
                if (isChecked) {
                    preRefEl.checked = false;
                }
            }
        });

        // Time inputs section (moved below other inputs)
        var timeInputsSection = document.createElement("div");
        timeInputsSection.style.cssText = "padding-top:12px;margin-top:12px;border-top:1px solid #444;";

        // Pre-Reference checkbox row (positioned right above time inputs)
        var preRefRow = document.createElement("div");
        preRefRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-bottom:12px;";

        var preRefCheckbox = document.createElement("input");
        preRefCheckbox.type = "checkbox";
        preRefCheckbox.id = "saBuilderPreReference";
        preRefCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var preRefLabel = document.createElement("label");
        preRefLabel.textContent = "Pre-Reference?";
        preRefLabel.setAttribute("for", "saBuilderPreReference");
        preRefLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        preRefRow.appendChild(preRefCheckbox);
        preRefRow.appendChild(preRefLabel);
        timeInputsSection.appendChild(preRefRow);

        timeInputsSection.appendChild(createTimeInput("Days", "saBuilderDays", 200));
        timeInputsSection.appendChild(createTimeInput("Hours", "saBuilderHours", 23));
        timeInputsSection.appendChild(createTimeInput("Minutes", "saBuilderMinutes", 59));
        timeInputsSection.appendChild(createTimeInput("Seconds", "saBuilderSeconds", 59));

        timeColumn.appendChild(timeInputsSection);

        // Initial render
        renderSegments("");
        renderStudyEvents("");
        renderForms("");

        // Restore saved time offset values
        try {
            var savedTimeOffset = localStorage.getItem(STORAGE_SA_BUILDER_TIME_OFFSET);
            if (savedTimeOffset) {
                var timeData = JSON.parse(savedTimeOffset);
                if (timeData.days !== undefined) document.getElementById("saBuilderDays").value = timeData.days;
                if (timeData.hours !== undefined) document.getElementById("saBuilderHours").value = timeData.hours;
                if (timeData.minutes !== undefined) document.getElementById("saBuilderMinutes").value = timeData.minutes;
                if (timeData.seconds !== undefined) document.getElementById("saBuilderSeconds").value = timeData.seconds;
            }
        } catch (e) {}

        // Select All checkbox handler
        function updateSelectAllCheckbox() {
            var allChecked = true;
            for (var i = 0; i < segments.length; i++) {
                if (!segmentCheckboxStates[segments[i].value]) {
                    allChecked = false;
                    break;
                }
            }
            selectAllCheckbox.checked = allChecked;
        }

        selectAllCheckbox.addEventListener("change", function() {
            var isChecked = this.checked;
            for (var i = 0; i < segments.length; i++) {
                segmentCheckboxStates[segments[i].value] = isChecked;
            }
            try {
                localStorage.setItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES, JSON.stringify(segmentCheckboxStates));
            } catch (e) {}
            renderSegments(segmentSearch.value);
        });

        // Initialize Select All checkbox state
        updateSelectAllCheckbox();

        // Search handlers
        segmentSearch.addEventListener("input", function() { renderSegments(this.value); });
        eventSearch.addEventListener("input", function() { renderStudyEvents(this.value); });
        formSearch.addEventListener("input", function() { renderForms(this.value); });

        contentRow.appendChild(segmentColumnWrapper);
        contentRow.appendChild(eventColumn);
        contentRow.appendChild(formColumn);
        contentRow.appendChild(timeColumn);
        container.appendChild(contentRow);

        // Confirm button row
        var buttonRow = document.createElement("div");
        buttonRow.style.cssText = "display:flex;justify-content:flex-end;gap:12px;padding-top:12px;border-top:1px solid #333;";

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.cssText = "padding:10px 24px;border-radius:6px;border:none;background:#28a745;color:#fff;font-weight:600;cursor:pointer;";
        confirmBtn.addEventListener("mouseenter", function() { this.style.background = "#218838"; });
        confirmBtn.addEventListener("mouseleave", function() { this.style.background = "#28a745"; });

        confirmBtn.addEventListener("click", function() {
            // Collect selected segments and their events
            var selectedSegments = [];
            var segmentCheckboxes = segmentColumn.querySelectorAll("input[type='checkbox']:checked");
            for (var i = 0; i < segmentCheckboxes.length; i++) {
                var cb = segmentCheckboxes[i];
                var segVal = cb.dataset.segmentValue;
                var segText = cb.closest("[data-segment-text]");
                var segName = segText ? segText.dataset.segmentText : "";
                var events = segmentEventMap[segVal] || [];
                if (events.length > 0) {
                    selectedSegments.push({ value: segVal, text: segName, events: events });
                }
            }

            // Collect selected forms
            var selectedForms = [];
            var formCheckboxes = formColumn.querySelectorAll("input[type='checkbox']:checked");
            for (var j = 0; j < formCheckboxes.length; j++) {
                var fcb = formCheckboxes[j];
                selectedForms.push({ value: fcb.dataset.formValue, text: fcb.dataset.formText });
            }

            // Validate
            if (selectedSegments.length === 0) {
                alert("Please check at least one segment that has study events attached.");
                return;
            }
            if (selectedForms.length === 0) {
                alert("Please select at least one form.");
                return;
            }

            // Get time offset
            var timeOffset = {
                days: parseInt(document.getElementById("saBuilderDays").value) || 0,
                hours: parseInt(document.getElementById("saBuilderHours").value) || 0,
                minutes: parseInt(document.getElementById("saBuilderMinutes").value) || 0,
                seconds: parseInt(document.getElementById("saBuilderSeconds").value) || 0
            };

            log("SA Builder: Time offset - Days: " + timeOffset.days + ", Hours: " + timeOffset.hours + ", Minutes: " + timeOffset.minutes + ", Seconds: " + timeOffset.seconds);

            var hiddenChecked = document.getElementById("saBuilderHidden").checked;
            var mandatoryChecked = document.getElementById("saBuilderMandatory").checked;
            var enforceChecked = document.getElementById("saBuilderEnforce").checked;
            var preWindowValue = document.getElementById("saBuilderPreWindow").value.trim();
            var postWindowValue = document.getElementById("saBuilderPostWindow").value.trim();
            var refActivityChecked = document.getElementById("saBuilderRefActivity").checked;
            var preReferenceChecked = document.getElementById("saBuilderPreReference").checked;

            // If Reference Activity is checked, void the time offset
            if (refActivityChecked) {
                timeOffset = { days: 0, hours: 0, minutes: 0, seconds: 0 };
            }

            var userSelection = {
                segments: selectedSegments,
                forms: selectedForms,
                timeOffset: timeOffset,
                hidden: hiddenChecked,
                mandatory: mandatoryChecked,
                enforce: enforceChecked,
                preWindow: preWindowValue,
                postWindow: postWindowValue,
                refActivity: refActivityChecked,
                preReference: preReferenceChecked
            };
            try {
                localStorage.setItem(STORAGE_SA_BUILDER_USER_SELECTION, JSON.stringify(userSelection));
                localStorage.setItem(STORAGE_SA_BUILDER_TIME_OFFSET, JSON.stringify(timeOffset));
            } catch (e) {}

            // Close selection popup and start adding
            if (SA_BUILDER_POPUP_REF) {
                SA_BUILDER_POPUP_REF.close();
                SA_BUILDER_POPUP_REF = null;
            }
            if (SA_BUILDER_CANCELLED) {
                log("SA Builder: cancelled");
                return;
            }
            // Start the adding process
            startSABuilderAddProcess(userSelection);
        });

        buttonRow.appendChild(confirmBtn);
        container.appendChild(buttonRow);

        return container;
    }

    // Create mapped item list from user selection
    function createMappedItemList(userSelection, existingItems) {
        var mappedItems = [];
        var segments = userSelection.segments || [];
        var forms = userSelection.forms || [];

        log("SA Builder: creating mapped items from " + segments.length + " segments and " + forms.length + " forms");

        for (var i = 0; i < segments.length; i++) {
            var seg = segments[i];
            var events = seg.events || [];
            for (var j = 0; j < events.length; j++) {
                var ev = events[j];
                for (var k = 0; k < forms.length; k++) {
                    var frm = forms[k];
                    var key = normalizeSAText(seg.text) + " - " + normalizeSAText(ev.text) + " - " + normalizeSAText(frm.text);

                    // Check for duplicates
                    var isDuplicate = existingItems.some(function(existing) {
                        return normalizeSAText(existing) === key;
                    });

                    mappedItems.push({
                        segmentValue: seg.value,
                        segmentText: seg.text,
                        eventValue: ev.value,
                        eventText: ev.text,
                        formValue: frm.value,
                        formText: frm.text,
                        key: key,
                        status: isDuplicate ? "Duplicate (Removed)" : "Incomplete"
                    });
                }
            }
        }

        log("SA Builder: created " + mappedItems.length + " mapped items");
        return mappedItems;
    }

    // Create progress popup
    function createSABuilderProgressPopup(mappedItems) {
        var container = document.createElement("div");
        container.style.cssText = "display:flex;flex-direction:column;gap:12px;max-height:500px;";

        var statusDiv = document.createElement("div");
        statusDiv.id = "saBuilderProgressStatus";
        statusDiv.style.cssText = "text-align:center;font-size:16px;font-weight:600;padding:8px;";
        statusDiv.textContent = "Processing...";

        var loadingDiv = document.createElement("div");
        loadingDiv.id = "saBuilderProgressLoading";
        loadingDiv.style.cssText = "text-align:center;font-size:14px;color:#9df;";
        loadingDiv.textContent = "Running.";

        var listContainer = document.createElement("div");
        listContainer.id = "saBuilderProgressList";
        listContainer.style.cssText = "flex:1;overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;max-height:350px;";

        // Render items
        function renderItems() {
            listContainer.innerHTML = "";
            for (var i = 0; i < mappedItems.length; i++) {
                var item = mappedItems[i];
                var row = document.createElement("div");
                row.style.cssText = "display:flex;justify-content:space-between;align-items:center;padding:6px 8px;margin-bottom:4px;border-radius:4px;font-size:12px;";

                if (item.status === "Complete") {
                    row.style.background = "#1a3a1a";
                } else if (item.status === "Duplicate (Removed)") {
                    row.style.background = "#3a3a1a";
                } else {
                    row.style.background = "#222";
                }

                var keySpan = document.createElement("span");
                keySpan.textContent = item.key;
                keySpan.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-right:8px;";

                var statusSpan = document.createElement("span");
                statusSpan.textContent = item.status;
                statusSpan.style.cssText = "font-weight:500;flex-shrink:0;";
                if (item.status === "Complete") {
                    statusSpan.style.color = "#4f4";
                } else if (item.status === "Duplicate (Removed)") {
                    statusSpan.style.color = "#ff4";
                } else {
                    statusSpan.style.color = "#aaa";
                }

                row.appendChild(keySpan);
                row.appendChild(statusSpan);
                listContainer.appendChild(row);
            }
        }

        renderItems();

        var buttonRow = document.createElement("div");
        buttonRow.style.cssText = "display:flex;justify-content:center;padding-top:12px;";

        var stopBtn = document.createElement("button");
        stopBtn.textContent = "Close";
        stopBtn.style.cssText = "padding:10px 24px;border-radius:6px;border:none;background:#dc3545;color:#fff;font-weight:600;cursor:pointer;";
        stopBtn.addEventListener("click", function() {
            SA_BUILDER_CANCELLED = true;
            log("SA Builder: stopped by user");
            clearSABuilderStorage();
            if (SA_BUILDER_PROGRESS_POPUP_REF && SA_BUILDER_PROGRESS_POPUP_REF.close) {
                try {
                    SA_BUILDER_PROGRESS_POPUP_REF.close();
                } catch (e) {
                    log("SA Builder: error closing popup - " + String(e));
                }
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
        });
        if (SA_BUILDER_CANCELLED) {
            log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
            log("SA Builder: cancelled");
            return;
        }
        buttonRow.appendChild(stopBtn);

        container.appendChild(statusDiv);
        container.appendChild(loadingDiv);
        container.appendChild(listContainer);
        container.appendChild(buttonRow);

        // Animate loading
        var dots = 1;
        var loadingInterval = setInterval(function() {
            if (!SA_BUILDER_PROGRESS_POPUP_REF || SA_BUILDER_CANCELLED) {
                clearInterval(loadingInterval);
                return;
            }
            dots = (dots % 3) + 1;
            var text = "Running";
            for (var i = 0; i < dots; i++) text += ".";
            loadingDiv.textContent = text;
        }, 500);

        return {
            element: container,
            updateStatus: function(text) {
                statusDiv.textContent = text;
            },
            updateItem: function(index, status) {
                if (mappedItems[index]) {
                    mappedItems[index].status = status;
                    renderItems();
                }
            },
            setComplete: function() {
                clearInterval(loadingInterval);
                loadingDiv.textContent = "Done!";
                loadingDiv.style.color = "#4f4";
                stopBtn.textContent = "Close";
                stopBtn.style.background = "#28a745";
            }
        };
    }

    // Start the adding process
    async function startSABuilderAddProcess(userSelection) {
        // Get existing items from storage
        var existingItems = [];
        try {
            var raw = localStorage.getItem(STORAGE_SA_BUILDER_EXISTING);
            if (raw) existingItems = JSON.parse(raw);
        } catch (e) {}

        // Create mapped items
        var mappedItems = createMappedItemList(userSelection, existingItems);

        // Log mapped items
        log("SA Builder: Mapped items list:");
        for (var i = 0; i < mappedItems.length; i++) {
            log("  " + (i + 1) + ". " + mappedItems[i].key + " [" + mappedItems[i].status + "]");
        }

        // Store for persistence
        try {
            localStorage.setItem(STORAGE_SA_BUILDER_MAPPED_ITEMS, JSON.stringify(mappedItems));
            localStorage.setItem(STORAGE_SA_BUILDER_RUNNING, "1");
            localStorage.setItem(STORAGE_SA_BUILDER_CURRENT_INDEX, "0");
        } catch (e) {}

        // Create progress popup
        var progressContent = createSABuilderProgressPopup(mappedItems);
        SA_BUILDER_PROGRESS_POPUP_REF = createPopup({
            title: "Adding Scheduled Activities",
            content: progressContent.element,
            width: "600px",
            height: "auto",
            maxHeight: "80%",
            onClose: function() {
                SA_BUILDER_CANCELLED = true;
                log("SA Builder: cancelled by user (X button)");
                clearSABuilderStorage();
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
        });
        log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled");
            return;
        }
        // Process items
        var timeOffset = userSelection.timeOffset || { days: 0, hours: 0, minutes: 0, seconds: 0 };
        log("SA Builder: Using time offset - Days: " + timeOffset.days + ", Hours: " + timeOffset.hours + ", Minutes: " + timeOffset.minutes + ", Seconds: " + timeOffset.seconds);

        for (var idx = 0; idx < mappedItems.length; idx++) {
            if (SA_BUILDER_CANCELLED) {
                log("SA Builder: cancelled");
                break;
            }

            var item = mappedItems[idx];

            // Skip duplicates
            if (item.status === "Duplicate (Removed)") {
                log("SA Builder: skipping duplicate - " + item.key);
                continue;
            }

            progressContent.updateStatus("Adding item " + (idx + 1) + " of " + mappedItems.length);
            log("SA Builder: adding item " + (idx + 1) + " - " + item.key);

            try {
                // Check cancellation before starting
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before Add button");
                    break;
                }

                // Click Add button
                if (!clickAddSaButton()) {
                    log("SA Builder: failed to click Add button");
                    break;
                }

                // Wait for modal
                var modal = await waitForSAModal(10000);
                if (!modal) {
                    log("SA Builder: modal did not appear");
                    break;
                }

                // Check cancellation after modal appears
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after modal appeared");
                    break;
                }

                await sleep(1000);

                // Check cancellation before selections
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before selections");
                    break;
                }
                log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
                // Select segment
                log("SA Builder: selecting segment " + item.segmentText + " (value: " + item.segmentValue + ")");
                await selectSelect2Value("segment", item.segmentValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after segment selection");
                    break;
                }

                // Select study event
                log("SA Builder: selecting study event " + item.eventText + " (value: " + item.eventValue + ")");
                await selectSelect2Value("studyEvent", item.eventValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after study event selection");
                    break;
                }

                // Select form
                log("SA Builder: selecting form " + item.formText + " (value: " + item.formValue + ")");
                await selectSelect2Value("form", item.formValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after form selection");
                    break;
                }
                // Check Hidden checkbox if user selected it
                if (userSelection.hidden) {
                    var hiddenCheckboxEl = document.querySelector("#uniform-hidden span input#hidden.checkbox");
                    if (!hiddenCheckboxEl) {
                        hiddenCheckboxEl = document.getElementById("hidden");
                    }
                    if (hiddenCheckboxEl && !hiddenCheckboxEl.checked) {
                        hiddenCheckboxEl.click();
                        log("SA Builder: Hidden checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Mandatory checkbox (default is checked, uncheck if user unchecked it)
                if (!userSelection.mandatory) {
                    var mandatoryEl = document.querySelector("#uniform-mandatory span input#mandatory.checkbox");
                    if (!mandatoryEl) {
                        mandatoryEl = document.getElementById("mandatory");
                    }
                    if (mandatoryEl && mandatoryEl.checked) {
                        mandatoryEl.click();
                        log("SA Builder: Mandatory checkbox unchecked");
                        await sleep(200);
                    }
                }

                // Handle Enforce Data Collection Order checkbox (default is unchecked, check if user checked it)
                if (userSelection.enforce) {
                    var enforceEl = document.querySelector("#uniform-enforceDataCollectionOrder span input#enforceDataCollectionOrder.checkbox");
                    if (!enforceEl) {
                        enforceEl = document.getElementById("enforceDataCollectionOrder");
                    }
                    if (enforceEl && !enforceEl.checked) {
                        enforceEl.click();
                        log("SA Builder: Enforce Data Collection Order checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Pre-Window and Post-Window fields
                var preWindow = document.getElementById("preWindow");
                var postWindow = document.getElementById("postWindow");
                if (preWindow) {
                    preWindow.value = userSelection.preWindow || "";
                    preWindow.dispatchEvent(new Event("input", { bubbles: true }));
                    log("SA Builder: Pre-Window set to '" + (userSelection.preWindow || "") + "'");
                }
                if (postWindow) {
                    postWindow.value = userSelection.postWindow || "";
                    postWindow.dispatchEvent(new Event("input", { bubbles: true }));
                    log("SA Builder: Post-Window set to '" + (userSelection.postWindow || "") + "'");
                }

                // Handle Reference Activity checkbox
                if (userSelection.refActivity) {
                    var refActivityEl = document.getElementById("referenceActivity");
                    if (refActivityEl && !refActivityEl.checked) {
                        refActivityEl.click();
                        log("SA Builder: Reference Activity checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Pre-Reference checkbox
                if (userSelection.preReference) {
                    var preRefEl = document.querySelector("#uniform-offset\\.preReference span input#offset\\.preReference.checkbox");
                    if (!preRefEl) {
                        preRefEl = document.getElementById("offset.preReference");
                    }
                    if (preRefEl && !preRefEl.checked) {
                        preRefEl.click();
                        log("SA Builder: Pre-Reference checkbox checked");
                        await sleep(200);
                    }
                }

                // Set time offset values (skip if Reference Activity is checked)
                if (!userSelection.refActivity) {
                    var daysInput = document.querySelector("input[name='offset.days']");
                    var hoursInput = document.querySelector("input[name='offset.hours']");
                    var minutesInput = document.querySelector("input[name='offset.minutes']");
                    var secondsInput = document.querySelector("input[name='offset.seconds']");

                    if (daysInput) {
                        daysInput.value = String(timeOffset.days);
                        daysInput.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                    if (hoursInput) {
                        hoursInput.value = String(timeOffset.hours);
                        hoursInput.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                    if (minutesInput) {
                        minutesInput.value = String(timeOffset.minutes);
                        minutesInput.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                    if (secondsInput) {
                        secondsInput.value = String(timeOffset.seconds);
                        secondsInput.dispatchEvent(new Event("input", { bubbles: true }));
                    }
                } else {
                    log("SA Builder: Skipping time offset values (Reference Activity checked)");
                }

                await sleep(300);

                // Check cancellation before saving
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before Save button");
                    break;
                }

                // Click Save
                var saveBtn = document.getElementById("actionButton");
                if (saveBtn) {
                    saveBtn.click();
                    log("SA Builder: Save button clicked");
                } else {
                    log("SA Builder: Save button not found");
                    break;
                }

                // Wait for modal to close
                var closed = await waitForSAModalClose(10000);
                if (!closed) {
                    log("SA Builder: modal did not close");
                    break;
                }

                // Check cancellation after save completes
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after save completed");
                    break;
                }

                // Update status
                mappedItems[idx].status = "Complete";
                progressContent.updateItem(idx, "Complete");
                log("SA Builder: item completed - " + item.key);

                // Wait for table refresh
                await sleep(2000);

                // Check cancellation after table refresh wait
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after table refresh wait");
                    break;
                }

            } catch (err) {
                log("SA Builder: error adding item - " + String(err));
                break;
            }
        }

        // Complete
        if (!SA_BUILDER_CANCELLED) {
            progressContent.updateStatus("Completed!");
            progressContent.setComplete();
            log("SA Builder: all items processed");
        }

        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled");
            return;
        }
        // Clear running state
        try {
            localStorage.removeItem(STORAGE_SA_BUILDER_RUNNING);
            localStorage.removeItem(STORAGE_SA_BUILDER_CURRENT_INDEX);
        } catch (e) {}
    }

    // Main entry point for SA Builder
    async function runSABuilder() {
        // Check if on correct page
        if (!isOnSABuilderPage()) {
            log("SA Builder: wrong page - " + location.href);
            showWrongPagePopup("Scheduled Activities Builder", SA_BUILDER_TARGET_URL, location.pathname);
            return;
        }

        // Show loading popup
        var loadingContent = document.createElement("div");
        loadingContent.style.cssText = "text-align:center;padding:30px;";
        loadingContent.innerHTML = '<div style="font-size:16px;margin-bottom:16px;">Collecting data...</div><div id="saBuilderLoadingDots" style="color:#9df;">Loading.</div>';

        var loadingPopup = createPopup({
            title: "Scheduled Activities Builder",
            content: loadingContent,
            width: "350px",
            height: "auto",
        });

        // Animate loading
        var dots = 1;
        var loadingInterval = setInterval(function() {
            var el = document.getElementById("saBuilderLoadingDots");
            if (!el || SA_BUILDER_CANCELLED) {
                clearInterval(loadingInterval);
                return;
            }
            dots = (dots % 3) + 1;
            var text = "Loading";
            for (var i = 0; i < dots; i++) text += ".";
            el.textContent = text;
        }, 500);

        // Scan existing table
        var existingItems = scanExistingSATable();
        try {
            localStorage.setItem(STORAGE_SA_BUILDER_EXISTING, JSON.stringify(existingItems));
        } catch (e) {}

        // Check if Add button is disabled
        if (isAddSaButtonDisabled()) {
            clearInterval(loadingInterval);
            loadingPopup.close();
            createPopup({
                title: "Scheduled Activities Builder",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#f66;font-size:16px;margin-bottom:16px;">⚠️ Add Button Disabled</p><p>The Add button is currently disabled. This activity plan may no longer be in design mode.</p></div>',
                width: "400px",
                height: "auto"
            });
            log("SA Builder: Add button is disabled");
            return;
        }

        // Click Add button to open modal
        if (!clickAddSaButton()) {
            clearInterval(loadingInterval);
            loadingPopup.close();
            createPopup({
                title: "Scheduled Activities Builder",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#f66;">Failed to find or click the Add button.</p></div>',
                width: "350px",
                height: "auto"
            });
            return;
        }

        // Wait for modal
        var modal = await waitForSAModal(10000);
        if (!modal) {
            clearInterval(loadingInterval);
            loadingPopup.close();
            createPopup({
                title: "Scheduled Activities Builder",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#f66;">Modal did not appear within timeout.</p></div>',
                width: "350px",
                height: "auto"
            });
            return;
        }

        // Collect dropdown data
        var dropdownData = await collectModalDropdownData();

        // Store for persistence
        try {
            localStorage.setItem(STORAGE_SA_BUILDER_SEGMENTS, JSON.stringify(dropdownData.segments));
            localStorage.setItem(STORAGE_SA_BUILDER_STUDY_EVENTS, JSON.stringify(dropdownData.studyEvents));
            localStorage.setItem(STORAGE_SA_BUILDER_FORMS, JSON.stringify(dropdownData.forms));
        } catch (e) {}

        // Close the modal
        var closeBtn = modal.querySelector(".close, [data-dismiss='modal']");
        if (closeBtn) {
            closeBtn.click();
        } else {
            var escEvent = new KeyboardEvent("keydown", { key: "Escape", bubbles: true });
            document.dispatchEvent(escEvent);
        }
        await sleep(500);

        clearInterval(loadingInterval);
        loadingPopup.close();

        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled during data collection");
            return;
        }

        // Show selection GUI
        var guiContent = createSABuilderSelectionGUI(dropdownData.segments, dropdownData.studyEvents, dropdownData.forms);
        SA_BUILDER_POPUP_REF = createPopup({
            title: "Scheduled Activities Builder - Select Items",
            content: guiContent,
            width: "95%",
            maxWidth: "1400px",
            height: "80%",
            maxHeight: "800px",
            onClose: function() {
                // Only clear storage if user explicitly cancelled, not on auto-close
                if (SA_BUILDER_CANCELLED) {
                    clearSABuilderStorage();
                    log("SA Builder: cancelled after table refresh wait");
                }
                SA_BUILDER_POPUP_REF = null;
            }
        });
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled after table refresh wait");
            return;
        }
        log("SA Builder: selection GUI displayed");
    }
    //==========================
    // BACKGROUND HTTP REQUEST HELPERS
    //==========================
    // Helper functions for making background HTTP requests without opening tabs
    //==========================

    function fetchPage(url) {
        return new Promise(function(resolve, reject) {
            if (typeof GM !== "undefined" && typeof GM.xmlHttpRequest === "function") {
                GM.xmlHttpRequest({
                    method: "GET",
                    url: url,
                    onload: function(response) {
                        resolve(response.responseText);
                    },
                    onerror: function(error) {
                        reject(error);
                    }
                });
            } else {
                reject(new Error("GM.xmlHttpRequest not available"));
            }
        });
    }

    function parseHtml(htmlString) {
        var parser = new DOMParser();
        return parser.parseFromString(htmlString, "text/html");
    }

    function containsNumber(str) {
        return /\d/.test(str);
    }

    //==========================
    // OPEN ELIGIBILITY FEATURE
    //==========================

    function cohortEligLog(msg) {
        var ts = new Date().toISOString().substring(11, 19);
        var line = "[" + ts + "] " + String(msg);
        console.log("[OpenElig] " + line);
        log("CohortElig: " + msg);
        if (COHORT_ELIG_LOG_CONTAINER) {
            var p = document.createElement("div");
            p.style.fontSize = "12px";
            p.style.color = "#aaa";
            p.style.marginBottom = "4px";
            p.textContent = line;
            COHORT_ELIG_LOG_CONTAINER.appendChild(p);
            COHORT_ELIG_LOG_CONTAINER.scrollTop = COHORT_ELIG_LOG_CONTAINER.scrollHeight;
        }
    }

    function clearCohortEligData() {
        COHORT_ELIG_CANCELED = true;
        try { localStorage.removeItem(STORAGE_COHORT_ELIG_DATA); } catch (e) {}
        try { localStorage.removeItem(STORAGE_COHORT_ELIG_RUNNING); } catch (e) {}
        try { localStorage.removeItem(STORAGE_COHORT_ELIG_AUTO_TAB); } catch (e) {}
        log("CohortElig: All Cohort Eligibility data cleared");
        cohortEligLog("All Cohort Eligibility data cleared");
    }

    function createCohortEligSpinner() {
        var spinner = document.createElement("div");
        spinner.style.display = "inline-block";
        spinner.style.width = "20px";
        spinner.style.height = "20px";
        spinner.style.border = "3px solid #333";
        spinner.style.borderTop = "3px solid #4a90e2";
        spinner.style.borderRadius = "50%";
        spinner.style.animation = "openEligSpin 1s linear infinite";
        if (!document.getElementById("openEligSpinStyle")) {
            var style = document.createElement("style");
            style.id = "openEligSpinStyle";
            style.textContent = "@keyframes openEligSpin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }";
            document.head.appendChild(style);
        }
        return spinner;
    }

    function cohortEligibilityFeature() {
        log("CohortElig: Starting Cohort Eligibility feature");
        cohortEligLog("Cohort Eligibility: starting");
        COHORT_ELIG_CANCELED = false;

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";
        container.style.minHeight = "300px";

        var statusRow = document.createElement("div");
        statusRow.style.display = "flex";
        statusRow.style.alignItems = "center";
        statusRow.style.justifyContent = "space-between";
        statusRow.style.padding = "10px";
        statusRow.style.background = "#1a1a1a";
        statusRow.style.borderRadius = "6px";

        var spinner = createCohortEligSpinner();
        statusRow.appendChild(spinner);

        var statusText = document.createElement("span");
        statusText.textContent = "Loading epochs and cohorts...";
        statusText.style.color = "#fff";
        statusRow.appendChild(statusText);
        container.appendChild(statusRow);

        var listContainer = document.createElement("div");
        listContainer.style.flex = "1";
        listContainer.style.overflowY = "auto";
        listContainer.style.maxHeight = "400px";
        container.appendChild(listContainer);
        COHORT_ELIG_POPUP_CONTENT = listContainer;

        var logContainer = document.createElement("div");
        logContainer.style.background = "#0a0a0a";
        logContainer.style.border = "1px solid #333";
        logContainer.style.borderRadius = "6px";
        logContainer.style.padding = "8px";
        logContainer.style.maxHeight = "150px";
        logContainer.style.overflowY = "auto";
        logContainer.style.fontSize = "11px";
        container.appendChild(logContainer);
        COHORT_ELIG_LOG_CONTAINER = logContainer;

        var popup = createPopup({
            title: "Cohort Eligibility",
            content: container,
            width: "600px",
            height: "auto",
            maxHeight: "80vh",
            onClose: function() {
                cohortEligLog("Popup closed by user - clearing all data");
                clearCohortEligData();
            }
        });
        COHORT_ELIG_POPUP = popup;

        fetchEpochsAndCohorts(statusRow, statusText, spinner, listContainer);
    }

    async function fetchEpochsAndCohorts(statusRow, statusText, spinner, listContainer) {
        cohortEligLog("Fetching study show page...");

        try {
            var html = await fetchPage(COHORT_ELIG_STUDY_SHOW_URL);
            if (COHORT_ELIG_CANCELED) { cohortEligLog("Canceled"); return; }

            var doc = parseHtml(html);
            var epochTbody = doc.getElementById("epochTableBody");
            if (!epochTbody) {
                cohortEligLog("ERROR: Could not find epochTableBody");
                statusText.textContent = "Error: Could not find epochs";
                spinner.style.display = "none";
                return;
            }

            var epochRows = epochTbody.querySelectorAll("tr[id^='ir_']");
            cohortEligLog("Found " + epochRows.length + " epochs");

            var epochs = [];
            for (var i = 0; i < epochRows.length; i++) {
                var row = epochRows[i];
                var link = row.querySelector("a[href*='/epoch/show/']");
                if (link) {
                    var epochName = (link.textContent || "").trim();
                    var epochHref = link.getAttribute("href") || "";
                    epochs.push({ name: epochName, href: epochHref, cohorts: [] });
                    cohortEligLog("Epoch: " + epochName);
                }
            }

            statusText.textContent = "Fetching cohorts for " + epochs.length + " epochs...";

            for (var ei = 0; ei < epochs.length; ei++) {
                if (COHORT_ELIG_CANCELED) { cohortEligLog("Canceled"); return; }

                var epoch = epochs[ei];
                var epochUrl = "https://cenexel.clinspark.com" + epoch.href;
                cohortEligLog("Fetching cohorts for: " + epoch.name);

                try {
                    var epochHtml = await fetchPage(epochUrl);
                    var epochDoc = parseHtml(epochHtml);
                    var cohortTbody = epochDoc.getElementById("cohortListBody");

                    if (cohortTbody) {
                        var cohortRows = cohortTbody.querySelectorAll("tr[id^='ir_']");
                        for (var ci = 0; ci < cohortRows.length; ci++) {
                            var cohortRow = cohortRows[ci];
                            var cohortLink = cohortRow.querySelector("a[href*='/cohort/show/']");
                            if (cohortLink) {
                                var cohortName = (cohortLink.textContent || "").trim();
                                var cohortHref = cohortLink.getAttribute("href") || "";
                                epoch.cohorts.push({ name: cohortName, href: cohortHref });
                                cohortEligLog("  Cohort: " + cohortName);
                            }
                        }
                    }
                } catch (err) {
                    cohortEligLog("  Error: " + String(err));
                }
            }

            spinner.style.display = "none";
            statusText.textContent = "Select a cohort to open eligibility:";
            renderEpochCohortList(listContainer, epochs);

        } catch (err) {
            cohortEligLog("ERROR: " + String(err));
            statusText.textContent = "Error loading data";
            spinner.style.display = "none";
        }
    }

    function renderEpochCohortList(container, epochs) {
        container.innerHTML = "";

        for (var i = 0; i < epochs.length; i++) {
            var epoch = epochs[i];
            var epochDiv = document.createElement("div");
            epochDiv.style.marginBottom = "12px";

            var epochHeader = document.createElement("div");
            epochHeader.style.fontWeight = "600";
            epochHeader.style.fontSize = "14px";
            epochHeader.style.color = "#4a90e2";
            epochHeader.style.padding = "8px";
            epochHeader.style.background = "#1a1a1a";
            epochHeader.style.borderRadius = "4px";
            epochHeader.style.marginBottom = "4px";
            epochHeader.textContent = epoch.name;
            epochDiv.appendChild(epochHeader);

            if (epoch.cohorts.length === 0) {
                var noCohorts = document.createElement("div");
                noCohorts.style.color = "#666";
                noCohorts.style.fontSize = "12px";
                noCohorts.style.paddingLeft = "16px";
                noCohorts.textContent = "No cohorts found";
                epochDiv.appendChild(noCohorts);
            } else {
                for (var ci = 0; ci < epoch.cohorts.length; ci++) {
                    var cohort = epoch.cohorts[ci];
                    var cohortRow = document.createElement("div");
                    cohortRow.style.display = "flex";
                    cohortRow.style.alignItems = "center";
                    cohortRow.style.justifyContent = "space-between";
                    cohortRow.style.padding = "6px 8px 6px 24px";
                    cohortRow.style.borderBottom = "1px solid #333";

                    var cohortName = document.createElement("span");
                    cohortName.style.color = "#ccc";
                    cohortName.style.fontSize = "13px";
                    cohortName.textContent = cohort.name;
                    cohortRow.appendChild(cohortName);

                    var confirmBtn = document.createElement("button");
                    confirmBtn.textContent = "Confirm";
                    confirmBtn.style.background = "#28a745";
                    confirmBtn.style.color = "#fff";
                    confirmBtn.style.border = "none";
                    confirmBtn.style.borderRadius = "4px";
                    confirmBtn.style.padding = "4px 12px";
                    confirmBtn.style.cursor = "pointer";
                    confirmBtn.style.fontSize = "12px";

                    (function(cohortData, epochName, btn) {
                        btn.addEventListener("click", function() {
                            onCohortConfirm(cohortData, epochName);
                        });
                    })(cohort, epoch.name, confirmBtn);

                    cohortRow.appendChild(confirmBtn);
                    epochDiv.appendChild(cohortRow);
                }
            }
            container.appendChild(epochDiv);
        }
    }

    function onCohortConfirm(cohort, epochName) {
        log("CohortElig: Confirmed cohort: " + cohort.name + " in " + epochName);

        var stateData = {
            phase: "collectInitials",
            cohortName: cohort.name,
            cohortHref: cohort.href,
            epochName: epochName,
            initials: []
        };

        try {
            localStorage.setItem(STORAGE_COHORT_ELIG_RUNNING, "1");
            localStorage.setItem(STORAGE_COHORT_ELIG_DATA, JSON.stringify(stateData));
        } catch (e) {
            log("CohortElig: Error saving state: " + String(e));
        }

        var cohortUrl = "https://cenexel.clinspark.com" + cohort.href;
        log("CohortElig: Navigating to cohort page: " + cohortUrl);
        location.href = cohortUrl;
    }

    function processCohortEligOnCohortPage() {
        log("CohortElig: On cohort page, collecting assignments...");

        var stateRaw = null;
        try {
            stateRaw = localStorage.getItem(STORAGE_COHORT_ELIG_DATA);
        } catch (e) {}

        if (!stateRaw) {
            log("CohortElig: No state data found");
            return;
        }

        var state = null;
        try {
            state = JSON.parse(stateRaw);
        } catch (e) {
            log("CohortElig: Error parsing state: " + String(e));
            return;
        }

        var assignmentTbody = document.getElementById("cohortAssignmentListBody");
        if (!assignmentTbody) {
            log("CohortElig: cohortAssignmentListBody not found, waiting...");
            setTimeout(processCohortEligOnCohortPage, 1000);
            return;
        }

        var assignmentRows = assignmentTbody.querySelectorAll("tr.cohortAssignmentRow");
        log("CohortElig: Found " + assignmentRows.length + " cohort assignments");

        var initialsSet = [];
        for (var i = 0; i < assignmentRows.length; i++) {
            var row = assignmentRows[i];
            var volunteerCell = row.querySelector("td a[href*='/volunteers/manage/show/']");
            if (volunteerCell) {
                var volunteerText = (volunteerCell.textContent || "").trim();
                var parts = volunteerText.split(",");
                if (parts.length > 0) {
                    var initials = parts[0].trim();
                    if (initials && initialsSet.indexOf(initials) === -1) {
                        initialsSet.push(initials);
                        log("CohortElig: Collected initials: " + initials);
                    }
                }
            }
        }

        log("CohortElig: Total unique initials: " + initialsSet.length);

        state.phase = "crossReference";
        state.initials = initialsSet;

        try {
            localStorage.setItem(STORAGE_COHORT_ELIG_DATA, JSON.stringify(state));
        } catch (e) {}

        log("CohortElig: Navigating to subjects list...");
        location.href = COHORT_ELIG_SUBJECTS_LIST_URL;
    }

    function processCohortEligOnSubjectsList() {
        log("CohortElig: On subjects list, cross-referencing...");

        var stateRaw = null;
        try {
            stateRaw = localStorage.getItem(STORAGE_COHORT_ELIG_DATA);
        } catch (e) {}

        if (!stateRaw) {
            log("CohortElig: No state data found");
            return;
        }

        var state = null;
        try {
            state = JSON.parse(stateRaw);
        } catch (e) {
            log("CohortElig: Error parsing state: " + String(e));
            return;
        }

        var initialsSet = state.initials || [];
        if (initialsSet.length === 0) {
            log("CohortElig: No initials to search for");
            showCohortEligResultPopup([], state.cohortName);
            return;
        }

        var subjectTbody = document.getElementById("subjectTableBody");
        if (!subjectTbody) {
            log("CohortElig: subjectTableBody not found, waiting...");
            setTimeout(processCohortEligOnSubjectsList, 1000);
            return;
        }

        var subjectRows = subjectTbody.querySelectorAll("tr");
        log("CohortElig: Found " + subjectRows.length + " subject rows");

        var matchedSubjects = [];

        for (var i = 0; i < subjectRows.length; i++) {
            var row = subjectRows[i];
            var tds = row.querySelectorAll("td");

            if (tds.length >= 7) {
                var secondTd = tds[1];
                var volunteerLink = secondTd.querySelector("a[href*='/volunteers/manage/show/']");

                if (volunteerLink) {
                    var volunteerText = (volunteerLink.textContent || "").trim();

                    for (var j = 0; j < initialsSet.length; j++) {
                        var searchInitials = initialsSet[j];
                        if (volunteerText.indexOf(searchInitials) !== -1) {
                            log("CohortElig: MATCH found: " + searchInitials + " in " + volunteerText);

                            var showLink = null;
                            var href = "";

                            // Search through all columns for the Show link
                            for (var tdIdx = 0; tdIdx < tds.length; tdIdx++) {
                                var td = tds[tdIdx];
                                var link = td.querySelector("a[href*='/subjects/show/']");
                                if (link) {
                                    showLink = link;
                                    href = link.getAttribute("href") || "";
                                    log("CohortElig: Found Show link in column " + tdIdx + ": " + href);
                                    break;
                                }
                            }

                            if (showLink && href) {
                                var fullUrl = href;
                                if (href.indexOf("https://") !== 0) {
                                    fullUrl = "https://cenexel.clinspark.com" + href;
                                }
                                matchedSubjects.push({ initials: searchInitials, url: fullUrl, volunteerText: volunteerText });
                                log("CohortElig: Subject URL: " + fullUrl);
                            } else {
                                log("CohortElig: ERROR - Could not find Show link for " + searchInitials + " (row has " + tds.length + " columns)");
                            }
                            break;
                        }
                    }
                }
            }
        }

        log("CohortElig: Total matched subjects: " + matchedSubjects.length);
        showCohortEligResultPopup(matchedSubjects, state.cohortName);
    }

    function showCohortEligResultPopup(matchedSubjects, cohortName) {
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";

        var headerDiv = document.createElement("div");
        headerDiv.style.padding = "10px";
        headerDiv.style.background = "#1a1a1a";
        headerDiv.style.borderRadius = "6px";
        headerDiv.style.color = "#4a90e2";
        headerDiv.style.fontWeight = "600";
        headerDiv.textContent = "Cohort: " + (cohortName || "Unknown");
        container.appendChild(headerDiv);

        if (matchedSubjects.length === 0) {
            var noMatch = document.createElement("div");
            noMatch.style.padding = "20px";
            noMatch.style.color = "#f90";
            noMatch.textContent = "No matching subjects found in the subjects list.";
            container.appendChild(noMatch);
        } else {
            var statusDiv = document.createElement("div");
            statusDiv.style.padding = "12px";
            statusDiv.style.color = "#4a4";
            statusDiv.textContent = "Found " + matchedSubjects.length + " matching subjects. Opening tabs...";
            container.appendChild(statusDiv);

            var listDiv = document.createElement("div");
            listDiv.style.maxHeight = "300px";
            listDiv.style.overflowY = "auto";
            listDiv.style.fontSize = "12px";
            listDiv.style.color = "#ccc";

            for (var m = 0; m < matchedSubjects.length; m++) {
                var subj = matchedSubjects[m];
                var itemDiv = document.createElement("div");
                itemDiv.style.padding = "6px 8px";
                itemDiv.style.borderBottom = "1px solid #333";
                itemDiv.textContent = (m + 1) + ". " + subj.initials + " - " + subj.volunteerText;
                listDiv.appendChild(itemDiv);
            }
            container.appendChild(listDiv);
        }

        var popup = createPopup({
            title: "Cohort Eligibility - Results",
            content: container,
            width: "500px",
            height: "auto",
            onClose: function() {
                log("CohortElig: Results popup closed, clearing state");
                clearCohortEligData();
            }
        });

        if (matchedSubjects.length > 0) {
            openCohortMatchedSubjectTabs(matchedSubjects, container);
        } else {
            clearCohortEligData();
        }
    }

    async function openCohortMatchedSubjectTabs(matchedSubjects, container) {
        log("CohortElig: Setting auto-tab flag for " + matchedSubjects.length + " subjects");
        try {
            localStorage.setItem(STORAGE_COHORT_ELIG_AUTO_TAB, "1");
        } catch (e) {
            log("CohortElig: Error setting auto-tab flag: " + String(e));
        }

        for (var k = 0; k < matchedSubjects.length; k++) {
            var subject = matchedSubjects[k];
            log("CohortElig: Opening tab for: " + subject.initials + " -> " + subject.url);

            try {
                if (typeof GM !== "undefined" && typeof GM.openInTab === "function") {
                    var url = subject.url;
                    if (url.indexOf("https://") !== 0) {
                        url = "https://cenexel.clinspark.com" + url;
                    }
                    GM.openInTab(url, { active: false });
                } else if (typeof GM_openInTab === "function") {
                    var url = subject.url;
                    if (url.indexOf("https://") !== 0) {
                        url = "https://cenexel.clinspark.com" + url;
                    }
                    GM_openInTab(url, { active: false });
                } else {
                    var url = subject.url;
                    if (url.indexOf("https://") !== 0) {
                        url = "https://cenexel.clinspark.com" + url;
                    }
                    window.open(url, "_blank");
                }
            } catch (err) {
                log("CohortElig: Error opening tab, using window.open: " + String(err));
                var url = subject.url;
                if (url.indexOf("https://") !== 0) {
                    url = "https://cenexel.clinspark.com" + url;
                }
                window.open(url, "_blank");
            }

            await new Promise(function(resolve) { setTimeout(resolve, 300); });
        }

        log("CohortElig: Completed! Opened " + matchedSubjects.length + " tabs");

        var completeDiv = document.createElement("div");
        completeDiv.style.padding = "12px";
        completeDiv.style.color = "#4a4";
        completeDiv.style.fontWeight = "600";
        completeDiv.style.marginTop = "10px";
        completeDiv.textContent = "✓ Complete! " + matchedSubjects.length + " tabs opened.";
        container.appendChild(completeDiv);

        // Clear other data but keep auto-tab flag
        try {
            localStorage.removeItem(STORAGE_COHORT_ELIG_RUNNING);
            localStorage.removeItem(STORAGE_COHORT_ELIG_DATA);
        } catch (e) {}
        log("CohortElig: Cohort Eligibility data cleared (auto-tab flag preserved for child tabs)");

        // Clear auto-tab flag after a delay to allow all tabs to load and navigate
        setTimeout(function() {
            try {
                localStorage.removeItem(STORAGE_COHORT_ELIG_AUTO_TAB);
                log("CohortElig: Auto-tab flag cleared after delay");
            } catch (e) {}
        }, 15000);
    }

    function isCohortEligRunning() {
        var running = null;
        try {
            running = localStorage.getItem(STORAGE_COHORT_ELIG_RUNNING);
        } catch (e) {}
        return running === "1";
    }

    function getCohortEligPhase() {
        var stateRaw = null;
        try {
            stateRaw = localStorage.getItem(STORAGE_COHORT_ELIG_DATA);
        } catch (e) {}
        if (!stateRaw) return null;
        try {
            var state = JSON.parse(stateRaw);
            return state.phase;
        } catch (e) {
            return null;
        }
    }

    function isCohortShowPage() {
        return location.pathname.indexOf("/cohort/show/") !== -1;
    }

    function isSubjectsListPage() {
        return location.pathname === "/secure/study/subjects/list";
    }

    function isSubjectEligPending() {
        var pending = null;
        try {
            pending = localStorage.getItem(STORAGE_SUBJECT_ELIG_PENDING);
        } catch (e) {}
        return pending === "1";
    }

    //==========================
    // SUBJECT ELIGIBILITY FEATURE IMPLEMENTATION
    //==========================

    function subjectEligibilityFeature() {
        log("SubjectElig: Starting Subject Eligibility feature");

        var info = getSubjectIdentifierForAE();
        var scannedId = info.raw || "";

        if (scannedId && !containsNumber(scannedId)) {
            log("SubjectElig: Scanned identifier '" + scannedId + "' excluded (no number)");
            scannedId = "";
        }

        showSubjectEligInputPopup(scannedId, function(userInput) {
            if (!userInput) {
                log("SubjectElig: Canceled by user");
                return;
            }

            try {
                localStorage.setItem(STORAGE_SUBJECT_ELIG_PENDING, "1");
                localStorage.setItem(STORAGE_SUBJECT_ELIG_IDENTIFIER, userInput);
                localStorage.setItem(STORAGE_SUBJECT_ELIG_AUTO_TAB, "1");
            } catch (e) {
                log("SubjectElig: Error saving state: " + String(e));
            }

            showSubjectEligSpinnerPopup(userInput);

            log("SubjectElig: Navigating to subjects list with identifier: " + userInput);
            setTimeout(function() {
                location.href = SUBJECT_ELIG_SUBJECTS_LIST_URL;
            }, 100);
        });
    }

    function showSubjectEligSpinnerPopup(identifier) {
        log("SubjectElig: Showing spinner popup");

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.alignItems = "center";
        container.style.gap = "16px";
        container.style.padding = "20px";

        var spinner = document.createElement("div");
        spinner.style.width = "40px";
        spinner.style.height = "40px";
        spinner.style.border = "4px solid #333";
        spinner.style.borderTop = "4px solid #28a745";
        spinner.style.borderRadius = "50%";
        spinner.style.animation = "spin 1s linear infinite";

        var style = document.createElement("style");
        style.textContent = "@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }";
        document.head.appendChild(style);

        var statusText = document.createElement("div");
        statusText.style.color = "#fff";
        statusText.style.fontSize = "14px";
        statusText.style.textAlign = "center";
        statusText.textContent = "Searching for subject: " + identifier;

        container.appendChild(spinner);
        container.appendChild(statusText);

        var popup = createPopup({
            title: "Subject Eligibility - Processing",
            content: container,
            width: "350px",
            height: "auto",
            onClose: function() {
                log("SubjectElig: Spinner popup closed by user, clearing data");
                clearSubjectEligData();
            }
        });

        SUBJECT_ELIG_POPUP = popup;
    }

    function clearSubjectEligData() {
        try {
            localStorage.removeItem(STORAGE_SUBJECT_ELIG_PENDING);
            localStorage.removeItem(STORAGE_SUBJECT_ELIG_IDENTIFIER);
            localStorage.removeItem(STORAGE_SUBJECT_ELIG_AUTO_TAB);
        } catch (e) {}
        log("SubjectElig: All Subject Eligibility data cleared");
    }

    function showSubjectEligInputPopup(prefill, onDone) {
        log("SubjectElig: Showing input popup");

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";
        container.style.padding = "10px";

        var fieldRow = document.createElement("div");
        fieldRow.style.display = "grid";
        fieldRow.style.gridTemplateColumns = "140px 1fr";
        fieldRow.style.alignItems = "center";
        fieldRow.style.gap = "8px";

        var label = document.createElement("div");
        label.textContent = "Subject Identifier";
        label.style.fontWeight = "600";
        label.style.color = "#fff";

        var input = document.createElement("input");
        input.type = "text";
        input.placeholder = "e.g., 101-001, S-101-001, L-101-001";
        input.value = prefill || "";
        input.style.width = "100%";
        input.style.boxSizing = "border-box";
        input.style.padding = "8px";
        input.style.borderRadius = "6px";
        input.style.border = "1px solid #444";
        input.style.background = "#1a1a1a";
        input.style.color = "#fff";

        fieldRow.appendChild(label);
        fieldRow.appendChild(input);
        container.appendChild(fieldRow);

        var btnRow = document.createElement("div");
        btnRow.style.display = "inline-flex";
        btnRow.style.justifyContent = "flex-end";
        btnRow.style.gap = scale(BUTTON_GAP_PX);
        btnRow.style.marginTop = "10px";

        var clearIdBtn = document.createElement("button");
        clearIdBtn.textContent = "Clear ID";
        clearIdBtn.style.background = "#6c757d";
        clearIdBtn.style.color = "#fff";
        clearIdBtn.style.border = "none";
        clearIdBtn.style.borderRadius = "6px";
        clearIdBtn.style.padding = "8px 16px";
        clearIdBtn.style.cursor = "pointer";

        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Cancel";
        cancelBtn.style.background = "#333";
        cancelBtn.style.color = "#fff";
        cancelBtn.style.border = "none";
        cancelBtn.style.borderRadius = "6px";
        cancelBtn.style.padding = "8px 16px";
        cancelBtn.style.cursor = "pointer";

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#28a745";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "none";
        confirmBtn.style.borderRadius = "6px";
        confirmBtn.style.padding = "8px 16px";
        confirmBtn.style.cursor = "pointer";

        btnRow.appendChild(clearIdBtn);
        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(confirmBtn);
        container.appendChild(btnRow);

        var popup = createPopup({
            title: "Subject Eligibility",
            content: container,
            width: "450px",
            height: "auto"
        });

        SUBJECT_ELIG_POPUP = popup;

        setTimeout(function() {
            input.focus();
            input.select();
        }, 50);

        function doConfirm() {
            var value = input.value.trim();
            if (!value) {
                log("SubjectElig: Empty input");
                return;
            }

            if (!containsNumber(value)) {
                log("SubjectElig: Identifier must contain a number");
                input.style.borderColor = "#f90";
                return;
            }

            if (popup && popup.close) {
                popup.close();
            }

            if (typeof onDone === "function") {
                onDone(value);
            }
        }

        clearIdBtn.addEventListener("click", function() {
            input.value = "";
            input.style.borderColor = "#444";
            input.focus();
        });

        confirmBtn.addEventListener("click", doConfirm);

        cancelBtn.addEventListener("click", function() {
            if (popup && popup.close) {
                popup.close();
            }
            if (typeof onDone === "function") {
                onDone(null);
            }
        });

        input.addEventListener("keydown", function(e) {
            if (e.key === "Enter") {
                doConfirm();
                e.preventDefault();
            } else if (e.key === "Escape") {
                if (popup && popup.close) {
                    popup.close();
                }
                if (typeof onDone === "function") {
                    onDone(null);
                }
                e.preventDefault();
            }
        });
    }

    function processSubjectEligOnSubjectsList() {
        log("SubjectElig: On subjects list, searching for identifier...");

        var identifier = null;
        try {
            identifier = localStorage.getItem(STORAGE_SUBJECT_ELIG_IDENTIFIER);
        } catch (e) {}

        if (!identifier) {
            log("SubjectElig: No identifier found in storage");
            return;
        }

        var subjectTbody = document.getElementById("subjectTableBody");
        if (!subjectTbody) {
            log("SubjectElig: subjectTableBody not found, waiting...");
            setTimeout(processSubjectEligOnSubjectsList, 1000);
            return;
        }

        var subjectRows = subjectTbody.querySelectorAll("tr");
        log("SubjectElig: Found " + subjectRows.length + " subject rows");

        // Split identifier by " / " to get individual identifiers
        var identifierParts = identifier.split(" / ");
        log("SubjectElig: Split identifier into " + identifierParts.length + " parts: " + JSON.stringify(identifierParts));

        var matchedUrl = null;

        for (var i = 0; i < subjectRows.length; i++) {
            var row = subjectRows[i];
            var tds = row.querySelectorAll("td");

            if (tds.length > 0) {
                for (var tdIdx = 0; tdIdx < tds.length; tdIdx++) {
                    var td = tds[tdIdx];
                    var divs = td.querySelectorAll("div.vertSepNoBorder, div.vertSep");

                    for (var divIdx = 0; divIdx < divs.length; divIdx++) {
                        var div = divs[divIdx];
                        var text = (div.textContent || "").trim();
                        // Extract just the identifier value (after the label like "R:", "L:", "S:")
                        var idValue = text.replace(/^[A-Z]:\s*/, "").trim();
                        var normalizedIdValue = idValue.toUpperCase();

                        // Check if any of the identifier parts match this value
                        for (var partIdx = 0; partIdx < identifierParts.length; partIdx++) {
                            var searchPart = identifierParts[partIdx].trim().toUpperCase();

                            if (searchPart && normalizedIdValue === searchPart) {
                                log("SubjectElig: MATCH found in row " + i + ", column " + tdIdx + ": '" + searchPart + "' matches '" + idValue + "'");

                                for (var linkIdx = 0; linkIdx < tds.length; linkIdx++) {
                                    var linkTd = tds[linkIdx];
                                    var showLink = linkTd.querySelector("a[href*='/subjects/show/']");
                                    if (showLink) {
                                        var href = showLink.getAttribute("href") || "";
                                        matchedUrl = href.indexOf("https://") === 0 ? href : "https://cenexel.clinspark.com" + href;
                                        log("SubjectElig: Found show URL: " + matchedUrl);
                                        break;
                                    }
                                }

                                if (matchedUrl) break;
                            }
                        }

                        if (matchedUrl) break;
                    }

                    if (matchedUrl) break;
                }
            }

            if (matchedUrl) break;
        }

        if (matchedUrl) {
            log("SubjectElig: Navigating to: " + matchedUrl);
            try {
                localStorage.removeItem(STORAGE_SUBJECT_ELIG_PENDING);
                localStorage.removeItem(STORAGE_SUBJECT_ELIG_IDENTIFIER);
            } catch (e) {}
            location.href = matchedUrl;
        } else {
            log("SubjectElig: No match found for identifier: " + identifier);
            try {
                localStorage.removeItem(STORAGE_SUBJECT_ELIG_PENDING);
                localStorage.removeItem(STORAGE_SUBJECT_ELIG_IDENTIFIER);
                localStorage.removeItem(STORAGE_SUBJECT_ELIG_AUTO_TAB);
            } catch (e) {}

            var noMatchDiv = document.createElement("div");
            noMatchDiv.style.padding = "20px";
            noMatchDiv.style.textAlign = "center";
            noMatchDiv.style.color = "#f90";
            noMatchDiv.textContent = "No subject found with identifier: " + identifier;
            createPopup({
                title: "Subject Eligibility - No Match",
                content: noMatchDiv,
                width: "400px",
                height: "auto"
            });
        }
    }


    //==========================
    // FIND FORM FEATURE
    //==========================
    // This section contains all functions related to finding forms.
    // This feature automates pull any subject identifier found on page,
    // request user for form keyword, and then search for form based on the keyword.
    //==========================

    function applyPanelHiddenState(panel) {
        if (!panel) {
            log("applyPanelHiddenState: panel not found");
            return;
        }
        var hidden = getPanelHidden();
        if (hidden) {
            panel.style.display = "none";
            panel.style.pointerEvents = "none";
            log("applyPanelHiddenState: applied hidden");
        } else {
            panel.style.display = "block";
            panel.style.pointerEvents = "auto";
            log("applyPanelHiddenState: applied visible");
        }
    }

    function togglePanelHiddenViaHotkey() {
        var panel = document.getElementById(PANEL_ID);
        if (!panel) {
            log("Hotkey toggle: panel element not found; nothing to toggle");
            return;
        }
        var isHidden = getPanelHidden();
        if (isHidden) {
            panel.style.display = "block";
            panel.style.pointerEvents = "auto";
            setPanelHidden(false);
            log("Hotkey toggle: panel unhidden");
        } else {
            panel.style.display = "none";
            panel.style.pointerEvents = "none";
            setPanelHidden(true);
            log("Hotkey toggle: panel hidden");
        }
    }

    function normalizeKeyForMatch(e) {
        var k = "";
        var c = "";
        try {
            k = e.key ? String(e.key) : "";
        } catch (ex1) {
            k = "";
        }
        try {
            c = e.code ? String(e.code) : "";
        } catch (ex2) {
            c = "";
        }
        return { key: k, code: c, keyCode: typeof e.keyCode === "number" ? e.keyCode : null };
    }

    function keyMatchesToggle(e) {
        var n = normalizeKeyForMatch(e);
        var savedHotkey = getPanelHotkey();
        var match = false;

        if (n.key && n.key.toUpperCase() === savedHotkey.toUpperCase()) {
            match = true;
        } else if (n.code && n.code.toUpperCase() === savedHotkey.toUpperCase()) {
            match = true;
        } else if (typeof n.keyCode === "number") {
            // Fallback for F-keys using keyCode
            if (savedHotkey.toUpperCase() === "F2" && n.keyCode === 113) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F1" && n.keyCode === 112) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F3" && n.keyCode === 114) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F4" && n.keyCode === 115) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F5" && n.keyCode === 116) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F6" && n.keyCode === 117) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F7" && n.keyCode === 118) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F8" && n.keyCode === 119) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F9" && n.keyCode === 120) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F10" && n.keyCode === 121) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F11" && n.keyCode === 122) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F12" && n.keyCode === 123) {
                match = true;
            }
        }

        return match;
    }

    function bindPanelHotkeyOnce() {
        if (window.__APS_HOTKEY_BOUND === true) {
            log("Hotkey: already bound; skipping rebind");
            return;
        }
        function handler(e) {
            if (keyMatchesToggle(e)) {
                log("Hotkey: toggle request received");
                togglePanelHiddenViaHotkey();
                e.preventDefault();
                e.stopPropagation();
            }
            if (e.key === "F4" || e.key === "F4") {
                if (PARSE_DEVIATION_POPUP_REF && document.body.contains(PARSE_DEVIATION_POPUP_REF.element)) {
                    log("Hotkey: Closing Parse Deviation popup via backtick");
                    PARSE_DEVIATION_POPUP_REF.close();
                } else {
                    log("Hotkey: Opening Parse Deviation popup via backtick");
                    APS_ParseDeviation();
                }
                e.preventDefault();
                e.stopPropagation();
            }
        }
        document.addEventListener("keydown", handler, true);
        window.addEventListener("keydown", handler, true);
        window.__APS_HOTKEY_BOUND = true;
        log("Hotkey: bound for " + String(getPanelHotkey()));
    }


    function resetStudyEventsOnList() {
        var cont = document.getElementById("s2id_studyEventIds");
        if (cont) {
            clearSelect2ChoicesByContainerId("s2id_studyEventIds");
        }
        var sel = document.getElementById("studyEventIds");
        if (!sel) {
            log("Find Form: studyEventIds select not found for reset");
            return;
        }
        var ops = sel.querySelectorAll("option");
        var cleared = 0;
        if (ops && ops.length > 0) {
            var i = 0;
            while (i < ops.length) {
                if (ops[i].selected) {
                    ops[i].selected = false;
                    cleared = cleared + 1;
                }
                i = i + 1;
            }
            var ev = new Event("change", { bubbles: true });
            sel.dispatchEvent(ev);
        }
        log("Find Form: studyEventIds reset; cleared count=" + String(cleared));
    }

    function deselectAllOptionsBySelect(selectEl) {
        if (!selectEl) {
            return 0;
        }
        var ops = selectEl.querySelectorAll("option");
        var count = 0;
        if (ops && ops.length > 0) {
            var i = 0;
            while (i < ops.length) {
                if (ops[i].selected) {
                    ops[i].selected = false;
                    count = count + 1;
                }
                i = i + 1;
            }
            var ev = new Event("change", { bubbles: true });
            selectEl.dispatchEvent(ev);
        }
        log("Find Form: deselected all options count=" + String(count));
        return count;
    }

    function formMatchContainsAllTokens(text, keyword) {
        var nt = formNormalize(text || "");
        var kw = formNormalize(keyword || "");
        if (!kw || kw.length === 0) {
            return false;
        }
        var tokens = kw.split(" ");
        var all = true;
        var i = 0;
        while (i < tokens.length) {
            var tok = tokens[i];
            if (tok && tok.length > 0) {
                if (nt.indexOf(tok) < 0) {
                    all = false;
                    break;
                }
            }
            i = i + 1;
        }
        return all;
    }

    function resetStatusValuesOnList() {
        var cont = document.getElementById("s2id_statusValues");
        if (cont) {
            clearSelect2ChoicesByContainerId("s2id_statusValues");
        }
        var sel = document.getElementById("statusValues");
        if (!sel) {
            log("Find Form: status select not found for reset");
            return;
        }
        var ops = sel.querySelectorAll("option");
        if (ops && ops.length > 0) {
            var i = 0;
            while (i < ops.length) {
                ops[i].selected = false;
                i = i + 1;
            }
            var ev = new Event("change", { bubbles: true });
            sel.dispatchEvent(ev);
        }
        log("Find Form: status values reset");
    }

    function applyStatusValuesOnList(values) {
        if (!Array.isArray(values)) {
            log("Find Form: status values not array; skipping apply");
            return;
        }
        var sel = document.getElementById("statusValues");
        if (!sel) {
            log("Find Form: status select not found for apply");
            return;
        }
        var selectedCount = 0;
        var j = 0;
        while (j < values.length) {
            var v = String(values[j]);
            var op = sel.querySelector('option[value="' + v.replace(/"/g, '\\"') + '"]');
            if (op) {
                op.selected = true;
                selectedCount = selectedCount + 1;
            }
            j = j + 1;
        }
        var ev = new Event("change", { bubbles: true });
        sel.dispatchEvent(ev);
        log("Find Form: status values applied count=" + String(selectedCount));
    }

    function showFormNoMatchPopup() {
        log("Find Form: showing 'No form is found' popup");
        var box = document.createElement("div");
        box.style.textAlign = "center";
        box.style.fontSize = "16px";
        box.style.color = "#fff";
        box.style.padding = "20px";
        box.textContent = FORM_NO_MATCH_MESSAGE;
        createPopup({ title: FORM_NO_MATCH_TITLE, content: box, width: "320px", height: "auto" });
    }

    function clearSelect2ChoicesByContainerId(containerId) {
        var cont = document.getElementById(containerId);
        if (!cont) {
            log("Find Form: select2 container not found id='" + String(containerId) + "'");
            return 0;
        }
        var removed = 0;
        var tries = 0;
        while (tries < 50) {
            var closeBtn = cont.querySelector("ul.select2-choices li.select2-search-choice a.select2-search-choice-close");
            if (!closeBtn) {
                break;
            }
            try {
                closeBtn.click();
                removed = removed + 1;
            } catch (e) {}
            tries = tries + 1;
        }
        log("Find Form: cleared select2 choices id='" + String(containerId) + "' count=" + String(removed));
        return removed;
    }

    function findSubjectIdentifierForFindForm() {
        var s = "";
        var cont = document.getElementById("s2id_subjectIds");
        if (cont) {
            var choiceDiv = cont.querySelector("ul.select2-choices li.select2-search-choice div");
            if (choiceDiv) {
                var t = (choiceDiv.textContent + "").trim();
                if (t.length > 0) {
                    s = t;
                    log("Find Form: subject from select2 choice='" + String(s) + "'");
                }
            }
        }
        if (!s || s.length === 0) {
            var bc = document.querySelector("ul.page-breadcrumb.breadcrumb > li:nth-child(3)");
            if (bc) {
                var bt = (bc.textContent + "").trim();
                if (bt.length > 0) {
                    var parts = bt.split("/");
                    if (parts && parts.length > 0) {
                        var p0 = (parts[0] + "").trim();
                        if (p0.length > 0) {
                            s = p0;
                            log("Find Form: subject from breadcrumb='" + String(s) + "'");
                        }
                    }
                }
            }
        }
        if (!s || s.length === 0) {
            var info = getSubjectIdentifierForAE();
            if (info && info.raw && info.raw.length > 0) {
                s = info.raw;
                log("Find Form: subject from AE fallback='" + String(s) + "'");
            }
        }
        if (s && s.length > 0 && !/\d/.test(s)) {
            log("Find Form: identifier '" + String(s) + "' excluded (no number)");
            s = "";
        }
        return s;
    }

    function getSubjectIdentifierForAE() {
        var normalized = "";
        var rawCandidate = "";
        var confident = false;

        var items = document.querySelectorAll(".page-breadcrumb.breadcrumb > li");
        if (items && items.length >= 3) {
            var el3 = items[2];
            var t3 = "";
            if (el3) {
                t3 = el3.textContent + "";
            }
            var n3 = aeNormalize(t3);
            if (n3.length > 0 && containsNumber(t3)) {
                rawCandidate = t3;
                normalized = n3;
            }
        }

        if ((!normalized || normalized.length === 0) && items) {
            var c = 0;
            while (c < items.length) {
                var li = items[c];
                var tt = "";
                if (li) {
                    tt = (li.textContent + "").trim();
                }
                var hasSlash = tt.indexOf("/") >= 0;
                if (hasSlash && containsNumber(tt)) {
                    rawCandidate = tt;
                    normalized = aeNormalize(tt);
                    if (normalized.length > 0) {
                        break;
                    }
                }
                c = c + 1;
            }
        }

        if (!normalized || normalized.length === 0) {
            var cap = document.querySelector("div.portlet-title div.caption");
            if (cap) {
                var ct = (cap.textContent + "").trim();
                var nc = aeNormalize(ct);
                if (nc.length > 0 && containsNumber(ct)) {
                    rawCandidate = ct;
                    normalized = nc;
                }
            }
        }

        var sid = "";
        var a = document.querySelector('span.tooltips[data-original-title="Subject"] a[data-target="#ajaxModal"][data-toggle="modal"]');
        if (a) {
            var href = a.getAttribute("href") + "";
            var m = href.match(/\/show\/subject\/(\d+)\//);
            if (m && m[1]) {
                sid = m[1];
            }
        }

        if (sid && sid.length > 0) {
            confident = true;
        } else {
            if (normalized && normalized.length > 0) {
                var looksNumeric = /^[0-9]+$/.test(rawCandidate.trim());
                var hasIdPattern = /^S?[0-9]{3,}$/.test(rawCandidate.trim());
                if (looksNumeric || hasIdPattern) {
                    confident = true;
                } else {
                    confident = false;
                }
            } else {
                confident = false;
            }
        }

        log("Find AE: subject normalized='" + String(normalized) + "' subjectId='" + String(sid) + "' confident=" + String(confident) + " raw='" + String(rawCandidate) + "'");
        return { normalizedIdentifier: normalized, subjectId: sid, confident: confident, raw: rawCandidate };
    }

    function previewFormsByKeyword(keyword, done) {
        log("Find Form: preview request started");
        var url = FORM_LIST_URL;
        try {
            GM.xmlHttpRequest({
                method: "GET",
                url: url,
                onload: function (resp) {
                    var html = resp && resp.responseText ? resp.responseText + "" : "";
                    var tmp = document.createElement("div");
                    tmp.innerHTML = html;
                    var formsSel = tmp.querySelector('select#formIds, select[name="formIds"]');
                    var matched = 0;
                    if (formsSel) {
                        var fos = formsSel.querySelectorAll("option");
                        if (fos && fos.length > 0) {
                            var i = 0;
                            while (i < fos.length) {
                                var fop = fos[i];
                                var ftxt = "";
                                if (fop) {
                                    ftxt = fop.textContent + "";
                                }
                                var hit = formMatchContainsAllTokens(ftxt, keyword);
                                if (hit) {
                                    matched = matched + 1;
                                }
                                i = i + 1;
                            }
                        }
                    }
                    log("Find Form: preview matched count=" + String(matched));
                    if (typeof done === "function") {
                        done(matched > 0);
                    }
                },
                onerror: function () {
                    log("Find Form: preview request error");
                    if (typeof done === "function") {
                        done(true);
                    }
                }
            });
        } catch (e) {
            log("Find Form: preview exception " + String(e));
            if (typeof done === "function") {
                done(true);
            }
        }
    }


    function formNormalize(x) {
        var s = x || "";
        s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        s = s.replace(/\s+/g, " ");
        s = s.trim();
        s = s.toUpperCase();
        return s;
    }


    function showFindFormPopup(prefillSubject, onDone) {
        log("Find Form: showing popup");
        var container = document.createElement("div");
        container.style.display = "grid";
        container.style.gridTemplateRows = "auto auto auto auto auto";
        container.style.gap = "10px";

        var row1 = document.createElement("div");
        row1.style.display = "grid";
        row1.style.gridTemplateColumns = "140px 1fr";
        row1.style.alignItems = "center";
        row1.style.gap = "8px";
        var label1 = document.createElement("div");
        label1.textContent = FORM_POPUP_KEYWORD_LABEL;
        label1.style.fontWeight = "600";
        var inputKeyword = document.createElement("input");
        inputKeyword.type = "text";
        inputKeyword.placeholder = "Required: any word";
        inputKeyword.style.width = "100%";
        inputKeyword.style.boxSizing = "border-box";
        inputKeyword.style.padding = "8px";
        inputKeyword.style.borderRadius = "6px";
        inputKeyword.style.border = "1px solid #444";
        inputKeyword.style.background = "#1a1a1a";
        inputKeyword.style.color = "#fff";
        row1.appendChild(label1);
        row1.appendChild(inputKeyword);
        container.appendChild(row1);

        var row2 = document.createElement("div");
        row2.style.display = "grid";
        row2.style.gridTemplateColumns = "140px 1fr";
        row2.style.alignItems = "center";
        row2.style.gap = "8px";
        var label2 = document.createElement("div");
        label2.textContent = FORM_POPUP_SUBJECT_LABEL;
        label2.style.fontWeight = "600";
        var inputSubject = document.createElement("input");
        inputSubject.type = "text";
        inputSubject.placeholder = "Optional: subject id or label";
        inputSubject.style.width = "100%";
        inputSubject.style.boxSizing = "border-box";
        inputSubject.style.padding = "8px";
        inputSubject.style.borderRadius = "6px";
        inputSubject.style.border = "1px solid #444";
        inputSubject.style.background = "#1a1a1a";
        inputSubject.style.color = "#fff";
        row2.appendChild(label2);
        row2.appendChild(inputSubject);
        container.appendChild(row2);

        var row3 = document.createElement("div");
        row3.style.display = "grid";
        row3.style.gridTemplateColumns = "140px 1fr";
        row3.style.alignItems = "center";
        row3.style.gap = "8px";
        var labelIG = document.createElement("div");
        labelIG.textContent = "Item Group Data";
        labelIG.style.fontWeight = "600";
        var igWrap = document.createElement("div");
        igWrap.style.display = "inline-flex";
        igWrap.style.gap = "12px";
        var igC = document.createElement("label");
        var igCBox = document.createElement("input");
        igCBox.type = "checkbox";
        igCBox.value = "Complete";
        igC.appendChild(igCBox);
        igC.appendChild(document.createTextNode(" Complete"));
        var igI = document.createElement("label");
        var igIBox = document.createElement("input");
        igIBox.type = "checkbox";
        igIBox.value = "Incomplete";
        igI.appendChild(igIBox);
        igI.appendChild(document.createTextNode(" Incomplete"));
        var igN = document.createElement("label");
        var igNBox = document.createElement("input");
        igNBox.type = "checkbox";
        igNBox.value = "Nonconformant";
        igN.appendChild(igNBox);
        igN.appendChild(document.createTextNode(" Nonconformant"));
        igWrap.appendChild(igC);
        igWrap.appendChild(igI);
        igWrap.appendChild(igN);
        row3.appendChild(labelIG);
        row3.appendChild(igWrap);
        container.appendChild(row3);

        var row4 = document.createElement("div");
        row4.style.display = "grid";
        row4.style.gridTemplateColumns = "140px 1fr";
        row4.style.alignItems = "center";
        row4.style.gap = "8px";
        var labelFD = document.createElement("div");
        labelFD.textContent = "Form Data";
        labelFD.style.fontWeight = "600";
        var fdWrap = document.createElement("div");
        fdWrap.style.display = "inline-flex";
        fdWrap.style.gap = "12px";
        var fdNC = document.createElement("label");
        var fdNCBox = document.createElement("input");
        fdNCBox.type = "checkbox";
        fdNCBox.value = "formDataNotCanceled";
        fdNC.appendChild(fdNCBox);
        fdNC.appendChild(document.createTextNode(" Not Canceled"));
        var fdTM = document.createElement("label");
        var fdTMBox = document.createElement("input");
        fdTMBox.type = "checkbox";
        fdTMBox.value = "timedAndMissed";
        fdTM.appendChild(fdTMBox);
        fdTM.appendChild(document.createTextNode(" Time and Missed"));
        fdWrap.appendChild(fdNC);
        fdWrap.appendChild(fdTM);
        row4.appendChild(labelFD);
        row4.appendChild(fdWrap);
        container.appendChild(row4);

        try {
            var prevRaw = localStorage.getItem(STORAGE_FIND_FORM_STATUS_VALUES);
            if (prevRaw) {
                var prevArr = JSON.parse(prevRaw);
                if (Array.isArray(prevArr) && prevArr.length > 0) {
                    var hasComplete = prevArr.indexOf("Complete") >= 0;
                    var hasIncomplete = prevArr.indexOf("Incomplete") >= 0;
                    var hasNonconf = prevArr.indexOf("Nonconformant") >= 0;
                    var hasFormNotCanceled = prevArr.indexOf("formDataNotCanceled") >= 0;
                    var hasTimedMissed = prevArr.indexOf("timedAndMissed") >= 0;
                    igCBox.checked = !!hasComplete;
                    igIBox.checked = !!hasIncomplete;
                    igNBox.checked = !!hasNonconf;
                    fdNCBox.checked = !!hasFormNotCanceled;
                    fdTMBox.checked = !!hasTimedMissed;
                    log("Find Form: restored checkbox prefs count=" + String(prevArr.length));
                } else {
                    log("Find Form: no prior checkbox prefs to restore");
                }
            } else {
                log("Find Form: checkbox prefs not present in storage");
            }
        } catch (e) {
            log("Find Form: error restoring checkbox prefs");
        }

        if (prefillSubject && prefillSubject.length > 0) {
            inputSubject.value = prefillSubject;
            var ev0 = new Event("input", { bubbles: true });
            inputSubject.dispatchEvent(ev0);
            log("Find Form: subject prefilled='" + String(prefillSubject) + "'");
        }

        var btnRow = document.createElement("div");
        btnRow.style.display = "inline-flex";
        btnRow.style.justifyContent = "flex-end";
        btnRow.style.gap = "8px";
        var clearIdBtn = document.createElement("button");
        clearIdBtn.textContent = "Clear ID";
        clearIdBtn.style.background = "#777";
        clearIdBtn.style.color = "#fff";
        clearIdBtn.style.border = "none";
        clearIdBtn.style.borderRadius = "6px";
        clearIdBtn.style.padding = "8px 12px";
        clearIdBtn.style.cursor = "pointer";
        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = FORM_POPUP_CANCEL_TEXT;
        cancelBtn.style.background = "#333";
        cancelBtn.style.color = "#fff";
        cancelBtn.style.border = "none";
        cancelBtn.style.borderRadius = "6px";
        cancelBtn.style.padding = "8px 12px";
        cancelBtn.style.cursor = "pointer";
        var okBtn = document.createElement("button");
        okBtn.textContent = FORM_POPUP_OK_TEXT;
        okBtn.style.background = "#0b82ff";
        okBtn.style.color = "#fff";
        okBtn.style.border = "none";
        okBtn.style.borderRadius = "6px";
        okBtn.style.padding = "8px 12px";
        okBtn.style.cursor = "pointer";
        btnRow.appendChild(clearIdBtn);
        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(okBtn);
        container.appendChild(btnRow);

        var popup = createPopup({ title: FORM_POPUP_TITLE , content: container, width: "520px", height: "auto" });

        window.setTimeout(function () {
            try {
                inputKeyword.focus();
                inputKeyword.select();
                log("Find Form: keyword input focused");
            } catch (e) {
                log("Find Form: failed to focus keyword input");
            }
        }, 50);

        function gatherStatusSelections() {
            var out = [];
            if (igCBox && igCBox.checked) {
                out.push("Complete");
            }
            if (igIBox && igIBox.checked) {
                out.push("Incomplete");
            }
            if (igNBox && igNBox.checked) {
                out.push("Nonconformant");
            }
            if (fdNCBox && fdNCBox.checked) {
                out.push("formDataNotCanceled");
            }
            if (fdTMBox && fdTMBox.checked) {
                out.push("timedAndMissed");
            }
            return out;
        }

        function doContinue() {
            var kw = inputKeyword.value + "";
            var sbj = inputSubject.value + "";
            var kwt = kw.trim();
            var sbjt = sbj.trim();
            if (!kwt || kwt.length === 0) {
                log("Find Form: keyword required");
                return;
            }
            var statuses = gatherStatusSelections();
            try {
                if (statuses && statuses.length > 0) {
                    localStorage.setItem(STORAGE_FIND_FORM_STATUS_VALUES, JSON.stringify(statuses));
                    log("Find Form: statuses saved count=" + String(statuses.length));
                } else {
                    localStorage.removeItem(STORAGE_FIND_FORM_STATUS_VALUES);
                    log("Find Form: no statuses selected; cleared storage");
                }
            } catch (e) {}
            log("Find Form: popup inputs keyword='" + String(kwt) + "' subject='" + String(sbjt) + "'");
            if (popup && popup.close) {
                popup.close();
            }
            if (typeof onDone === "function") {
                onDone({ keyword: kwt, subject: sbjt });
            }
        }

        function keyHandler(e) {
            var code = e.key || e.code || "";
            if (code === "Enter") {
                log("Find Form: Enter pressed; continuing");
                doContinue();
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            if (code === "Escape" || code === "Esc") {
                log("Find Form: Esc pressed; closing");
                if (popup && popup.close) {
                    popup.close();
                }
                document.removeEventListener("keydown", keyHandler, true);
                if (typeof onDone === "function") {
                    onDone(null);
                }
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        document.addEventListener("keydown", keyHandler, true);

        clearIdBtn.addEventListener("click", function () {
            inputSubject.value = "";
            var ev = new Event("input", { bubbles: true });
            inputSubject.dispatchEvent(ev);
            log("Find Form: subject cleared by user");
        });

        cancelBtn.addEventListener("click", function () {
            log("Find Form: popup canceled");
            if (popup && popup.close) {
                popup.close();
            }
            document.removeEventListener("keydown", keyHandler, true);
            if (typeof onDone === "function") {
                onDone(null);
            }
        });

        okBtn.addEventListener("click", function () {
            document.removeEventListener("keydown", keyHandler, true);
            doContinue();
        });
    }

    function openFindForm() {
        log("Find Form: starting");
        var prefill = findSubjectIdentifierForFindForm();
        showFindFormPopup(prefill, function (userInput) {
            if (userInput === null) {
                log("Find Form: canceled by user; stopping");
                return;
            }
            var kw = userInput.keyword || "";
            var sbj = userInput.subject || "";
            previewFormsByKeyword(kw, function (hasMatch) {
                if (!hasMatch) {
                    log("Find Form: preview found no matching forms; stopping");
                    showFormNoMatchPopup();
                    return;
                }
                localStorage.setItem(STORAGE_FIND_FORM_PENDING, "1");
                localStorage.setItem(STORAGE_FIND_FORM_KEYWORD, String(kw));
                localStorage.setItem(STORAGE_FIND_FORM_SUBJECT, String(sbj));
                log("Find Form: state saved; redirecting same tab");
                window.location.href = FORM_LIST_URL;
            });
        });
    }

    function processFindFormOnList() {
        var pending = localStorage.getItem(STORAGE_FIND_FORM_PENDING);
        if (!pending || pending !== "1") {
            return;
        }
        var kw = localStorage.getItem(STORAGE_FIND_FORM_KEYWORD) || "";
        var sbj = localStorage.getItem(STORAGE_FIND_FORM_SUBJECT) || "";
        log("Find Form: post-reload processing started keyword='" + String(kw) + "' subject='" + String(sbj) + "'");
        var t = 0;
        var r = setInterval(function () {
            t = t + 1;
            if (t > 120) {
                clearInterval(r);
                log("Find Form: timeout waiting for controls");
                localStorage.removeItem(STORAGE_FIND_FORM_PENDING);
                return;
            }
            var d = document;
            var formsSel = d.querySelector('select#formIds, select[name="formIds"]');
            var subjSel = d.querySelector('select#subjectIds, select[name="subjectIds"]');
            var formsBoxReady = document.getElementById("s2id_formIds");
            var statusSel = d.getElementById("statusValues");
            var statusBoxReady = d.getElementById("s2id_statusValues");
            if (!formsSel || !formsBoxReady) {
                log("Find Form: selects not ready at t=" + String(t));
                return;
            }

            resetStudyEventsOnList();
            resetStatusValuesOnList();

            var rawStatuses = null;
            try {
                rawStatuses = localStorage.getItem(STORAGE_FIND_FORM_STATUS_VALUES);
            } catch (e) {}
            if (rawStatuses) {
                try {
                    var arr = JSON.parse(rawStatuses);
                    if (Array.isArray(arr) && arr.length > 0) {
                        applyStatusValuesOnList(arr);
                    } else {
                        log("Find Form: no statuses to apply");
                    }
                } catch (e2) {
                    log("Find Form: status parse error");
                }
            } else {
                log("Find Form: no stored statuses");
            }

            clearSelect2ChoicesByContainerId("s2id_formIds");
            deselectAllOptionsBySelect(formsSel);

            var selectedCount = 0;
            var fos = formsSel.querySelectorAll("option");
            if (fos && fos.length > 0) {
                var i = 0;
                while (i < fos.length) {
                    var fop = fos[i];
                    var ftxt = "";
                    if (fop) {
                        ftxt = fop.textContent + "";
                    }
                    var hit = formMatchContainsAllTokens(ftxt, kw);
                    if (hit) {
                        fop.selected = true;
                        selectedCount = selectedCount + 1;
                    }
                    i = i + 1;
                }
                var evtForms = new Event("change", { bubbles: true });
                formsSel.dispatchEvent(evtForms);
                log("Find Form: selected forms count=" + String(selectedCount));
            } else {
                log("Find Form: forms options list empty");
            }

            if (selectedCount === 0) {
                clearInterval(r);
                log("Find Form: no forms selected; notifying user");
                showFormNoMatchPopup();
                localStorage.removeItem(STORAGE_FIND_FORM_PENDING);
                localStorage.removeItem(STORAGE_FIND_FORM_KEYWORD);
                localStorage.removeItem(STORAGE_FIND_FORM_SUBJECT);
                return;
            }

            var subjNorm = aeNormalize(sbj || "");
            if (subjNorm && subjNorm.length > 0 && subjSel) {
                var os = subjSel.querySelectorAll("option");
                var y = "";
                if (os && os.length > 0) {
                    var j = 0;
                    while (j < os.length) {
                        var op = os[j];
                        var txt = "";
                        var val = "";
                        if (op) {
                            txt = op.textContent + "";
                            val = op.value + "";
                        }
                        var nt = aeNormalize(txt);
                        var matchByTextExact = nt === subjNorm;
                        var matchByTextContains = nt.indexOf(subjNorm) >= 0;
                        var matchById = val === sbj;
                        var matched = false;
                        if (matchByTextExact) {
                            matched = true;
                        } else {
                            if (matchByTextContains) {
                                matched = true;
                            } else {
                                if (matchById) {
                                    matched = true;
                                }
                            }
                        }
                        if (matched) {
                            y = val;
                            break;
                        }
                        j = j + 1;
                    }
                }
                if (y && y.length > 0) {
                    subjSel.value = y;
                    var evtSubj = new Event("change", { bubbles: true });
                    subjSel.dispatchEvent(evtSubj);
                    log("Find Form: subject applied value='" + String(y) + "'");
                } else {
                    log("Find Form: no subject match applied");
                }
            } else {
                log("Find Form: subject input empty or select not found");
            }

            var searchBtn = d.getElementById("dataSearchButton");
            if (searchBtn) {
                searchBtn.click();
                log("Find Form: Search button clicked");
            } else {
                log("Find Form: Search button not found");
            }

            clearInterval(r);
            localStorage.removeItem(STORAGE_FIND_FORM_PENDING);
            localStorage.removeItem(STORAGE_FIND_FORM_KEYWORD);
            localStorage.removeItem(STORAGE_FIND_FORM_SUBJECT);
            log("Find Form: post-reload processing completed");
        }, 200);
    }

    function aeNormalize(x) {
        var s = x || "";
        s = s.replace(/\s+/g, "");
        s = s.replace(/›|»|▶|►|→/g, "");
        s = s.trim();
        s = s.toUpperCase();
        return s;
    }

    //==========================
    // FIND STUDY EVENTS FEATURE
    //==========================
    // This section contains all functions related to finding study events.
    // This feature automates pull any subject identifier found on page,
    // request user for study event keyword, and then search for study events based on the keyword.
    //==========================

    function showStudyEventNoMatchPopup() {
        log("Find Study Events: showing 'No study event is found' popup");
        var box = document.createElement("div");
        box.style.textAlign = "center";
        box.style.fontSize = "16px";
        box.style.color = "#fff";
        box.style.padding = "20px";
        box.textContent = STUDY_EVENT_NO_MATCH_MESSAGE;
        createPopup({ title: STUDY_EVENT_NO_MATCH_TITLE, content: box, width: "320px", height: "auto" });
    }

    function studyEventMatchContainsAllTokens(text, keyword) {
        var nt = formNormalize(text || "");
        var kw = formNormalize(keyword || "");
        if (!kw || kw.length === 0) {
            return false;
        }
        var tokens = kw.split(" ");
        var all = true;
        var i = 0;
        while (i < tokens.length) {
            var tok = tokens[i];
            if (tok && tok.length > 0) {
                if (nt.indexOf(tok) < 0) {
                    all = false;
                    break;
                }
            }
            i = i + 1;
        }
        return all;
    }

    function previewStudyEventsByKeyword(keyword, done) {
        log("Find Study Events: preview request started");
        var url = FORM_LIST_URL;
        try {
            GM.xmlHttpRequest({
                method: "GET",
                url: url,
                onload: function (resp) {
                    var html = resp && resp.responseText ? resp.responseText + "" : "";
                    var tmp = document.createElement("div");
                    tmp.innerHTML = html;
                    var eventsSel = tmp.querySelector('select#studyEventIds, select[name="studyEventIds"]');
                    var matched = 0;
                    if (eventsSel) {
                        var eos = eventsSel.querySelectorAll("option");
                        if (eos && eos.length > 0) {
                            var i = 0;
                            while (i < eos.length) {
                                var eop = eos[i];
                                var etxt = "";
                                if (eop) {
                                    etxt = eop.textContent + "";
                                }
                                var hit = studyEventMatchContainsAllTokens(etxt, keyword);
                                if (hit) {
                                    matched = matched + 1;
                                }
                                i = i + 1;
                            }
                        }
                    }
                    log("Find Study Events: preview matched count=" + String(matched));
                    if (typeof done === "function") {
                        done(matched > 0);
                    }
                },
                onerror: function () {
                    log("Find Study Events: preview request error");
                    if (typeof done === "function") {
                        done(true);
                    }
                }
            });
        } catch (e) {
            log("Find Study Events: preview exception " + String(e));
            if (typeof done === "function") {
                done(true);
            }
        }
    }

    function showFindStudyEventsPopup(prefillSubject, onDone) {
        log("Find Study Events: showing popup");
        var container = document.createElement("div");
        container.style.display = "grid";
        container.style.gridTemplateRows = "auto auto auto auto auto";
        container.style.gap = "10px";

        var row1 = document.createElement("div");
        row1.style.display = "grid";
        row1.style.gridTemplateColumns = "140px 1fr";
        row1.style.alignItems = "center";
        row1.style.gap = "8px";
        var label1 = document.createElement("div");
        label1.textContent = STUDY_EVENT_POPUP_KEYWORD_LABEL;
        label1.style.fontWeight = "600";
        var inputKeyword = document.createElement("input");
        inputKeyword.type = "text";
        inputKeyword.placeholder = "Required: any word";
        inputKeyword.style.width = "100%";
        inputKeyword.style.boxSizing = "border-box";
        inputKeyword.style.padding = "8px";
        inputKeyword.style.borderRadius = "6px";
        inputKeyword.style.border = "1px solid #444";
        inputKeyword.style.background = "#1a1a1a";
        inputKeyword.style.color = "#fff";
        row1.appendChild(label1);
        row1.appendChild(inputKeyword);
        container.appendChild(row1);

        var row2 = document.createElement("div");
        row2.style.display = "grid";
        row2.style.gridTemplateColumns = "140px 1fr";
        row2.style.alignItems = "center";
        row2.style.gap = "8px";
        var label2 = document.createElement("div");
        label2.textContent = STUDY_EVENT_POPUP_SUBJECT_LABEL;
        label2.style.fontWeight = "600";
        var inputSubject = document.createElement("input");
        inputSubject.type = "text";
        inputSubject.placeholder = "Optional: subject id or label";
        inputSubject.style.width = "100%";
        inputSubject.style.boxSizing = "border-box";
        inputSubject.style.padding = "8px";
        inputSubject.style.borderRadius = "6px";
        inputSubject.style.border = "1px solid #444";
        inputSubject.style.background = "#1a1a1a";
        inputSubject.style.color = "#fff";
        row2.appendChild(label2);
        row2.appendChild(inputSubject);
        container.appendChild(row2);

        var row3 = document.createElement("div");
        row3.style.display = "grid";
        row3.style.gridTemplateColumns = "140px 1fr";
        row3.style.alignItems = "center";
        row3.style.gap = "8px";
        var labelIG = document.createElement("div");
        labelIG.textContent = "Item Group Data";
        labelIG.style.fontWeight = "600";
        var igWrap = document.createElement("div");
        igWrap.style.display = "inline-flex";
        igWrap.style.gap = "12px";
        var igC = document.createElement("label");
        var igCBox = document.createElement("input");
        igCBox.type = "checkbox";
        igCBox.value = "Complete";
        igC.appendChild(igCBox);
        igC.appendChild(document.createTextNode(" Complete"));
        var igI = document.createElement("label");
        var igIBox = document.createElement("input");
        igIBox.type = "checkbox";
        igIBox.value = "Incomplete";
        igI.appendChild(igIBox);
        igI.appendChild(document.createTextNode(" Incomplete"));
        var igN = document.createElement("label");
        var igNBox = document.createElement("input");
        igNBox.type = "checkbox";
        igNBox.value = "Nonconformant";
        igN.appendChild(igNBox);
        igN.appendChild(document.createTextNode(" Nonconformant"));
        igWrap.appendChild(igC);
        igWrap.appendChild(igI);
        igWrap.appendChild(igN);
        row3.appendChild(labelIG);
        row3.appendChild(igWrap);
        container.appendChild(row3);

        var row4 = document.createElement("div");
        row4.style.display = "grid";
        row4.style.gridTemplateColumns = "140px 1fr";
        row4.style.alignItems = "center";
        row4.style.gap = "8px";
        var labelFD = document.createElement("div");
        labelFD.textContent = "Form Data";
        labelFD.style.fontWeight = "600";
        var fdWrap = document.createElement("div");
        fdWrap.style.display = "inline-flex";
        fdWrap.style.gap = "12px";
        var fdNC = document.createElement("label");
        var fdNCBox = document.createElement("input");
        fdNCBox.type = "checkbox";
        fdNCBox.value = "formDataNotCanceled";
        fdNC.appendChild(fdNCBox);
        fdNC.appendChild(document.createTextNode(" Not Canceled"));
        var fdTM = document.createElement("label");
        var fdTMBox = document.createElement("input");
        fdTMBox.type = "checkbox";
        fdTMBox.value = "timedAndMissed";
        fdTM.appendChild(fdTMBox);
        fdTM.appendChild(document.createTextNode(" Time and Missed"));
        fdWrap.appendChild(fdNC);
        fdWrap.appendChild(fdTM);
        row4.appendChild(labelFD);
        row4.appendChild(fdWrap);
        container.appendChild(row4);

        try {
            var prevRaw = localStorage.getItem(STORAGE_FIND_STUDY_EVENT_STATUS_VALUES);
            if (prevRaw) {
                var prevArr = JSON.parse(prevRaw);
                if (Array.isArray(prevArr) && prevArr.length > 0) {
                    var hasComplete = prevArr.indexOf("Complete") >= 0;
                    var hasIncomplete = prevArr.indexOf("Incomplete") >= 0;
                    var hasNonconf = prevArr.indexOf("Nonconformant") >= 0;
                    var hasFormNotCanceled = prevArr.indexOf("formDataNotCanceled") >= 0;
                    var hasTimedMissed = prevArr.indexOf("timedAndMissed") >= 0;
                    igCBox.checked = !!hasComplete;
                    igIBox.checked = !!hasIncomplete;
                    igNBox.checked = !!hasNonconf;
                    fdNCBox.checked = !!hasFormNotCanceled;
                    fdTMBox.checked = !!hasTimedMissed;
                    log("Find Study Events: restored checkbox prefs count=" + String(prevArr.length));
                } else {
                    log("Find Study Events: no prior checkbox prefs to restore");
                }
            } else {
                log("Find Study Events: checkbox prefs not present in storage");
            }
        } catch (e) {
            log("Find Study Events: error restoring checkbox prefs");
        }

        if (prefillSubject && prefillSubject.length > 0) {
            inputSubject.value = prefillSubject;
            var ev0 = new Event("input", { bubbles: true });
            inputSubject.dispatchEvent(ev0);
            log("Find Study Events: subject prefilled='" + String(prefillSubject) + "'");
        }

        var btnRow = document.createElement("div");
        btnRow.style.display = "inline-flex";
        btnRow.style.justifyContent = "flex-end";
        btnRow.style.gap = "8px";
        var clearIdBtn = document.createElement("button");
        clearIdBtn.textContent = "Clear ID";
        clearIdBtn.style.background = "#777";
        clearIdBtn.style.color = "#fff";
        clearIdBtn.style.border = "none";
        clearIdBtn.style.borderRadius = "6px";
        clearIdBtn.style.padding = "8px 12px";
        clearIdBtn.style.cursor = "pointer";
        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = STUDY_EVENT_POPUP_CANCEL_TEXT;
        cancelBtn.style.background = "#333";
        cancelBtn.style.color = "#fff";
        cancelBtn.style.border = "none";
        cancelBtn.style.borderRadius = "6px";
        cancelBtn.style.padding = "8px 12px";
        cancelBtn.style.cursor = "pointer";
        var okBtn = document.createElement("button");
        okBtn.textContent = STUDY_EVENT_POPUP_OK_TEXT;
        okBtn.style.background = "#0b82ff";
        okBtn.style.color = "#fff";
        okBtn.style.border = "none";
        okBtn.style.borderRadius = "6px";
        okBtn.style.padding = "8px 12px";
        okBtn.style.cursor = "pointer";
        btnRow.appendChild(clearIdBtn);
        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(okBtn);
        container.appendChild(btnRow);

        var popup = createPopup({ title: STUDY_EVENT_POPUP_TITLE, content: container, width: "520px", height: "auto" });

        window.setTimeout(function () {
            try {
                inputKeyword.focus();
                inputKeyword.select();
                log("Find Study Events: keyword input focused");
            } catch (e) {
                log("Find Study Events: failed to focus keyword input");
            }
        }, 50);

        function gatherStatusSelections() {
            var out = [];
            if (igCBox && igCBox.checked) {
                out.push("Complete");
            }
            if (igIBox && igIBox.checked) {
                out.push("Incomplete");
            }
            if (igNBox && igNBox.checked) {
                out.push("Nonconformant");
            }
            if (fdNCBox && fdNCBox.checked) {
                out.push("formDataNotCanceled");
            }
            if (fdTMBox && fdTMBox.checked) {
                out.push("timedAndMissed");
            }
            return out;
        }

        function doContinue() {
            var kw = inputKeyword.value + "";
            var sbj = inputSubject.value + "";
            var kwt = kw.trim();
            var sbjt = sbj.trim();
            if (!kwt || kwt.length === 0) {
                log("Find Study Events: keyword required");
                return;
            }
            var statuses = gatherStatusSelections();
            try {
                if (statuses && statuses.length > 0) {
                    localStorage.setItem(STORAGE_FIND_STUDY_EVENT_STATUS_VALUES, JSON.stringify(statuses));
                    log("Find Study Events: statuses saved count=" + String(statuses.length));
                } else {
                    localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_STATUS_VALUES);
                    log("Find Study Events: no statuses selected; cleared storage");
                }
            } catch (e) {}
            log("Find Study Events: popup inputs keyword='" + String(kwt) + "' subject='" + String(sbjt) + "'");
            if (popup && popup.close) {
                popup.close();
            }
            if (typeof onDone === "function") {
                onDone({ keyword: kwt, subject: sbjt });
            }
        }

        function keyHandler(e) {
            var code = e.key || e.code || "";
            if (code === "Enter") {
                log("Find Study Events: Enter pressed; continuing");
                doContinue();
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            if (code === "Escape" || code === "Esc") {
                log("Find Study Events: Esc pressed; closing");
                if (popup && popup.close) {
                    popup.close();
                }
                document.removeEventListener("keydown", keyHandler, true);
                if (typeof onDone === "function") {
                    onDone(null);
                }
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        document.addEventListener("keydown", keyHandler, true);

        clearIdBtn.addEventListener("click", function () {
            inputSubject.value = "";
            var ev = new Event("input", { bubbles: true });
            inputSubject.dispatchEvent(ev);
            log("Find Study Events: subject cleared by user");
        });

        cancelBtn.addEventListener("click", function () {
            log("Find Study Events: popup canceled");
            if (popup && popup.close) {
                popup.close();
            }
            document.removeEventListener("keydown", keyHandler, true);
            if (typeof onDone === "function") {
                onDone(null);
            }
        });

        okBtn.addEventListener("click", function () {
            document.removeEventListener("keydown", keyHandler, true);
            doContinue();
        });
    }

    function openFindStudyEvents() {
        log("Find Study Events: starting");
        var prefill = findSubjectIdentifierForFindForm();
        showFindStudyEventsPopup(prefill, function (userInput) {
            if (userInput === null) {
                log("Find Study Events: canceled by user; stopping");
                return;
            }
            var kw = userInput.keyword || "";
            var sbj = userInput.subject || "";
            previewStudyEventsByKeyword(kw, function (hasMatch) {
                if (!hasMatch) {
                    log("Find Study Events: preview found no matching study events; stopping");
                    showStudyEventNoMatchPopup();
                    return;
                }
                localStorage.setItem(STORAGE_FIND_STUDY_EVENT_PENDING, "1");
                localStorage.setItem(STORAGE_FIND_STUDY_EVENT_KEYWORD, String(kw));
                localStorage.setItem(STORAGE_FIND_STUDY_EVENT_SUBJECT, String(sbj));
                log("Find Study Events: state saved; redirecting same tab");
                window.location.href = FORM_LIST_URL;
            });
        });
    }

    function processFindStudyEventsOnList() {
        var pending = localStorage.getItem(STORAGE_FIND_STUDY_EVENT_PENDING);
        if (!pending || pending !== "1") {
            return;
        }
        var kw = localStorage.getItem(STORAGE_FIND_STUDY_EVENT_KEYWORD) || "";
        var sbj = localStorage.getItem(STORAGE_FIND_STUDY_EVENT_SUBJECT) || "";
        log("Find Study Events: post-reload processing started keyword='" + String(kw) + "' subject='" + String(sbj) + "'");
        var t = 0;
        var r = setInterval(function () {
            t = t + 1;
            if (t > 120) {
                clearInterval(r);
                log("Find Study Events: timeout waiting for controls");
                localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_PENDING);
                return;
            }
            var d = document;
            var eventsSel = d.querySelector('select#studyEventIds, select[name="studyEventIds"]');
            var subjSel = d.querySelector('select#subjectIds, select[name="subjectIds"]');
            var eventsBoxReady = document.getElementById("s2id_studyEventIds");
            var statusSel = d.getElementById("statusValues");
            var statusBoxReady = d.getElementById("s2id_statusValues");
            if (!eventsSel || !eventsBoxReady) {
                log("Find Study Events: selects not ready at t=" + String(t));
                return;
            }

            resetStatusValuesOnList();

            var rawStatuses = null;
            try {
                rawStatuses = localStorage.getItem(STORAGE_FIND_STUDY_EVENT_STATUS_VALUES);
            } catch (e) {}
            if (rawStatuses) {
                try {
                    var arr = JSON.parse(rawStatuses);
                    if (Array.isArray(arr) && arr.length > 0) {
                        applyStatusValuesOnList(arr);
                    } else {
                        log("Find Study Events: no statuses to apply");
                    }
                } catch (e2) {
                    log("Find Study Events: status parse error");
                }
            } else {
                log("Find Study Events: no stored statuses");
            }

            clearSelect2ChoicesByContainerId("s2id_studyEventIds");
            deselectAllOptionsBySelect(eventsSel);

            var selectedCount = 0;
            var eos = eventsSel.querySelectorAll("option");
            if (eos && eos.length > 0) {
                var i = 0;
                while (i < eos.length) {
                    var eop = eos[i];
                    var etxt = "";
                    if (eop) {
                        etxt = eop.textContent + "";
                    }
                    var hit = studyEventMatchContainsAllTokens(etxt, kw);
                    if (hit) {
                        eop.selected = true;
                        selectedCount = selectedCount + 1;
                    }
                    i = i + 1;
                }
                var evtEvents = new Event("change", { bubbles: true });
                eventsSel.dispatchEvent(evtEvents);
                log("Find Study Events: selected study events count=" + String(selectedCount));
            } else {
                log("Find Study Events: study events options list empty");
            }

            if (selectedCount === 0) {
                clearInterval(r);
                log("Find Study Events: no study events selected; notifying user");
                showStudyEventNoMatchPopup();
                localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_PENDING);
                localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_KEYWORD);
                localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_SUBJECT);
                return;
            }

            var subjNorm = aeNormalize(sbj || "");
            if (subjNorm && subjNorm.length > 0 && subjSel) {
                var os = subjSel.querySelectorAll("option");
                var y = "";
                if (os && os.length > 0) {
                    var j = 0;
                    while (j < os.length) {
                        var op = os[j];
                        var txt = "";
                        var val = "";
                        if (op) {
                            txt = op.textContent + "";
                            val = op.value + "";
                        }
                        var nt = aeNormalize(txt);
                        var matchByTextExact = nt === subjNorm;
                        var matchByTextContains = nt.indexOf(subjNorm) >= 0;
                        var matchById = val === sbj;
                        var matched = false;
                        if (matchByTextExact) {
                            matched = true;
                        } else {
                            if (matchByTextContains) {
                                matched = true;
                            } else {
                                if (matchById) {
                                    matched = true;
                                }
                            }
                        }
                        if (matched) {
                            y = val;
                            break;
                        }
                        j = j + 1;
                    }
                }
                if (y && y.length > 0) {
                    subjSel.value = y;
                    var evtSubj = new Event("change", { bubbles: true });
                    subjSel.dispatchEvent(evtSubj);
                    log("Find Study Events: subject applied value='" + String(y) + "'");
                } else {
                    log("Find Study Events: no subject match applied");
                }
            } else {
                log("Find Study Events: subject input empty or select not found");
            }

            var searchBtn = d.getElementById("dataSearchButton");
            if (searchBtn) {
                searchBtn.click();
                log("Find Study Events: Search button clicked");
            } else {
                log("Find Study Events: Search button not found");
            }

            clearInterval(r);
            localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_PENDING);
            localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_KEYWORD);
            localStorage.removeItem(STORAGE_FIND_STUDY_EVENT_SUBJECT);
            log("Find Study Events: post-reload processing completed");
        }, 200);
    }

    //==========================
    //==========================
    // Parse Method Functions
    //==========================
    //==========================

    function ParseMethodFunctions() {}

    function clearParseMethodState() {
        PARSE_METHOD_CANCELED = false;
        PARSE_METHOD_COLLECTED_METHODS = [];
        PARSE_METHOD_COLLECTED_FORMS = [];
        try { localStorage.removeItem(STORAGE_PARSE_METHOD_RUNNING); localStorage.removeItem(STORAGE_PARSE_METHOD_ITEM_NAME); } catch (e) {}
    }

    function stopParseMethodAutomation() {
        PARSE_METHOD_CANCELED = true;
        clearParseMethodState();
        try { localStorage.removeItem(STORAGE_RUN_MODE); } catch (e) {}
    }

    function clearParseMethodStoredResults() {
        try { localStorage.removeItem(STORAGE_PARSE_METHOD_RESULTS); localStorage.removeItem(STORAGE_PARSE_METHOD_COMPLETED); localStorage.removeItem(STORAGE_PARSE_METHOD_ITEM_NAME); } catch (e) {}
    }

    function openParseMethod() {
        log("ParseMethod: button clicked");
        if (isPaused()) return;
        clearParseMethodState();
        showParseMethodPopup();
    }

    function showParseMethodPopup() {
        var container = document.createElement("div");
        container.style.cssText = "display:flex;flex-direction:column;gap:16px";
        var inputRow = document.createElement("div");
        inputRow.style.cssText = "display:grid;grid-template-columns:100px 1fr;align-items:center;gap:12px";
        var label = document.createElement("div");
        label.textContent = "Item name";
        label.style.fontWeight = "600";
        var inputEl = document.createElement("input");
        inputEl.type = "text";
        inputEl.placeholder = "Required (case sensitive)";
        inputEl.style.cssText = "width:100%;box-sizing:border-box;padding:10px;border-radius:6px;border:1px solid #444;background:#1a1a1a;color:#fff";
        inputRow.appendChild(label);
        inputRow.appendChild(inputEl);
        container.appendChild(inputRow);
        var btnRow = document.createElement("div");
        btnRow.style.cssText = "display:flex;justify-content:flex-end";
        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirmed";
        confirmBtn.style.cssText = "background:#28a745;color:#fff;border:none;border-radius:6px;padding:10px 20px;cursor:pointer;font-weight:600";
        btnRow.appendChild(confirmBtn);
        container.appendChild(btnRow);
        var statusBox = document.createElement("div");
        statusBox.style.cssText = "display:none;margin-top:8px;padding:12px;background:#1a1a1a;border-radius:6px;border:1px solid #333;text-align:center";
        statusBox.textContent = "Running.";
        container.appendChild(statusBox);
        var resultsBox = document.createElement("div");
        resultsBox.style.cssText = "display:none;margin-top:8px;max-height:300px;overflow-y:auto;background:#1a1a1a;border-radius:6px;border:1px solid #333;padding:12px";
        container.appendChild(resultsBox);
        var okRow = document.createElement("div");
        okRow.style.cssText = "display:none;justify-content:center;margin-top:8px";
        var okBtn = document.createElement("button");
        okBtn.textContent = "Ok";
        okBtn.style.cssText = "background:#0b82ff;color:#fff;border:none;border-radius:6px;padding:10px 30px;cursor:pointer;font-weight:600";
        okRow.appendChild(okBtn);
        container.appendChild(okRow);
        var popup = createPopup({ title: "Parse Method", content: container, width: "480px", height: "auto", onClose: function() { stopParseMethodAutomation(); } });
        setTimeout(function() { try { inputEl.focus(); } catch (e) {} }, 50);

        function doConfirm() {
            var itemName = (inputEl.value + "").trim();
            if (!itemName) { inputEl.style.border = "2px solid #dc3545"; return; }
            inputEl.style.border = "1px solid #444";
            log("ParseMethod: Confirmed itemName=" + itemName);
            try { localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_PARSE_METHOD); localStorage.setItem(STORAGE_PARSE_METHOD_RUNNING, "1"); localStorage.setItem(STORAGE_PARSE_METHOD_ITEM_NAME, itemName); } catch (e) {}
            inputRow.style.display = "none";
            btnRow.style.display = "none";
            statusBox.style.display = "block";
            resultsBox.style.display = "block";
            var dots = 1;
            var interval = setInterval(function() { if (PARSE_METHOD_CANCELED) { clearInterval(interval); return; } dots = (dots % 3) + 1; var t = "Running"; for (var d = 0; d < dots; d++) t += "."; statusBox.textContent = t; }, 400);
            runParseMethodAutomation(itemName, function(methods, forms) {
                clearInterval(interval);
                if (PARSE_METHOD_CANCELED) return;
                statusBox.textContent = "Completed. Proceed to Study -> Data ?";
                showParseMethodResults(resultsBox, methods);
                okRow.style.display = "flex";
                okBtn.addEventListener("click", function() { okRow.style.display = "none"; statusBox.textContent = "Running Find Form..."; doFindFormWithForms(forms); });
            }, function(name, fms) { addParseMethodResult(resultsBox, name, fms); });
        }

        function keyHandler(e) {
            var code = e.key || e.code || "";
            if (code === "Enter") {
                log("ParseMethod: Enter pressed; confirming");
                doConfirm();
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        inputEl.addEventListener("keydown", keyHandler);

        confirmBtn.addEventListener("click", function() {
            doConfirm();
        });
    }

    function showParseMethodResults(box, methods) {
        box.innerHTML = "";
        if (!methods || methods.length === 0) { box.innerHTML = "<div style='color:#999;text-align:center;padding:20px'>No methods found.</div>"; return; }
        for (var i = 0; i < methods.length; i++) {
            var m = methods[i];
            var block = document.createElement("div");
            block.style.cssText = "margin-bottom:12px;padding:10px;background:#222;border-radius:6px;border:1px solid #444";
            var title = document.createElement("div");
            title.style.cssText = "font-weight:600;color:#4a90e2;margin-bottom:6px";
            title.textContent = m.name;
            block.appendChild(title);
            if (m.forms && m.forms.length > 0) {
                var list = document.createElement("div");
                list.style.paddingLeft = "16px";
                for (var j = 0; j < m.forms.length; j++) { var it = document.createElement("div"); it.style.cssText = "color:#ccc;font-size:12px"; it.textContent = "- " + m.forms[j]; list.appendChild(it); }
                block.appendChild(list);
            }
            box.appendChild(block);
        }
    }

    function addParseMethodResult(box, name, forms) {
        var block = document.createElement("div");
        block.style.cssText = "margin-bottom:12px;padding:10px;background:#222;border-radius:6px;border:1px solid #444";
        var title = document.createElement("div");
        title.style.cssText = "font-weight:600;color:#4a90e2;margin-bottom:6px";
        title.textContent = name;
        block.appendChild(title);
        if (forms && forms.length > 0) {
            var list = document.createElement("div");
            list.style.paddingLeft = "16px";
            for (var j = 0; j < forms.length; j++) { var it = document.createElement("div"); it.style.cssText = "color:#ccc;font-size:12px"; it.textContent = "- " + forms[j]; list.appendChild(it); }
            block.appendChild(list);
        }
        box.appendChild(block);
        box.scrollTop = box.scrollHeight;
    }

    async function runParseMethodAutomation(itemName, onComplete, onFound) {
        PARSE_METHOD_CANCELED = false;
        PARSE_METHOD_COLLECTED_METHODS = [];
        PARSE_METHOD_COLLECTED_FORMS = [];
        try {
            var methodsData = await getMethodLibraryData();
            if (PARSE_METHOD_CANCELED) return;
            if (!methodsData || methodsData.length === 0) { onComplete([], []); return; }
            for (var idx = 0; idx < methodsData.length; idx++) {
                if (PARSE_METHOD_CANCELED) return;
                if (isPaused()) { stopParseMethodAutomation(); return; }
                var method = methodsData[idx];
                var formalExpr = await getMethodFormalExpr(method.url);
                if (PARSE_METHOD_CANCELED) return;
                if (!formalExpr || formalExpr.indexOf('"' + itemName + '"') < 0) continue;
                log("ParseMethod: " + method.name + " matches");
                var itemRefs = await getMethodItemRefs(method.url);
                if (PARSE_METHOD_CANCELED) return;
                if (!itemRefs || itemRefs.length === 0) continue;
                var mForms = [];
                for (var ri = 0; ri < itemRefs.length; ri++) {
                    if (PARSE_METHOD_CANCELED) return;
                    var fs = await getItemRefForms(itemRefs[ri].url);
                    if (PARSE_METHOD_CANCELED) return;
                    if (fs) { for (var fi = 0; fi < fs.length; fi++) { if (mForms.indexOf(fs[fi]) < 0) mForms.push(fs[fi]); if (PARSE_METHOD_COLLECTED_FORMS.indexOf(fs[fi]) < 0) PARSE_METHOD_COLLECTED_FORMS.push(fs[fi]); } }
                }
                if (mForms.length > 0) { PARSE_METHOD_COLLECTED_METHODS.push({ name: method.name, forms: mForms }); if (typeof onFound === "function") onFound(method.name, mForms); }
            }
            onComplete(PARSE_METHOD_COLLECTED_METHODS, PARSE_METHOD_COLLECTED_FORMS);
        } catch (err) { log("ParseMethod: error - " + err); onComplete([], []); }
    }

    async function getMethodLibraryData() {
        return new Promise(function(resolve) {
            GM.xmlHttpRequest({ method: "GET", url: METHOD_LIBRARY_URL, onload: function(resp) {
                var html = resp && resp.responseText ? resp.responseText : "";
                var methods = [];
                var tmp = document.createElement("div"); tmp.innerHTML = html;
                var table = tmp.querySelector("table#listTable") || tmp.querySelector("table");
                if (table) { var tbody = table.querySelector("tbody"); if (tbody) { var rows = tbody.querySelectorAll("tr"); for (var i = 0; i < rows.length; i++) { var tds = rows[i].querySelectorAll("td"); if (tds.length >= 1) { var link = tds[0].querySelector("a"); if (link) { var name = (link.textContent || "").trim(); var href = link.getAttribute("href") || ""; if (name && href) { var url = href.indexOf("http") === 0 ? href : "https://cenexel.clinspark.com" + href; methods.push({ name: name, url: url }); } } } } } }
                resolve(methods);
            }, onerror: function() { resolve([]); } });
        });
    }

    async function getMethodFormalExpr(methodUrl) {
        return new Promise(function(resolve) {
            GM.xmlHttpRequest({ method: "GET", url: methodUrl, onload: function(resp) {
                var html = resp && resp.responseText ? resp.responseText : "";
                var tmp = document.createElement("div"); tmp.innerHTML = html;
                var tables = tmp.querySelectorAll("table.table.table-striped.table-bordered");
                for (var i = 0; i < tables.length; i++) { var tbody = tables[i].querySelector("tbody"); if (tbody) { var rows = tbody.querySelectorAll("tr"); for (var j = 0; j < rows.length; j++) { var tds = rows[j].querySelectorAll("td"); if (tds.length >= 2 && (tds[0].textContent || "").trim() === "Formal Expression:") { resolve((tds[1].textContent || "").trim()); return; } } } }
                resolve("");
            }, onerror: function() { resolve(""); } });
        });
    }

    async function getMethodItemRefs(methodUrl) {
        return new Promise(function(resolve) {
            GM.xmlHttpRequest({ method: "GET", url: methodUrl + "?references=references", onload: function(resp) {
                var html = resp && resp.responseText ? resp.responseText : "";
                var refs = [];
                var tmp = document.createElement("div"); tmp.innerHTML = html;
                var groups = tmp.querySelectorAll("ul.list-group");
                for (var i = 0; i < groups.length; i++) { var items = groups[i].querySelectorAll("li.list-group-item a"); for (var j = 0; j < items.length; j++) { var href = items[j].getAttribute("href") || ""; if (href.indexOf("/show/itemgroup/") >= 0) { var fullUrl = href.indexOf("http") === 0 ? href : "https://cenexel.clinspark.com" + href; refs.push({ url: fullUrl }); } } }
                resolve(refs);
            }, onerror: function() { resolve([]); } });
        });
    }

    async function getItemRefForms(itemRefUrl) {
        return new Promise(function(resolve) {
            GM.xmlHttpRequest({ method: "GET", url: itemRefUrl + "?references=references", onload: function(resp) {
                var html = resp && resp.responseText ? resp.responseText : "";
                var forms = [];
                var tmp = document.createElement("div"); tmp.innerHTML = html;
                var groups = tmp.querySelectorAll("ul.list-group");
                for (var i = 0; i < groups.length; i++) { var items = groups[i].querySelectorAll("li.list-group-item a"); for (var j = 0; j < items.length; j++) { var href = items[j].getAttribute("href") || ""; var text = (items[j].textContent || "").trim(); if (href.indexOf("/show/form/") >= 0 && text) { forms.push(text); } } }
                resolve(forms);
            }, onerror: function() { resolve([]); } });
        });
    }

    function doFindFormWithForms(formNames) {
        if (!formNames || formNames.length === 0) { log("ParseMethod: no forms"); return; }
        log("ParseMethod: Find Form with " + formNames.length + " forms");
        try { localStorage.setItem(STORAGE_PARSE_METHOD_RESULTS, JSON.stringify(PARSE_METHOD_COLLECTED_METHODS)); localStorage.setItem(STORAGE_PARSE_METHOD_COMPLETED, "1"); } catch (e) {}
        try { localStorage.removeItem(STORAGE_FIND_FORM_STATUS_VALUES); } catch (e) {}
        var formNamesJson = JSON.stringify(formNames);
        localStorage.setItem(STORAGE_FIND_FORM_KEYWORD, formNamesJson);
        localStorage.setItem(STORAGE_FIND_FORM_SUBJECT, "");
        window.location.href = FORM_LIST_URL;
    }

    function showParseMethodCompletedPopup(methods) {
        var container = document.createElement("div");
        container.style.cssText = "display:flex;flex-direction:column;gap:16px";
        var statusBox = document.createElement("div");
        statusBox.style.cssText = "padding:12px;background:#1a1a1a;border-radius:6px;border:1px solid #333;text-align:center";
        statusBox.textContent = "Find Form completed.";
        container.appendChild(statusBox);
        var resultsBox = document.createElement("div");
        resultsBox.style.cssText = "max-height:300px;overflow-y:auto;background:#1a1a1a;border-radius:6px;border:1px solid #333;padding:12px";
        showParseMethodResults(resultsBox, methods);
        container.appendChild(resultsBox);
        createPopup({ title: "Parse Method", content: container, width: "480px", height: "auto", onClose: function() { clearParseMethodStoredResults(); } });
    }

    function checkAndRestoreParseMethodPopup() {
        var completed = null;
        try { completed = localStorage.getItem(STORAGE_PARSE_METHOD_COMPLETED); } catch (e) {}
        if (completed !== "1") return false;
        var resultsRaw = null;
        try { resultsRaw = localStorage.getItem(STORAGE_PARSE_METHOD_RESULTS); } catch (e) {}
        if (!resultsRaw) { clearParseMethodStoredResults(); return false; }
        try {
            var methods = JSON.parse(resultsRaw);
            if (Array.isArray(methods)) {
                log("ParseMethod: restoring completed popup");
                populateFormsFromParseMethod();
                showParseMethodCompletedPopup(methods);
                return true;
            }
        } catch (e) { log("ParseMethod: failed to parse stored results"); }
        clearParseMethodStoredResults();
        return false;
    }

    function populateFormsFromParseMethod() {
        var kw = "";
        try { kw = localStorage.getItem(STORAGE_FIND_FORM_KEYWORD) || ""; } catch (e) {}
        if (!kw) { log("ParseMethod: no stored keyword for forms"); return; }
        var formNames = [];
        try { formNames = JSON.parse(kw); } catch (e) { log("ParseMethod: failed to parse form names JSON"); return; }
        if (!Array.isArray(formNames) || formNames.length === 0) { log("ParseMethod: no form names in array"); return; }
        log("ParseMethod: looking for " + formNames.length + " form names");
        var waitCount = 0;
        var waitInterval = setInterval(function() {
            waitCount++;
            if (waitCount > 60) { clearInterval(waitInterval); log("ParseMethod: timeout waiting for form select"); return; }
            var formsSel = document.querySelector('select#formIds, select[name="formIds"]');
            if (!formsSel) return;
            clearInterval(waitInterval);
            var fos = formsSel.querySelectorAll("option");
            var formIds = [];
            for (var i = 0; i < fos.length; i++) {
                var fop = fos[i];
                var ftxt = (fop.textContent || "").trim();
                var fval = fop.value || "";
                for (var j = 0; j < formNames.length; j++) {
                    var formName = formNames[j].trim();
                    if (formName && ftxt === formName && fval) {
                        formIds.push(fval);
                        log("ParseMethod: matched form '" + formName + "' with id=" + fval);
                        break;
                    }
                }
            }
            log("ParseMethod: found " + formIds.length + " form IDs");
            if (formIds.length > 0) {
                var url = FORM_LIST_URL + "?search=true";
                for (var k = 0; k < formIds.length; k++) { url += "&formIds=" + encodeURIComponent(formIds[k]); }
                log("ParseMethod: navigating to URL with formIds");
                try { localStorage.removeItem(STORAGE_FIND_FORM_KEYWORD); localStorage.removeItem(STORAGE_FIND_FORM_SUBJECT); } catch (e) {}
                window.location.href = url;
            } else {
                log("ParseMethod: no matching forms found");
                try { localStorage.removeItem(STORAGE_FIND_FORM_KEYWORD); localStorage.removeItem(STORAGE_FIND_FORM_SUBJECT); } catch (e) {}
            }
        }, 100);
    }

    //==========================
    // RUN SUBJECT ELIGIBILITY FEATURE
    //==========================
    // This section contains all functions related to subject eligibility.
    // This feature automates storing all existing eligibility mapping in the table,
    // adding new eligibility item that cannot be found in the table,
    // saving those new items and storing it so there is no duplicate.
    //==========================

    function SubjectEligibilityFunctions() {}
    const ELIGIBILITY_LIST_URL_PROD = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const ELIGIBILITY_VALID_HOSTNAMES = ["cenexeltest.clinspark.com", "cenexel.clinspark.com"];
    const ELIGIBILITY_LIST_PATH = "/secure/crfdesign/studylibrary/eligibility/list";
    const IE_CODE_REGEX = /\b(INC|EXC)\s*(\d+)\b/i;
    const IE_CODE_REGEX_GLOBAL = /\b(INC|EXC)\s*(\d+)\b/gi;
    const IMPORT_IE_HELPER_TIMEOUT = 15000;
    const IMPORT_IE_POLL_INTERVAL = 120;
    const IMPORT_IE_MODAL_TIMEOUT = 12000;
    const IMPORT_IE_SHORT_DELAY_MIN = 150;
    const IMPORT_IE_SHORT_DELAY_MAX = 400;
    var IMPORT_IE_CANCELED = false;

    // Update Import Eligibility popup lists
    function addToImportEligCompletedList(itemCode) {
        try {
            var list = document.getElementById("importEligCompletedList");
            if (!list && IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.element) {
                list = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligCompletedList");
            }
            if (list) {
                var item = document.createElement("div");
                item.textContent = itemCode;
                item.style.padding = "2px 4px";
                list.appendChild(item);
                list.scrollTop = list.scrollHeight;
            }
        } catch (e) {
            log("Error adding to completed list: " + e);
        }
    }

    function addToImportEligFailedList(itemCode) {
        try {
            var list = document.getElementById("importEligFailedList");
            if (!list && IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.element) {
                list = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligFailedList");
            }
            if (list) {
                var item = document.createElement("div");
                item.textContent = itemCode;
                item.style.padding = "2px 4px";
                list.appendChild(item);
                list.scrollTop = list.scrollHeight;
            }
        } catch (e) {
            log("Error adding to failed list: " + e);
        }
    }

    function addToImportEligExcludedList(itemCode) {
        try {
            var list = document.getElementById("importEligExcludedList");
            if (!list && IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.element) {
                list = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligExcludedList");
            }
            if (list) {
                var item = document.createElement("div");
                item.textContent = itemCode;
                item.style.padding = "2px 4px";
                list.appendChild(item);
                list.scrollTop = list.scrollHeight;
            }
        } catch (e) {
            log("Error adding to excluded list: " + e);
        }
    }

    function isRunModeSet(expectedMode) {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            return false;
        }
        return raw === expectedMode;
    }

    function setLastMatchSelection(planVal, saVal, itemVal) {
        try {
            localStorage.setItem(STORAGE_ELIG_LAST_PLAN, String(planVal));
            log("ImportElig: pinned plan=" + String(planVal));
        } catch(e) {}
        try {
            localStorage.setItem(STORAGE_ELIG_LAST_SA, String(saVal));
            log("ImportElig: pinned SA=" + String(saVal));
        } catch(e) {}
        try {
            localStorage.setItem(STORAGE_ELIG_LAST_ITEMREF, String(itemVal));
            log("ImportElig: pinned itemRef=" + String(itemVal));
        } catch(e) {}
    }

    function getLastMatchSelection() {
        var p = "";
        var s = "";
        var i = "";
        try {
            p = localStorage.getItem(STORAGE_ELIG_LAST_PLAN);
            if (!p) {
                p = "";
            }
        } catch(e1) {}
        try {
            s = localStorage.getItem(STORAGE_ELIG_LAST_SA);
            if (!s) {
                s = "";
            }
        } catch(e2) {}
        try {
            i = localStorage.getItem(STORAGE_ELIG_LAST_ITEMREF);
            if (!i) {
                i = "";
            }
        } catch(e3) {}
        return { plan: String(p), sa: String(s), itemRef: String(i) };
    }

    function clearLastMatchSelection() {
        try {
            localStorage.removeItem(STORAGE_ELIG_LAST_PLAN);
        } catch(e1) {}
        try {
            localStorage.removeItem(STORAGE_ELIG_LAST_SA);
        } catch(e2) {}
        try {
            localStorage.removeItem(STORAGE_ELIG_LAST_ITEMREF);
        } catch(e3) {}
        log("ImportElig: cleared pinned selections");
    }

    async function stabilizeSelectionBeforeComparator(timeoutMs) {
        log("ImportElig: stabilizeSelectionBeforeComparator start");
        var max = typeof timeoutMs === "number" ? timeoutMs : 6000;
        var start = Date.now();
        var last = getLastMatchSelection();
        if (!last.plan || !last.sa) {
            log("ImportElig: no pinned plan/sa to stabilize");
            return;
        }
        while (Date.now() - start < max) {
            var planSel = document.querySelector("select#activityPlan");
            var saSel = document.querySelector("select#scheduledActivity");
            var itemSel = document.querySelector("select#itemRef");
            if (!planSel || !saSel) {
                await sleep(150);
                continue;
            }
            var changed = false;
            if (String(planSel.value) !== String(last.plan)) {
                planSel.value = last.plan;
                var evtP = new Event("change", { bubbles: true });
                planSel.dispatchEvent(evtP);
                log("ImportElig: restored plan to pinned value");
                await sleep(700);
                changed = true;
            }
            var saSel2 = document.querySelector("select#scheduledActivity");
            if (saSel2) {
                saSel = saSel2;
            }
            if (String(saSel.value) !== String(last.sa)) {
                var prevSig = getItemRefOptionsSignature();
                saSel.value = last.sa;
                var evtS = new Event("change", { bubbles: true });
                saSel.dispatchEvent(evtS);
                log("ImportElig: restored SA to pinned value");
                var reloaded = await waitForItemRefReload(prevSig, 1200);
                if (!reloaded) {
                    log("ImportElig: itemRef did not reload after SA restore");
                }
                changed = true;
            }
            itemSel = document.querySelector("select#itemRef");
            if (itemSel) {
                if (String(itemSel.value) !== String(last.itemRef) && String(last.itemRef).length > 0) {
                    itemSel.value = last.itemRef;
                    var evtI = new Event("change", { bubbles: true });
                    itemSel.dispatchEvent(evtI);
                    log("ImportElig: restored itemRef to pinned value");
                    await sleep(350);
                }
            }
            if (!changed) {
                log("ImportElig: selection stable");
                return;
            }
        }
        log("ImportElig: stabilization timeout");
    }

    async function waitForComparatorReady(timeoutMs) {
        log("ImportElig: waitForComparatorReady start");
        var max = typeof timeoutMs === "number" ? timeoutMs : 8000;
        var start = Date.now();
        var lastLog = 0;
        while (Date.now() - start < max) {
            var sel = document.querySelector("select#eligibilityComparator");
            if (sel) {
                var opts = sel.querySelectorAll("option");
                var count = 0;
                var i = 0;
                while (i < opts.length) {
                    var v = (opts[i].value + "").trim();
                    if (v.length > 0) {
                        count = count + 1;
                    }
                    i = i + 1;
                }
                if (count >= 2) {
                    log("ImportElig: comparator ready with " + String(count) + " options");
                    return sel;
                }
            }
            if (Date.now() - lastLog > 1000) {
                log("ImportElig: comparator not ready yet");
                lastLog = Date.now();
            }
            await sleep(300);
        }
        log("ImportElig: waitForComparatorReady timeout");
        return null;
    }
    function clearEligibilityWorkingState() {
        try {
            localStorage.removeItem(STORAGE_ELIG_IMPORTED);
            log("ImportElig: cleared imported items");
        } catch(e) {}

        try {
            localStorage.removeItem(STORAGE_ELIG_CHECKITEM_CACHE);
            log("ImportElig: cleared checkItemCache");
        } catch(e) {}

        log("ImportElig: working state fully cleared");
    }

    function getItemRefOptionsSignature() {
        var sel = document.querySelector("select#itemRef");
        var sig = "";
        if (!sel) {
            return sig;
        }
        var opts = sel.querySelectorAll("option");
        var parts = [];
        var i = 0;
        while (i < opts.length) {
            var o = opts[i];
            var v = (o.value + "").trim();
            var t = (o.textContent + "").trim();
            parts.push(v + "|" + t);
            i = i + 1;
        }
        sig = parts.join("||");
        return sig;
    }

    async function waitForItemRefReload(prevSig, timeoutMs) {
        var start = Date.now();
        var step = 500;
        var max = typeof timeoutMs === "number" ? timeoutMs : 900;
        while (Date.now() - start < max) {
            var curSig = getItemRefOptionsSignature();
            if (curSig && curSig !== prevSig) {
                return true;
            }
            await sleep(step);
        }
        log("ImportElig: itemRef did not reload before timeout");
        return false;
    }

    function buildImportEligPopup(onConfirm) {
        log("ImportElig: building popup");

        var storedPlanPriStr = localStorage.getItem(STORAGE_ELIG_PLAN_PRIORITY);
        if (!storedPlanPriStr || storedPlanPriStr.trim() === "" || storedPlanPriStr.trim() === "[]") {
            storedPlanPriStr = "";
        }
        var storedExcStr = localStorage.getItem(STORAGE_ELIG_FORM_EXCLUSION);
        if (!storedExcStr || storedExcStr.trim() === "" || storedExcStr.trim() === "[]") {
            storedExcStr = "";
        }
        var storedPriStr = localStorage.getItem(STORAGE_ELIG_FORM_PRIORITY);
        if (!storedPriStr || storedPriStr.trim() === "" || storedPriStr.trim() === "[]") {
            storedPriStr = "";
        }
        var storedIgnoreStr = localStorage.getItem(STORAGE_ELIG_IGNORE);
        if (!storedIgnoreStr || storedIgnoreStr.trim() === "" || storedIgnoreStr.trim() === "[]") {
            storedIgnoreStr = "";
        }
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "20px";

        var planPriBox = document.createElement("div");
        planPriBox.style.display = "flex";
        planPriBox.style.flexDirection = "column";
        planPriBox.style.gap = "8px";

        var planPriHeaderRow = document.createElement("div");
        planPriHeaderRow.style.display = "flex";
        planPriHeaderRow.style.alignItems = "center";
        planPriHeaderRow.style.justifyContent = "space-between";
        planPriHeaderRow.style.gap = "8px";

        var planPriLabel = document.createElement("div");
        planPriLabel.textContent = "Activity Plan Priority (comma-separated)";
        planPriLabel.style.fontWeight = "600";
        planPriLabel.style.flex = "1";

        var planPriDefaultBtn = document.createElement("button");
        planPriDefaultBtn.textContent = "Clear All";
        planPriDefaultBtn.style.background = "#444";
        planPriDefaultBtn.style.color = "#fff";
        planPriDefaultBtn.style.border = "none";
        planPriDefaultBtn.style.padding = "4px 12px";
        planPriDefaultBtn.style.borderRadius = "4px";
        planPriDefaultBtn.style.cursor = "pointer";
        planPriDefaultBtn.style.fontSize = "12px";
        planPriDefaultBtn.style.whiteSpace = "nowrap";

        planPriHeaderRow.appendChild(planPriLabel);
        planPriHeaderRow.appendChild(planPriDefaultBtn);

        var planPriInput = document.createElement("textarea");
        planPriInput.style.width = "100%";
        planPriInput.style.height = "70px";
        planPriInput.style.background = "#222";
        planPriInput.style.color = "#fff";
        planPriInput.style.border = "1px solid #555";
        planPriInput.style.borderRadius = "4px";
        planPriInput.value = storedPlanPriStr;

        planPriDefaultBtn.addEventListener("click", function () {
            planPriInput.value = "";
            log("ImportElig: Activity Plan Priority cleared (no default)");
        });

        planPriBox.appendChild(planPriHeaderRow);
        planPriBox.appendChild(planPriInput);

        var ignoreBox = document.createElement("div");
        ignoreBox.style.display = "flex";
        ignoreBox.style.flexDirection = "column";
        ignoreBox.style.gap = "8px";

        var ignoreHeaderRow = document.createElement("div");
        ignoreHeaderRow.style.display = "flex";
        ignoreHeaderRow.style.alignItems = "center";
        ignoreHeaderRow.style.justifyContent = "space-between";
        ignoreHeaderRow.style.gap = "8px";

        var ignoreLabel = document.createElement("div");
        ignoreLabel.textContent = "Ignore I/E (comma-separated)";
        ignoreLabel.style.fontWeight = "600";
        ignoreLabel.style.flex = "1";

        var ignoreDefaultBtn = document.createElement("button");
        ignoreDefaultBtn.textContent = "Clear All";
        ignoreDefaultBtn.style.background = "#444";
        ignoreDefaultBtn.style.color = "#fff";
        ignoreDefaultBtn.style.border = "none";
        ignoreDefaultBtn.style.padding = "4px 12px";
        ignoreDefaultBtn.style.borderRadius = "4px";
        ignoreDefaultBtn.style.cursor = "pointer";
        ignoreDefaultBtn.style.fontSize = "12px";
        ignoreDefaultBtn.style.whiteSpace = "nowrap";

        ignoreHeaderRow.appendChild(ignoreLabel);
        ignoreHeaderRow.appendChild(ignoreDefaultBtn);

        var ignoreInput = document.createElement("textarea");
        ignoreInput.style.width = "100%";
        ignoreInput.style.height = "70px";
        ignoreInput.style.background = "#222";
        ignoreInput.style.color = "#fff";
        ignoreInput.style.border = "1px solid #555";
        ignoreInput.style.borderRadius = "4px";
        ignoreInput.value = storedIgnoreStr;

        ignoreDefaultBtn.addEventListener("click", function () {
            ignoreInput.value = "";
            log("ImportElig: Ignore I/E cleared (no default)");
        });

        ignoreBox.appendChild(ignoreHeaderRow);
        ignoreBox.appendChild(ignoreInput);

        var excBox = document.createElement("div");
        excBox.style.display = "flex";
        excBox.style.flexDirection = "column";
        excBox.style.gap = "8px";

        var excHeaderRow = document.createElement("div");
        excHeaderRow.style.display = "flex";
        excHeaderRow.style.alignItems = "center";
        excHeaderRow.style.justifyContent = "space-between";
        excHeaderRow.style.gap = "8px";

        var excLabel = document.createElement("div");
        excLabel.textContent = "Form Exclusion (comma-separated)";
        excLabel.style.fontWeight = "600";
        excLabel.style.flex = "1";

        var excDefaultBtn = document.createElement("button");
        excDefaultBtn.textContent = "Use Default";
        excDefaultBtn.style.background = "#444";
        excDefaultBtn.style.color = "#fff";
        excDefaultBtn.style.border = "none";
        excDefaultBtn.style.padding = "4px 12px";
        excDefaultBtn.style.borderRadius = "4px";
        excDefaultBtn.style.cursor = "pointer";
        excDefaultBtn.style.fontSize = "12px";
        excDefaultBtn.style.whiteSpace = "nowrap";

        excHeaderRow.appendChild(excLabel);
        excHeaderRow.appendChild(excDefaultBtn);

        var excInput = document.createElement("textarea");
        excInput.style.width = "100%";
        excInput.style.height = "70px";
        excInput.style.background = "#222";
        excInput.style.color = "#fff";
        excInput.style.border = "1px solid #555";
        excInput.style.borderRadius = "4px";
        excInput.value = storedExcStr;

        excDefaultBtn.addEventListener("click", function () {
            excInput.value = DEFAULT_FORM_EXCLUSION;
            log("ImportElig: Form Exclusion populated with default value");
        });

        excBox.appendChild(excHeaderRow);
        excBox.appendChild(excInput);

        var priBox = document.createElement("div");
        priBox.style.display = "flex";
        priBox.style.flexDirection = "column";
        priBox.style.gap = "8px";

        var priHeaderRow = document.createElement("div");
        priHeaderRow.style.display = "flex";
        priHeaderRow.style.alignItems = "center";
        priHeaderRow.style.justifyContent = "space-between";
        priHeaderRow.style.gap = "8px";

        var priLabel = document.createElement("div");
        priLabel.textContent = "Form Priority (comma-separated)";
        priLabel.style.fontWeight = "600";
        priLabel.style.flex = "1";

        var priOnlyCheckbox = document.createElement("input");
        priOnlyCheckbox.type = "checkbox";
        priOnlyCheckbox.id = "formPriorityOnlyCheckbox";
        var storedPriOnly = localStorage.getItem(STORAGE_ELIG_FORM_PRIORITY_ONLY);
        priOnlyCheckbox.checked = storedPriOnly === "1";
        priOnlyCheckbox.style.cursor = "pointer";
        priOnlyCheckbox.style.marginRight = "8px";

        var priOnlyLabel = document.createElement("label");
        priOnlyLabel.textContent = "Only";
        priOnlyLabel.style.fontSize = "12px";
        priOnlyLabel.style.cursor = "pointer";
        priOnlyLabel.style.marginRight = "8px";
        priOnlyLabel.setAttribute("for", "formPriorityOnlyCheckbox");

        var priDefaultBtn = document.createElement("button");
        priDefaultBtn.textContent = "Use Default";
        priDefaultBtn.style.background = "#444";
        priDefaultBtn.style.color = "#fff";
        priDefaultBtn.style.border = "none";
        priDefaultBtn.style.padding = "4px 12px";
        priDefaultBtn.style.borderRadius = "4px";
        priDefaultBtn.style.cursor = "pointer";
        priDefaultBtn.style.fontSize = "12px";
        priDefaultBtn.style.whiteSpace = "nowrap";

        var priRightControls = document.createElement("div");
        priRightControls.style.display = "flex";
        priRightControls.style.alignItems = "center";
        priRightControls.style.gap = "4px";
        priRightControls.appendChild(priOnlyCheckbox);
        priRightControls.appendChild(priOnlyLabel);
        priRightControls.appendChild(priDefaultBtn);

        priHeaderRow.appendChild(priLabel);
        priHeaderRow.appendChild(priRightControls);

        var priInput = document.createElement("textarea");
        priInput.style.width = "100%";
        priInput.style.height = "70px";
        priInput.style.background = "#222";
        priInput.style.color = "#fff";
        priInput.style.border = "1px solid #555";
        priInput.style.borderRadius = "4px";
        priInput.value = storedPriStr;

        priDefaultBtn.addEventListener("click", function () {
            priInput.value = DEFAULT_FORM_PRIORITY;
            log("ImportElig: Form Priority populated with default value");
        });

        priBox.appendChild(priHeaderRow);
        priBox.appendChild(priInput);

        container.appendChild(planPriBox);
        container.appendChild(ignoreBox);
        container.appendChild(excBox);
        container.appendChild(priBox);

        var buttonRow = document.createElement("div");
        buttonRow.style.display = "flex";
        buttonRow.style.gap = "10px";
        buttonRow.style.justifyContent = "flex-start";

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#0a0";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "none";
        confirmBtn.style.padding = "10px";
        confirmBtn.style.borderRadius = "6px";
        confirmBtn.style.cursor = "pointer";
        confirmBtn.style.flex = "1";

        var clearAllBtn = document.createElement("button");
        clearAllBtn.textContent = "Clear All";
        clearAllBtn.style.background = "#666";
        clearAllBtn.style.color = "#fff";
        clearAllBtn.style.border = "none";
        clearAllBtn.style.padding = "10px";
        clearAllBtn.style.borderRadius = "6px";
        clearAllBtn.style.cursor = "pointer";
        clearAllBtn.style.flex = "1";

        clearAllBtn.addEventListener("click", function () {
            planPriInput.value = "";
            ignoreInput.value = "";
            excInput.value = "";
            priInput.value = "";
            priOnlyCheckbox.checked = false;
            log("ImportElig: All input boxes cleared");
        });

        buttonRow.appendChild(confirmBtn);
        buttonRow.appendChild(clearAllBtn);
        container.appendChild(buttonRow);

        var runningContainer = document.createElement("div");
        runningContainer.id = "importEligRunningContainer";
        runningContainer.style.display = "none";
        runningContainer.style.marginTop = "10px";
        runningContainer.style.flexDirection = "column";
        runningContainer.style.gap = "12px";

        var runningBox = document.createElement("div");
        runningBox.id = "importEligRunningText";
        runningBox.style.textAlign = "center";
        runningBox.style.fontSize = "16px";
        runningBox.style.color = "#fff";
        runningBox.style.fontWeight = "500";
        runningBox.textContent = "Running";
        runningContainer.appendChild(runningBox);

        var listsContainer = document.createElement("div");
        listsContainer.style.display = "flex";
        listsContainer.style.flexDirection = "column";
        listsContainer.style.gap = "12px";
        listsContainer.style.maxHeight = "300px";
        listsContainer.style.overflowY = "auto";

        var completedListContainer = document.createElement("div");
        completedListContainer.style.display = "flex";
        completedListContainer.style.flexDirection = "column";
        completedListContainer.style.gap = "4px";
        var completedTitle = document.createElement("div");
        completedTitle.textContent = "Completed:";
        completedTitle.style.fontWeight = "600";
        completedTitle.style.color = "#5cb85c";
        completedTitle.style.fontSize = "13px";
        completedListContainer.appendChild(completedTitle);
        var completedList = document.createElement("div");
        completedList.id = "importEligCompletedList";
        completedList.style.fontSize = "12px";
        completedList.style.color = "#9f9";
        completedList.style.maxHeight = "100px";
        completedList.style.overflowY = "auto";
        completedList.style.padding = "4px";
        completedList.style.background = "#1a1a1a";
        completedList.style.borderRadius = "4px";
        completedListContainer.appendChild(completedList);
        listsContainer.appendChild(completedListContainer);

        var failedListContainer = document.createElement("div");
        failedListContainer.style.display = "flex";
        failedListContainer.style.flexDirection = "column";
        failedListContainer.style.gap = "4px";
        var failedTitle = document.createElement("div");
        failedTitle.textContent = "Failed to Collect:";
        failedTitle.style.fontWeight = "600";
        failedTitle.style.color = "#d9534f";
        failedTitle.style.fontSize = "13px";
        failedListContainer.appendChild(failedTitle);
        var failedList = document.createElement("div");
        failedList.id = "importEligFailedList";
        failedList.style.fontSize = "12px";
        failedList.style.color = "#f99";
        failedList.style.maxHeight = "100px";
        failedList.style.overflowY = "auto";
        failedList.style.padding = "4px";
        failedList.style.background = "#1a1a1a";
        failedList.style.borderRadius = "4px";
        failedListContainer.appendChild(failedList);
        listsContainer.appendChild(failedListContainer);

        var excludedListContainer = document.createElement("div");
        excludedListContainer.style.display = "flex";
        excludedListContainer.style.flexDirection = "column";
        excludedListContainer.style.gap = "4px";
        var excludedTitle = document.createElement("div");
        excludedTitle.textContent = "Form Exclusion:";
        excludedTitle.style.fontWeight = "600";
        excludedTitle.style.color = "#f0ad4e";
        excludedTitle.style.fontSize = "13px";
        excludedListContainer.appendChild(excludedTitle);
        var excludedList = document.createElement("div");
        excludedList.id = "importEligExcludedList";
        excludedList.style.fontSize = "12px";
        excludedList.style.color = "#ff9";
        excludedList.style.maxHeight = "100px";
        excludedList.style.overflowY = "auto";
        excludedList.style.padding = "4px";
        excludedList.style.background = "#1a1a1a";
        excludedList.style.borderRadius = "4px";
        excludedListContainer.appendChild(excludedList);
        listsContainer.appendChild(excludedListContainer);

        runningContainer.appendChild(listsContainer);
        container.appendChild(runningContainer);

        var popup = createPopup({
            title: "Import Eligibility Mapping",
            content: container,
            width: "500px",
            height: "auto",
            maxHeight: "80vh",
            onClose: function () {
                // X pressed → STOP everything (but don't pause)
                log("ImportElig: popup X pressed → stopping automation");
                clearAllRunState();
                clearEligibilityWorkingState();
                // Clear pending popup flag
                try {
                    localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {}
                IMPORT_ELIG_POPUP_REF = null;
            }
        });

        IMPORT_ELIG_POPUP_REF = popup;
        try {
            localStorage.setItem(STORAGE_IMPORT_ELIG_POPUP, "1");
        } catch (e) {}

        confirmBtn.addEventListener("click", function () {
            log("ImportElig: Confirm clicked");

            // Set run mode only when Confirm is actually pressed
            try {
                localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ELIG_IMPORT);
                // Clear pending popup flag since we're now starting automation
                localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
            } catch (e) {
            }

            var planPriText = (planPriInput.value + "").trim();
            var ignoreText = (ignoreInput.value + "").trim();
            var excText = (excInput.value + "").trim();
            var priText = (priInput.value + "").trim();
            var priOnlyChecked = priOnlyCheckbox.checked;

            // Save empty strings if inputs are empty (don't save default values)
            setPlanPriority(planPriText);
            setIgnoreKeywords(ignoreText);
            setFormExclusion(excText);
            setFormPriority(priText);
            setFormPriorityOnly(priOnlyChecked);

            planPriInput.disabled = true;
            ignoreInput.disabled = true;
            excInput.disabled = true;
            priInput.disabled = true;
            priOnlyCheckbox.disabled = true;

            planPriDefaultBtn.disabled = true;
            ignoreDefaultBtn.disabled = true;
            excDefaultBtn.disabled = true;
            priDefaultBtn.disabled = true;
            clearAllBtn.disabled = true;

            planPriBox.style.display = "none";
            ignoreBox.style.display = "none";
            excBox.style.display = "none";
            priBox.style.display = "none";

            confirmBtn.style.display = "none";
            clearAllBtn.style.display = "none";

            runningContainer.style.display = "flex";

            // Set popup flag so it persists across page changes
            try {
                localStorage.setItem(STORAGE_IMPORT_ELIG_POPUP, "1");
            } catch (e) {}

            var dots = 1;
            var interval = setInterval(function () {
                if (!popup || !document.body.contains(popup.element)) {
                    clearInterval(interval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var t = "Running";
                var i = 0;
                while (i < dots) {
                    t = t + ".";
                    i = i + 1;
                }
                if (runningBox) {
                    runningBox.textContent = t;
                }
            }, 350);

            onConfirm(function (message) {
                clearInterval(interval);
                if (message) {
                    if (runningBox) {
                        runningBox.textContent = "No more I/E items to add";
                    }
                    // Don't auto-close - user must click X
                } else {
                    if (runningBox) {
                        runningBox.textContent = "No more I/E items to add";
                    }
                    // Don't auto-close - user must click X
                }
            });
        });
    }

    function parseStoredKeywords(rawText) {
        var arr = [];
        if (rawText) {
            var parts = rawText.split(",");
            var i = 0;
            while (i < parts.length) {
                var t = (parts[i] + "").trim();
                if (t.length > 0) {
                    arr.push(t);
                }
                i = i + 1;
            }
        }
        return arr;
    }

    function getFormExclusion() {
        var raw = null;

        try {
            raw = localStorage.getItem(STORAGE_ELIG_FORM_EXCLUSION);
        } catch (e) {
            raw = null;
        }

        // If raw is empty/null, return empty array (don't use default)
        // This allows scanning all SAs when exclusion is not set
        if (!raw || raw.trim() === "" || raw.trim() === "[]") {
            var arr = [];
            log("ImportElig: loaded Form Exclusion = [] (empty, will scan all SAs)");
            return arr;
        }

        var aarr = parseStoredKeywords(raw);

        log("ImportElig: loaded Form Exclusion = " + JSON.stringify(aarr));
        return aarr;
    }

    function setFormExclusion(str) {
        try {
            localStorage.setItem(STORAGE_ELIG_FORM_EXCLUSION, str);
            log("ImportElig: saved Form Exclusion = " + str);
        } catch (e) {
            log("ImportElig: error saving Form Exclusion");
        }
    }

    function getFormPriority() {
        var raw = null;

        try {
            raw = localStorage.getItem(STORAGE_ELIG_FORM_PRIORITY);
        } catch (e) {
            raw = null;
        }

        // If raw is empty/null, return empty array (don't use default)
        // This allows scanning all SAs without priority when priority is not set
        if (!raw || raw.trim() === "" || raw.trim() === "[]") {
            var arr = [];
            log("ImportElig: loaded Form Priority = [] (empty, will scan all SAs)");
            return arr;
        }

        var aarr = parseStoredKeywords(raw);

        log("ImportElig: loaded Form Priority = " + JSON.stringify(aarr));
        return aarr;
    }

    function setFormPriority(str) {
        try {
            localStorage.setItem(STORAGE_ELIG_FORM_PRIORITY, str);
            log("ImportElig: saved Form Priority = " + str);
        } catch (e) {
            log("ImportElig: error saving Form Priority");
        }
    }

    function getFormPriorityOnly() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ELIG_FORM_PRIORITY_ONLY);
        } catch (e) {
            raw = null;
        }
        if (raw === "1") {
            return true;
        }
        return false;
    }

    function setFormPriorityOnly(flag) {
        try {
            localStorage.setItem(STORAGE_ELIG_FORM_PRIORITY_ONLY, flag ? "1" : "0");
            log("ImportElig: saved Form Priority Only = " + String(flag));
        } catch (e) {
            log("ImportElig: error saving Form Priority Only");
        }
    }

    function getPlanPriority() {
        var raw = null;

        try {
            raw = localStorage.getItem(STORAGE_ELIG_PLAN_PRIORITY);
        } catch (e) {
            raw = null;
        }

        // If raw is empty/null, return empty array (don't use default)
        // This allows scanning all plans without priority when priority is not set
        if (!raw || raw.trim() === "" || raw.trim() === "[]") {
            var arr = [];
            log("ImportElig: loaded Plan Priority = [] (empty, will scan all plans)");
            return arr;
        }

        var aarr = parseStoredKeywords(raw);

        log("ImportElig: loaded Plan Priority = " + JSON.stringify(aarr));
        return aarr;
    }

    function setPlanPriority(str) {
        try {
            localStorage.setItem(STORAGE_ELIG_PLAN_PRIORITY, str);
            log("ImportElig: saved Plan Priority = " + str);
        } catch (e) {
            log("ImportElig: error saving Plan Priority");
        }
    }

    function getIgnoreKeywords() {
        var raw = null;

        try {
            raw = localStorage.getItem(STORAGE_ELIG_IGNORE);
        } catch (e) {
            raw = null;
        }

        // If raw is empty/null, return empty array (don't use default)
        // This allows processing all items when ignore is not set
        if (!raw || raw.trim() === "" || raw.trim() === "[]") {
            var arr = [];
            log("ImportElig: loaded Ignore I/E = [] (empty, will process all items)");
            return arr;
        }

        var aarr = parseStoredKeywords(raw);

        log("ImportElig: loaded Ignore I/E = " + JSON.stringify(aarr));
        return aarr;
    }

    function setIgnoreKeywords(str) {
        try {
            localStorage.setItem(STORAGE_ELIG_IGNORE, str);
            log("ImportElig: saved Ignore I/E = " + str);
        } catch (e) {
            log("ImportElig: error saving Ignore I/E");
        }
    }

    function addToIgnoreKeywords(code) {
        if (!code || code.length === 0) {
            return;
        }
        var currentKeywords = getIgnoreKeywords();
        var codeStr = (code + "").trim();
        var codeLower = codeStr.toLowerCase();

        // Check if code is already in the ignore list
        var alreadyIgnored = false;
        var i = 0;
        while (i < currentKeywords.length) {
            var kw = (currentKeywords[i] + "").trim().toLowerCase();
            if (kw === codeLower) {
                alreadyIgnored = true;
                break;
            }
            i = i + 1;
        }

        if (alreadyIgnored) {
            log("ImportElig: code '" + String(codeStr) + "' already in ignore list");
            return;
        }

        // Add the code to the list
        currentKeywords.push(codeStr);

        // Convert back to comma-separated string
        var newStr = currentKeywords.join(", ");
        setIgnoreKeywords(newStr);
        log("ImportElig: added code '" + String(codeStr) + "' to Ignore I/E list");
    }


    function getCheckItemCache() {
        var obj = {};
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ELIG_CHECKITEM_CACHE);
        } catch (e) {
            raw = null;
        }
        if (raw) {
            try {
                var parsed = JSON.parse(raw);
                if (parsed && typeof parsed === "object") {
                    obj = parsed;
                }
            } catch (e2) {}
        }
        return obj;
    }

    function persistCheckItemCache(cache) {
        try {
            localStorage.setItem(STORAGE_ELIG_CHECKITEM_CACHE, JSON.stringify(cache));
        } catch (e) {}
    }

    function getCachedCheckItem(code) {
        var cache = getCheckItemCache();
        if (cache.hasOwnProperty(code)) {
            var rec = cache[code];
            if (rec && typeof rec === "object") {
                log("ImportElig: cache hit for code=" + String(code) + " plan=" + String(rec.plan) + " sched=" + String(rec.sched) + " item=" + String(rec.item));
                return rec;
            }
        }
        log("ImportElig: cache miss for code=" + String(code));
        return null;
    }

    function cacheCheckItem(code, planVal, schedVal, itemVal) {
        var cache = getCheckItemCache();
        cache[String(code)] = { plan: String(planVal), sched: String(schedVal), item: String(itemVal) };
        persistCheckItemCache(cache);
        log("ImportElig: cached mapping for code=" + String(code) + " plan=" + String(planVal) + " sched=" + String(schedVal) + " item=" + String(itemVal));
    }

    async function setSexOptionFromEligibilitySelection(code, valueMap) {
        log("ImportElig: setSexOptionFromEligibilitySelection started for code=" + String(code));
        var selElig = await waitForSelector("select#eligibilityItemRef", 8000);
        if (!selElig) {
            log("ImportElig: select#eligibilityItemRef not found while setting sex");
            return false;
        }
        var val = "";
        if (valueMap && typeof valueMap === "object") {
            if (valueMap.hasOwnProperty(code)) {
                val = String(valueMap[code]);
            }
        }
        if (!val || val.length === 0) {
            var opts = selElig.querySelectorAll("option");
            var i = 0;
            while (i < opts.length) {
                var op = opts[i];
                var v = (op.value + "").trim();
                var t = (op.textContent + "").trim();
                if (v.length > 0) {
                    var parsed = parseItemCodeFromEligibilityOptionText(t);
                    if (parsed === code) {
                        val = v;
                        break;
                    }
                }
                i = i + 1;
            }
        }
        if (!val || val.length === 0) {
            log("ImportElig: cannot resolve option value to parse sex for code=" + String(code));
            return false;
        }
        var opt = selElig.querySelector("option[value='" + CSS.escape(val) + "']");
        if (!opt) {
            log("ImportElig: option element not found for value=" + String(val));
            return false;
        }
        var txt = (opt.textContent + "").trim();
        var sex = parseSexFromEligibilityText(txt);
        var selSex = await waitForSelector("select#sexOption", 8000);
        if (!selSex) {
            log("ImportElig: select#sexOption not found");
            return false;
        }
        var prior = (selSex.value + "");
        selSex.value = sex;
        var evt = new Event("change", { bubbles: true });
        selSex.dispatchEvent(evt);
        log("ImportElig: sexOption set from '" + String(prior) + "' to '" + String(sex) + "'");
        await sleep(150);
        return true;
    }

    function parseSexFromEligibilityText(t) {
        var s = (t + "").toLowerCase();
        var maleTokens = ["male", "man", "men", "boy", "boys", "males"];
        var femaleTokens = ["female", "woman", "women", "girl", "girls", "females"];
        var foundMale = false;
        var foundFemale = false;
        var i = 0;
        while (i < maleTokens.length) {
            var mt = maleTokens[i];
            var reM = new RegExp("\\b" + mt.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&") + "\\b", "i");
            if (reM.test(s)) {
                foundMale = true;
                log("ImportElig: sex parse matched male token=" + mt);
                break;
            }
            i = i + 1;
        }
        var j = 0;
        while (j < femaleTokens.length) {
            var ft = femaleTokens[j];
            var reF = new RegExp("\\b" + ft.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&") + "\\b", "i");
            if (reF.test(s)) {
                foundFemale = true;
                log("ImportElig: sex parse matched female token=" + ft);
                break;
            }
            j = j + 1;
        }
        if (foundMale && foundFemale) {
            log("ImportElig: sex parse result=Both");
            return "Both";
        }
        if (foundMale) {
            log("ImportElig: sex parse result=Male");
            return "Male";
        }
        if (foundFemale) {
            log("ImportElig: sex parse result=Female");
            return "Female";
        }
        log("ImportElig: sex parse result=Both (no tokens)");
        return "Both";
    }

    function parseItemCodeFromEligibilityOptionText(t) {
        var s = (t + "").trim();
        var code = "";
        var parts = s.split("-");
        if (parts.length >= 3) {
            var third = parts[2];
            var semiParts = third.split(";");
            if (semiParts.length >= 1) {
                code = (semiParts[0] + "").trim();
            } else {
                code = (third + "").trim();
            }
        }
        return code;
    }

    async function readEligibilityItemCodesFromSelect() {
        log("ImportElig: opening Eligibility Item from hidden select#eligibilityItemRef");
        var sel = await waitForSelector("select#eligibilityItemRef", 12000);
        if (!sel) {
            log("ImportElig: select#eligibilityItemRef not found");
            return { codes: [], valueMap: {} };
        }
        var opts = sel.querySelectorAll("option");
        var codes = [];
        var valueMap = {};
        var i = 0;
        while (i < opts.length) {
            var op = opts[i];
            var val = (op.value + "").trim();
            var txt = (op.textContent + "").trim();
            if (val.length > 0) {
                var code = parseItemCodeFromEligibilityOptionText(txt);
                if (code && code.length > 0) {
                    codes.push(code);
                    valueMap[code] = val;
                } else {
                    log("ImportElig: skipping option with unparsable text index=" + String(i) + " text=" + JSON.stringify(txt));
                }
            }
            i = i + 1;
        }
        log("ImportElig: total parsed eligibility codes=" + String(codes.length));
        return { codes: codes, valueMap: valueMap };
    }

    async function unlockEligibilityMapping() {
        var actionsBtn = await waitForSelector("button.btn.blue.dropdown-toggle", 8000);
        if (!actionsBtn) {
            log("ImportElig: Actions dropdown button not found; proceeding (control not visible)");
            return true;
        }
        actionsBtn.click();
        await sleep(400);
        var lockToggle = await waitForSelector("a[href='/secure/crfdesign/studylibrary/eligibility/locking'][data-toggle='modal']", 6000);
        if (!lockToggle) {
            log("ImportElig: Lock/Unlock link not found; proceeding");
            return true;
        }
        var labelText = (lockToggle.textContent + "").replace(/\s+/g, " ").trim().toLowerCase();
        var shouldUnlock = labelText.indexOf("unlock") >= 0;
        var shouldLock = labelText.indexOf("lock") >= 0 && !shouldUnlock;
        if (shouldLock) {
            log("ImportElig: mapping already unlocked; skipping unlock step");
            return true;
        }
        if (!shouldUnlock) {
            log("ImportElig: unable to determine lock state from label; proceeding without unlock");
            return true;
        }
        lockToggle.click();
        var modal = await waitForSelector("#ajaxModal .modal-content", 10000);
        if (!modal) {
            log("ImportElig: Unlock modal did not appear");
            return false;
        }
        var reason = await waitForSelector("textarea#reasonForChange", 8000);
        if (!reason) {
            log("ImportElig: reasonForChange textarea not found");
            return false;
        }
        reason.value = "Adding more I/E";
        var evt = new Event("input", { bubbles: true });
        reason.dispatchEvent(evt);
        var saveBtn = await waitForSelector("button#actionButton.btn.green", 8000);
        if (!saveBtn) {
            log("ImportElig: Unlock Save button not found");
            return false;
        }
        try {
            localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ELIG_IMPORT);
        } catch (e) {}
        saveBtn.click();
        await sleep(1000);
        log("ImportElig: unlock flow completed");
        return true;
    }

    function isEligibilityListPage() {
        var path = location.pathname;
        if (path === "/secure/crfdesign/studylibrary/eligibility/list") {
            return true;
        }
        return false;
    }

    function getImportedItemsSet() {
        var set = new Set();
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ELIG_IMPORTED);
        } catch (e) {
            raw = null;
        }
        if (raw) {
            try {
                var arr = JSON.parse(raw);
                if (Array.isArray(arr)) {
                    var i = 0;
                    while (i < arr.length) {
                        set.add(String(arr[i]));
                        i = i + 1;
                    }
                }
            } catch (e2) {
            }
        }
        return set;
    }

    function persistImportedItemsSet(set) {
        try {
            var arr = Array.from(set);
            localStorage.setItem(STORAGE_ELIG_IMPORTED, JSON.stringify(arr));
        } catch (e) {
        }
    }

    function clearImportedItemsSet() {
        try {
            localStorage.removeItem(STORAGE_ELIG_IMPORTED);
        } catch (e) {
        }
    }

    function parseItemCodeFromEligibilityLabel(t) {
        var s = (t + "").trim();
        var code = "";
        var parts = s.split("-");
        if (parts.length >= 3) {
            var third = parts[2];
            var semiParts = third.split(";");
            if (semiParts.length >= 1) {
                code = (semiParts[0] + "").trim();
            } else {
                code = (third + "").trim();
            }
        }
        return code;
    }

    async function collectEligibilityTableMap() {
        var map = {};
        var codeSet = new Set();
        var tbody = await waitForSelector("tbody#eligibilityRefTableBody", 15000);
        if (!tbody) {
            log("ImportElig: tbody#eligibilityRefTableBody not found");
            return { map: map, codeSet: codeSet };
        }
        var rows = tbody.querySelectorAll("tr");
        log("ImportElig: table rows found=" + String(rows.length));
        var i = 0;
        while (i < rows.length) {
            var tr = rows[i];
            var tds = tr.querySelectorAll("td");
            if (tds && tds.length >= 8) {
                var nameTd = tds[0];
                var code = "";
                var a = nameTd.querySelector("a[href*='/secure/crfdesign/studylibrary/show/item/']");
                if (a) {
                    code = (a.textContent + "").trim();
                } else {
                    var fallback = (nameTd.textContent + "").trim().replace(/\s+/g, " ");
                    code = fallback;
                }
                if (code && code.length > 0) {
                    var sex = (tds[1].textContent + "").trim();
                    var cohort = (tds[2].textContent + "").trim();
                    var cohortType = (tds[3].textContent + "").trim();
                    var scheduledActivity = (tds[4].textContent + "").trim().replace(/\s+/g, " ");
                    var checkItem = (tds[5].textContent + "").trim().replace(/\s+/g, " ");
                    var operator = (tds[6].textContent + "").trim();
                    var value = (tds[7].textContent + "").trim();
                    map[code] = {
                        sex: sex,
                        cohort: cohort,
                        cohortType: cohortType,
                        scheduledActivity: scheduledActivity,
                        checkItem: checkItem,
                        operator: operator,
                        value: value
                    };
                    codeSet.add(code);
                    log("ImportElig: stored item=" + JSON.stringify({ code: code, sex: sex, cohort: cohort, cohortType: cohortType, scheduledActivity: scheduledActivity, checkItem: checkItem, operator: operator, value: value }));
                } else {
                    log("ImportElig: skipping row without code at index=" + String(i));
                }
            } else {
                log("ImportElig: skipping row with insufficient columns at index=" + String(i));
            }
            i = i + 1;
        }
        var codesList = Array.from(codeSet);
        log("ImportElig: collected eligibility table rows=" + String(codeSet.size));
        log("ImportElig: collected item codes=" + JSON.stringify(codesList));
        log("ImportElig: full table map=" + JSON.stringify(map));
        return { map: map, codeSet: codeSet };
    }

    async function openAddEligibilityModal() {
        var addBtn = await waitForSelector("a#addEligButton", 8000);
        if (!addBtn) {
            log("ImportElig: Add button #addEligButton not found");
            return false;
        }
        addBtn.click();
        var modal = await waitForSelector("#ajaxModal .modal-content", 10000);
        if (!modal) {
            log("ImportElig: modal did not appear");
            return false;
        }
        return true;
    }

    async function openSelect2(containerSelector) {
        var cont = await waitForSelector(containerSelector + " a.select2-choice", 8000);
        if (!cont) {
            return false;
        }
        cont.click();
        var drop = await waitForSelector("div.select2-drop-active ul.select2-results", 8000);
        if (!drop) {
            return false;
        }
        return true;
    }

    async function readEligibilityItemCodesFromDropdown() {
        var ok = await openSelect2("#s2id_eligibilityItemRef");
        if (!ok) {
            log("ImportElig: unable to open Eligibility Item dropdown");
            return [];
        }
        var list = document.querySelectorAll("div.select2-drop-active ul.select2-results li.select2-result");
        var codes = [];
        var i = 0;
        while (i < list.length) {
            var li = list[i];
            var label = li.querySelector("div.select2-result-label");
            if (label) {
                var text = (label.textContent + "").trim();
                var code = parseItemCodeFromEligibilityLabel(text);
                if (code && code.length > 0) {
                    codes.push(code);
                }
            }
            i = i + 1;
        }
        return codes;
    }


    async function selectEligibilityItemByCode(code, valueMap) {
        log("ImportElig: selecting Eligibility Item via select for code=" + String(code));
        var sel = await waitForSelector("select#eligibilityItemRef", 12000);
        if (!sel) {
            log("ImportElig: select#eligibilityItemRef not found during selection");
            return false;
        }
        var val = "";
        if (valueMap && typeof valueMap === "object") {
            if (valueMap.hasOwnProperty(code)) {
                val = String(valueMap[code]);
            }
        }
        if (!val || val.length === 0) {
            log("ImportElig: valueMap not provided or code missing; rebuilding map");
            var opts = sel.querySelectorAll("option");
            var i = 0;
            while (i < opts.length) {
                var op = opts[i];
                var v = (op.value + "").trim();
                var t = (op.textContent + "").trim();
                if (v.length > 0) {
                    var parsed = parseItemCodeFromEligibilityOptionText(t);
                    if (parsed === code) {
                        val = v;
                        break;
                    }
                }
                i = i + 1;
            }
        }
        if (!val || val.length === 0) {
            log("ImportElig: no option value found for code=" + String(code));
            return false;
        }
        sel.value = val;
        var evt = new Event("change", { bubbles: true });
        sel.dispatchEvent(evt);
        log("ImportElig: Eligibility Item selected value=" + String(val));
        await sleep(300);
        return true;
    }

    // Return index of first option with non-empty value.
    function firstNonEmptyOptionIndex(selectEl) {
        if (!selectEl) {
            return -1;
        }
        var opts = selectEl.querySelectorAll("option");
        var i = 0;
        while (i < opts.length) {
            var v = (opts[i].value + "").trim();
            if (v.length > 0) {
                return i;
            }
            i = i + 1;
        }
        return -1;
    }

    async function trySelectCheckItemForCodeThroughAllPlansAndActivities(code) {
        var formExclusion = getFormExclusion();
        var formPriority = getFormPriority();
        var formPriorityOnly = getFormPriorityOnly();
        var planPriority = getPlanPriority();
        var priorityCheckedSet = new Set();

        function isExcluded(saText) {
            var k = 0;
            while (k < formExclusion.length) {
                var ex = (formExclusion[k] + "").toLowerCase();
                var s = (saText + "").toLowerCase();
                if (ex.length > 0) {
                    if (s.indexOf(ex) >= 0) {
                        return true;
                    }
                }
                k = k + 1;
            }
            return false;
        }

        function isPriority(saText) {
            var k = 0;
            while (k < formPriority.length) {
                var pr = (formPriority[k] + "").toLowerCase();
                var s = (saText + "").toLowerCase();
                if (pr.length > 0) {
                    if (s.indexOf(pr) >= 0) {
                        return true;
                    }
                }
                k = k + 1;
            }
            return false;
        }

        function isPlanPriority(planText) {
            var k = 0;
            while (k < planPriority.length) {
                var pr = (planPriority[k] + "").toLowerCase();
                var s = (planText + "").toLowerCase();
                if (pr.length > 0) {
                    if (s.indexOf(pr) >= 0) {
                        return true;
                    }
                }
                k = k + 1;
            }
            return false;
        }

        log("ImportElig: TrySelect start with priority and exclusion");

        // If Form Priority Only mode is enabled and Form Priority has keywords
        if (formPriorityOnly && formPriority.length > 0) {
            // Scan priority plans first (if they exist), then non-priority plans, but only priority SAs
            if (planPriority.length > 0) {
                var formPriOnlyPlan = await scanActivitiesForMatch(true, true);
                if (formPriOnlyPlan) {
                    log("ImportElig: TrySelect found match during Form Priority Only scan (priority plans)");
                    return true;
                }
            }
            // Scan non-priority plans (or all plans if no plan priority) with priority SAs only
            var formPriOnly = await scanActivitiesForMatch(true, false);
            if (formPriOnly) {
                log("ImportElig: TrySelect found match during Form Priority Only scan");
                return true;
            }
            // No match found in Form Priority Only mode - return false (will be added to ignore list)
            log("ImportElig: TrySelect no match in Form Priority Only mode");
            return false;
        }

        // If plan priorities are set, scan priority plans first
        if (planPriority.length > 0) {
            var pri = await scanActivitiesForMatch(true, true);
            if (pri) {
                log("ImportElig: TrySelect found match during priority plan scan");
                return true;
            }

            var pri2 = await scanActivitiesForMatch(false, true);
            if (pri2) {
                log("ImportElig: TrySelect found match during non-priority SA scan (priority plans)");
                return true;
            }
        }

        // Normal scan: priority SAs first (if Form Priority is set), then all SAs
        if (formPriority.length > 0) {
            var priSA = await scanActivitiesForMatch(true, false);
            if (priSA) {
                log("ImportElig: TrySelect found match during priority SA scan");
                return true;
            }
        }

        var norm = await scanActivitiesForMatch(false, false);
        if (norm) {
            log("ImportElig: TrySelect found match during normal scan");
            return true;
        }

        log("ImportElig: TrySelect no match after all scans");
        return false;


        async function scanActivitiesForMatch(priorityOnly, planPriorityOnly) {
            var planSel = await waitForSelector("select#activityPlan", 10000);
            var schedSel = await waitForSelector("select#scheduledActivity", 10000);
            var itemRefSel = await waitForSelector("select#itemRef", 10000);
            if (!planSel) {
                log("ImportElig: scan planSel missing");
                return false;
            }
            if (!schedSel) {
                log("ImportElig: scan schedSel missing");
                return false;
            }
            if (!itemRefSel) {
                log("ImportElig: scan itemRefSel missing");
                return false;
            }
            var plans = planSel.querySelectorAll("option");
            var pStart = firstNonEmptyOptionIndex(planSel);
            if (pStart < 0) {
                log("ImportElig: scan no activity plan options");
                return false;
            }
            var p = pStart;
            while (p < plans.length) {
                if (isPaused()) {
                    log("ImportElig: scan paused at plan loop");
                    return false;
                }
                if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                    log("ImportElig: run mode cleared (X pressed) at plan loop");
                    return false;
                }
                var pVal = (plans[p].value + "").trim();
                var pTxt = (plans[p].textContent + "").trim();
                if (pVal.length > 0) {
                    if (planPriorityOnly) {
                        var isPlanPri = isPlanPriority(pTxt);
                        if (!isPlanPri) {
                            log("ImportElig: skipping non-priority plan '" + String(pTxt) + "'");
                            p = p + 1;
                            continue;
                        }
                    } else {
                        var isPlanPri2 = isPlanPriority(pTxt);
                        if (isPlanPri2) {
                            log("ImportElig: skipping priority plan '" + String(pTxt) + "' (already scanned)");
                            p = p + 1;
                            continue;
                        }
                    }
                }
                if (pVal.length > 0) {
                    planSel.value = pVal;
                    var evtP = new Event("change", { bubbles: true });
                    planSel.dispatchEvent(evtP);
                    log("ImportElig: plan changed to '" + String(pTxt) + "'");
                    await sleep(500);
                }
                schedSel = document.querySelector("select#scheduledActivity");
                if (!schedSel) {
                    log("ImportElig: scan missing schedSel after plan change");
                    return false;
                }
                var sOps = schedSel.querySelectorAll("option");
                var sStart = firstNonEmptyOptionIndex(schedSel);
                if (sStart < 0) {
                    log("ImportElig: scan no SA under plan '" + String(pTxt) + "'");
                    p = p + 1;
                    continue;
                }
                var s = sStart;
                while (s < sOps.length) {
                    if (isPaused()) {
                        log("ImportElig: scan paused at SA loop");
                        return false;
                    }
                    if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                        log("ImportElig: run mode cleared (X pressed) at SA loop");
                        return false;
                    }
                    var sVal = (sOps[s].value + "").trim();
                    var sTxt = (sOps[s].textContent + "").trim();
                    if (isExcluded(sTxt)) {
                        log("ImportElig: SA excluded '" + String(sTxt) + "'");
                        s = s + 1;
                        continue;
                    }
                    var priSA = isPriority(sTxt);
                    if (priorityOnly && !priSA) {
                        s = s + 1;
                        continue;
                    }
                    if (!priorityOnly && priSA) {
                        if (priorityCheckedSet.has(sTxt)) {
                            log("ImportElig: skipping prioritized SA already checked '" + String(sTxt) + "'");
                            s = s + 1;
                            continue;
                        }
                    }
                    var prevSig = getItemRefOptionsSignature();
                    schedSel.value = sVal;
                    var evtS = new Event("change", { bubbles: true });
                    schedSel.dispatchEvent(evtS);
                    log("ImportElig: Scheduled Activity selected text='" + String(sTxt) + "' value='" + String(sVal) + "'");
                    var reloaded = await waitForItemRefReload(prevSig, 2500);
                    if (!reloaded) {
                        log("ImportElig: itemRef not reloaded for SA '" + String(sTxt) + "', skipping match attempt");
                        if (priorityOnly && priSA) {
                            priorityCheckedSet.add(sTxt);
                            log("ImportElig: marked prioritized SA as checked '" + String(sTxt) + "'");
                        }
                        s = s + 1;
                        continue;
                    }
                    // Verify we're still on the correct SA after itemRef reloaded
                    var verifySAEl = document.querySelector("select#scheduledActivity");
                    if (!verifySAEl || String(verifySAEl.value) !== String(sVal)) {
                        log("ImportElig: SA changed after itemRef reload, expected '" + String(sVal) + "' but got '" + String(verifySAEl ? verifySAEl.value : "null") + "'; skipping");
                        s = s + 1;
                        continue;
                    }
                    // Wait a bit more to ensure Check Item list is fully populated
                    await sleep(300);
                    // Verify again after additional wait
                    verifySAEl = document.querySelector("select#scheduledActivity");
                    if (!verifySAEl || String(verifySAEl.value) !== String(sVal)) {
                        log("ImportElig: SA changed during wait period, expected '" + String(sVal) + "' but got '" + String(verifySAEl ? verifySAEl.value : "null") + "'; skipping");
                        s = s + 1;
                        continue;
                    }
                    var haveOptions = false;
                    var selProbe = document.querySelector("select#itemRef");
                    if (selProbe) {
                        var cnt = selProbe.querySelectorAll("option").length;
                        if (cnt > 0) {
                            haveOptions = true;
                        }
                    }
                    var matched = await attemptCheckItemMatch(code, pVal, sVal);
                    if (matched) {
                        var curSAEl = document.querySelector("select#scheduledActivity");
                        var curSAVal = "";
                        if (curSAEl) {
                            curSAVal = (curSAEl.value + "");
                        }
                        if (String(curSAVal) !== String(sVal)) {
                            log("ImportElig: stale match detected under different SA; discarding");
                            clearLastMatchSelection();
                            matched = false;
                        }
                    }
                    if (!matched) {
                        if (priorityOnly && priSA) {
                            if (haveOptions) {
                                priorityCheckedSet.add(sTxt);
                                log("ImportElig: marked prioritized SA as checked '" + String(sTxt) + "'");
                            } else {
                                log("ImportElig: prioritized SA had no options to check '" + String(sTxt) + "'");
                            }
                        }
                    }
                    if (matched) {
                        log("ImportElig: match found under SA '" + String(sTxt) + "'");
                        return true;
                    }
                    s = s + 1;
                }
                p = p + 1;
            }
            return false;
        }

        async function attemptCheckItemMatch(code, expectedPlanVal, expectedSAVal) {
            var elapsed = 0;
            var step = 160;
            var max = 3200;
            while (elapsed <= max) {
                // Verify we're still on the correct Scheduled Activity
                var saEl = document.querySelector("select#scheduledActivity");
                if (!saEl) {
                    await sleep(step);
                    elapsed = elapsed + step;
                    continue;
                }
                var curSA = (saEl.value + "");
                if (expectedSAVal && String(curSA) !== String(expectedSAVal)) {
                    log("ImportElig: attemptCheckItemMatch SA changed during scan, expected '" + String(expectedSAVal) + "' but got '" + String(curSA) + "'; aborting");
                    return false;
                }
                var planEl = document.querySelector("select#activityPlan");
                if (planEl && expectedPlanVal && String(planEl.value) !== String(expectedPlanVal)) {
                    log("ImportElig: attemptCheckItemMatch Plan changed during scan, expected '" + String(expectedPlanVal) + "' but got '" + String(planEl.value) + "'; aborting");
                    return false;
                }
                var itemRefSel = document.querySelector("select#itemRef");
                if (itemRefSel) {
                    var opts = itemRefSel.querySelectorAll("option");
                    var len = opts.length;
                    if (len === 0) {
                        await sleep(step);
                        elapsed = elapsed + step;
                        continue;
                    }
                    // Verify SA hasn't changed after options are loaded
                    saEl = document.querySelector("select#scheduledActivity");
                    if (!saEl || (expectedSAVal && String(saEl.value) !== String(expectedSAVal))) {
                        log("ImportElig: attemptCheckItemMatch SA changed after options loaded; aborting");
                        return false;
                    }
                    var i = 0;
                    while (i < len) {
                        var op = opts[i];
                        var txt = (op.textContent + "").trim();
                        var val = (op.value + "").trim();
                        if (val.length > 0) {
                            var parts = txt.split("-");
                            if (parts.length >= 2) {
                                var tail = (parts[parts.length - 1] + "").trim();
                                if (tail === code) {
                                    // Final verification before setting value
                                    saEl = document.querySelector("select#scheduledActivity");
                                    if (!saEl || (expectedSAVal && String(saEl.value) !== String(expectedSAVal))) {
                                        log("ImportElig: attemptCheckItemMatch SA changed before setting match; aborting");
                                        return false;
                                    }
                                    itemRefSel.value = val;
                                    var evt = new Event("change", { bubbles: true });
                                    itemRefSel.dispatchEvent(evt);
                                    log("ImportElig: Check Item matched '" + String(txt) + "'");
                                    var planSelCap = document.querySelector("select#activityPlan");
                                    var schedSelCap = document.querySelector("select#scheduledActivity");
                                    var planVal = "";
                                    var saVal = "";
                                    if (planSelCap) {
                                        planVal = (planSelCap.value + "");
                                    }
                                    if (schedSelCap) {
                                        saVal = (schedSelCap.value + "");
                                    }
                                    setLastMatchSelection(planVal, saVal, val);
                                    return true;
                                }
                            }
                        }
                        i = i + 1;
                    }
                    // No match found, but we verified we're on the correct SA
                    return false;
                }
                await sleep(step);
                elapsed = elapsed + step;
            }
            return false;
        }
    }


    async function setComparatorEQ() {
        log("ImportElig: setComparatorEQ started");
        var sel = await waitForComparatorReady(1000);
        if (!sel) {
            log("ImportElig: comparator select not found or not ready");
            return false;
        }
        var attempt = 0;
        var maxAttempts = 10;
        var delay = 400;
        while (attempt < maxAttempts) {
            sel = document.querySelector("select#eligibilityComparator");
            if (!sel) {
                log("ImportElig: comparator disappeared, retrying");
                await sleep(delay);
                attempt = attempt + 1;
                continue;
            }
            var op = sel.querySelector("option[value='EQ']");
            if (op) {
                sel.value = "EQ";
                var evt = new Event("change", { bubbles: true });
                sel.dispatchEvent(evt);
                log("ImportElig: comparator set to EQ");
                await sleep(350);
                return true;
            }
            log("ImportElig: EQ not found (attempt " + String(attempt + 1) + "/" + String(maxAttempts) + ")");
            await sleep(delay);
            attempt = attempt + 1;
        }
        log("ImportElig: comparator EQ option not found after all retries");
        return false;
    }



    async function selectCodeListValueContainingSF() {
        var sel = await waitForSelector("select#codeListItem", 8000);
        if (!sel) {
            log("ImportElig: codeListItem select not found");
            return false;
        }
        var options = sel.querySelectorAll("option");
        var chosen = null;
        var i = 0;
        while (i < options.length) {
            var o = options[i];
            var txt = (o.textContent + "").trim();
            var val = (o.value + "").trim();
            if (val.length > 0) {
                if (txt.indexOf("SF") >= 0) {
                    chosen = o;
                    break;
                }
            }
            i = i + 1;
        }
        if (!chosen) {
            log("ImportElig: no Code List option containing 'SF' found");
            return false;
        }
        sel.value = chosen.value;
        var evt = new Event("change", { bubbles: true });
        sel.dispatchEvent(evt);
        log("ImportElig: Code List value set to '" + String((chosen.textContent + "").trim()) + "'");
        await sleep(200);
        return true;
    }

    async function clickSaveAndWait() {
        var btn = await waitForSelector("button#actionButton.btn.green", 8000);
        if (!btn) {
            log("ImportElig: Save button not found");
            return false;
        }
        btn.click();
        log("ImportElig: clicked Save");
        await sleep(2000);
        return true;
    }



    function importIERandomDelay() {
        var range = IMPORT_IE_SHORT_DELAY_MAX - IMPORT_IE_SHORT_DELAY_MIN;
        var val = IMPORT_IE_SHORT_DELAY_MIN + Math.floor(Math.random() * range);
        return val;
    }

    async function importIEDelay(ms) {
        log("ImportIE: delay " + String(ms) + "ms");
        await sleep(ms);
    }

    async function waitForElement(selector, timeoutMs) {
        log("ImportIE: waitForElement selector='" + String(selector) + "' timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_HELPER_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var el = document.querySelector(selector);
            if (el) {
                log("ImportIE: waitForElement found selector='" + String(selector) + "' in " + String(Date.now() - start) + "ms");
                return el;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForElement timeout selector='" + String(selector) + "'");
        return null;
    }

    async function waitForSelectOptions(selectEl, minCount, timeoutMs) {
        log("ImportIE: waitForSelectOptions minCount=" + String(minCount) + " timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_HELPER_TIMEOUT;
        var minC = typeof minCount === "number" ? minCount : 1;
        while (Date.now() - start < maxMs) {
            if (selectEl) {
                var opts = selectEl.querySelectorAll("option");
                var nonEmpty = 0;
                var oi = 0;
                while (oi < opts.length) {
                    var v = (opts[oi].value + "").trim();
                    if (v.length > 0) {
                        nonEmpty = nonEmpty + 1;
                    }
                    oi = oi + 1;
                }
                if (nonEmpty >= minC) {
                    log("ImportIE: waitForSelectOptions got " + String(nonEmpty) + " options in " + String(Date.now() - start) + "ms");
                    return true;
                }
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForSelectOptions timeout");
        return false;
    }

    async function waitForModalOpen(timeoutMs) {
        log("ImportIE: waitForModalOpen timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_MODAL_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var modal = document.querySelector("#ajaxModal .modal-content");
            if (modal) {
                var header = modal.querySelector(".modal-header");
                if (header) {
                    var txt = (header.textContent + "").trim();
                    if (txt.indexOf("Eligibility Management") >= 0) {
                        log("ImportIE: waitForModalOpen modal found in " + String(Date.now() - start) + "ms");
                        return true;
                    }
                }
                log("ImportIE: waitForModalOpen modal content found in " + String(Date.now() - start) + "ms");
                return true;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForModalOpen timeout");
        return false;
    }

    async function waitForModalClose(timeoutMs) {
        log("ImportIE: waitForModalClose timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_MODAL_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var modal = document.querySelector("#ajaxModal .modal-content");
            if (!modal) {
                log("ImportIE: waitForModalClose modal closed in " + String(Date.now() - start) + "ms");
                return true;
            }
            var display = window.getComputedStyle(modal.closest("#ajaxModal") || modal).display;
            if (display === "none") {
                log("ImportIE: waitForModalClose modal hidden in " + String(Date.now() - start) + "ms");
                return true;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForModalClose timeout");
        return false;
    }

    async function clickWithRetry(elOrSelector) {
        log("ImportIE: clickWithRetry start");
        var el = null;
        if (typeof elOrSelector === "string") {
            el = document.querySelector(elOrSelector);
        } else {
            el = elOrSelector;
        }
        if (!el) {
            log("ImportIE: clickWithRetry element not found on first attempt");
            await sleep(500);
            if (typeof elOrSelector === "string") {
                el = document.querySelector(elOrSelector);
            }
            if (!el) {
                log("ImportIE: clickWithRetry element not found on retry");
                return false;
            }
        }
        try {
            el.click();
            log("ImportIE: clickWithRetry clicked");
            return true;
        } catch (e) {
            log("ImportIE: clickWithRetry first click error=" + String(e));
            await sleep(300);
            try {
                el.click();
                log("ImportIE: clickWithRetry retry click succeeded");
                return true;
            } catch (e2) {
                log("ImportIE: clickWithRetry retry click failed=" + String(e2));
                return false;
            }
        }
    }

    function select2TriggerChange(selectEl) {
        log("ImportIE: select2TriggerChange");
        var evt = new Event("change", { bubbles: true });
        selectEl.dispatchEvent(evt);
        if (typeof jQuery !== "undefined" && jQuery && jQuery.fn) {
            try {
                jQuery(selectEl).trigger("change");
                log("ImportIE: select2TriggerChange jQuery trigger fired");
            } catch (e) {
                log("ImportIE: select2TriggerChange jQuery trigger error=" + String(e));
            }
        }
    }

    async function select2SelectByValue(containerOrSelect, value) {
        log("ImportIE: select2SelectByValue value='" + String(value) + "'");
        var sel = null;
        if (containerOrSelect && containerOrSelect.tagName && containerOrSelect.tagName.toLowerCase() === "select") {
            sel = containerOrSelect;
        } else if (typeof containerOrSelect === "string") {
            sel = document.querySelector(containerOrSelect);
        } else {
            var under = containerOrSelect.querySelector("select");
            if (under) {
                sel = under;
            }
        }
        if (!sel) {
            log("ImportIE: select2SelectByValue select element not found");
            return false;
        }
        sel.value = String(value);
        select2TriggerChange(sel);
        await sleep(importIERandomDelay());
        var after = (sel.value + "").trim();
        if (after === String(value).trim()) {
            log("ImportIE: select2SelectByValue confirmed value='" + String(after) + "'");
            return true;
        }
        log("ImportIE: select2SelectByValue value mismatch expected='" + String(value) + "' got='" + String(after) + "'");
        sel.value = String(value);
        select2TriggerChange(sel);
        await sleep(300);
        after = (sel.value + "").trim();
        log("ImportIE: select2SelectByValue retry result='" + String(after) + "'");
        return after === String(value).trim();
    }

    async function select2SelectByText(containerOrSelect, text) {
        log("ImportIE: select2SelectByText text='" + String(text) + "'");
        var sel = null;
        if (containerOrSelect && containerOrSelect.tagName && containerOrSelect.tagName.toLowerCase() === "select") {
            sel = containerOrSelect;
        } else if (typeof containerOrSelect === "string") {
            sel = document.querySelector(containerOrSelect);
        } else {
            var under = containerOrSelect.querySelector("select");
            if (under) {
                sel = under;
            }
        }
        if (!sel) {
            log("ImportIE: select2SelectByText select element not found");
            return false;
        }
        var opts = sel.querySelectorAll("option");
        var normalTarget = normalizeTextForCompare(text);
        var oi = 0;
        while (oi < opts.length) {
            var oTxt = normalizeTextForCompare((opts[oi].textContent + ""));
            if (oTxt === normalTarget) {
                sel.value = opts[oi].value;
                select2TriggerChange(sel);
                await sleep(importIERandomDelay());
                log("ImportIE: select2SelectByText matched exact text index=" + String(oi));
                return true;
            }
            oi = oi + 1;
        }
        log("ImportIE: select2SelectByText no exact match for text='" + String(text) + "'");
        return false;
    }

    function normalizeTextForCompare(t) {
        var s = (t + "").trim();
        s = s.replace(/\s+/g, " ");
        s = s.toLowerCase();
        return s;
    }

    function extractIECode(text) {
        var s = (text + "").trim();
        var match = IE_CODE_REGEX.exec(s);
        if (match) {
            var prefix = match[1].toUpperCase();
            var num = match[2];
            var code = prefix + num;
            return code;
        }
        return "";
    }

    function extractIECodeStrict(text) {
        var s = (text + "").trim();
        var re = /\b(INC|EXC)(\d+)\b/i;
        var match = re.exec(s);
        if (match) {
            var prefix = match[1].toUpperCase();
            var num = match[2];
            return prefix + num;
        }
        var re2 = /\b(INC|EXC)\s+(\d+)\b/i;
        var match2 = re2.exec(s);
        if (match2) {
            var prefix2 = match2[1].toUpperCase();
            var num2 = match2[2];
            return prefix2 + num2;
        }
        return "";
    }

    function isValidEligibilityPage() {
        log("ImportIE: isValidEligibilityPage checking");
        var hostname = location.hostname;
        var path = location.pathname;
        var validHost = false;
        var hi = 0;
        while (hi < ELIGIBILITY_VALID_HOSTNAMES.length) {
            if (hostname === ELIGIBILITY_VALID_HOSTNAMES[hi]) {
                validHost = true;
                break;
            }
            hi = hi + 1;
        }
        if (!validHost) {
            log("ImportIE: invalid hostname='" + String(hostname) + "'");
            return false;
        }
        if (path !== ELIGIBILITY_LIST_PATH) {
            log("ImportIE: invalid path='" + String(path) + "'");
            return false;
        }
        log("ImportIE: valid eligibility page");
        return true;
    }

    function getBaseUrl() {
        return location.protocol + "//" + location.hostname;
    }

    function showWarningPopup(title, message) {
        log("ImportIE: showWarningPopup title='" + String(title) + "'");
        var msgEl = document.createElement("div");
        msgEl.style.textAlign = "center";
        msgEl.style.fontSize = "15px";
        msgEl.style.color = "#fff";
        msgEl.style.padding = "20px";
        msgEl.style.lineHeight = "1.6";
        msgEl.textContent = message;
        createPopup({
            title: title,
            content: msgEl,
            width: "460px",
            height: "auto"
        });
    }

    async function collectAllTableCodes() {
        log("ImportIE: collectAllTableCodes start");
        var codeSet = new Set();
        var tbody = document.querySelector("tbody#eligibilityreftablebody");
        if (!tbody) {
            tbody = document.querySelector("tbody#eligibilityRefTableBody");
            if (tbody) {
                log("ImportIE: collectAllTableCodes fallback selector tbody#eligibilityRefTableBody matched");
            }
        }
        if (!tbody) {
            tbody = document.querySelector("#eligibilityRefTable tbody");
            if (tbody) {
                log("ImportIE: collectAllTableCodes fallback selector #eligibilityRefTable tbody matched");
            }
        }
        if (!tbody) {
            log("ImportIE: collectAllTableCodes no table body found");
            return codeSet;
        }

        var paginationExists = document.querySelector(".pagination") || document.querySelector("[data-page]") || document.querySelector(".dataTables_paginate");
        if (paginationExists) {
            log("ImportIE: collectAllTableCodes pagination detected, iterating pages");
            var maxPages = 50;
            var pageNum = 0;
            while (pageNum < maxPages) {
                var rows = tbody.querySelectorAll("tr");
                var ri = 0;
                while (ri < rows.length) {
                    var tr = rows[ri];
                    var anchor = tr.querySelector("a[href*='/secure/crfdesign/studylibrary/show/item/']");
                    if (anchor) {
                        var anchorText = (anchor.textContent + "").trim();
                        var code = extractIECodeStrict(anchorText);
                        if (code.length > 0) {
                            codeSet.add(code.toUpperCase());
                        } else {
                            var cleaned = anchorText.replace(/\s+/g, "").toUpperCase();
                            if (cleaned.length > 0) {
                                codeSet.add(cleaned);
                            }
                        }
                    } else {
                        var tds = tr.querySelectorAll("td");
                        if (tds && tds.length > 0) {
                            var firstText = (tds[0].textContent + "").trim().replace(/\s+/g, " ");
                            var code2 = extractIECodeStrict(firstText);
                            if (code2.length > 0) {
                                codeSet.add(code2.toUpperCase());
                            }
                        }
                    }
                    ri = ri + 1;
                }
                var nextBtn = document.querySelector(".pagination .next:not(.disabled) a");
                if (!nextBtn) {
                    nextBtn = document.querySelector(".dataTables_paginate .next:not(.disabled)");
                }
                if (!nextBtn) {
                    log("ImportIE: collectAllTableCodes no more pages");
                    break;
                }
                nextBtn.click();
                await sleep(1500);
                tbody = document.querySelector("tbody#eligibilityreftablebody") || document.querySelector("tbody#eligibilityRefTableBody") || document.querySelector("#eligibilityRefTable tbody");
                if (!tbody) {
                    log("ImportIE: collectAllTableCodes tbody lost after page change");
                    break;
                }
                pageNum = pageNum + 1;
            }
        } else {
            var loadMoreBtn = document.querySelector(".load-more, [data-load-more]");
            if (loadMoreBtn) {
                log("ImportIE: collectAllTableCodes load-more detected");
                var clickCount = 0;
                var maxClicks = 50;
                while (clickCount < maxClicks) {
                    loadMoreBtn.click();
                    await sleep(1500);
                    loadMoreBtn = document.querySelector(".load-more, [data-load-more]");
                    if (!loadMoreBtn) {
                        break;
                    }
                    clickCount = clickCount + 1;
                }
            }
            var rows2 = tbody.querySelectorAll("tr");
            log("ImportIE: collectAllTableCodes rows=" + String(rows2.length));
            var ri2 = 0;
            while (ri2 < rows2.length) {
                var tr2 = rows2[ri2];
                var anchor2 = tr2.querySelector("a[href*='/secure/crfdesign/studylibrary/show/item/']");
                if (anchor2) {
                    var anchorText2 = (anchor2.textContent + "").trim();
                    var code3 = extractIECodeStrict(anchorText2);
                    if (code3.length > 0) {
                        codeSet.add(code3.toUpperCase());
                    } else {
                        var cleaned2 = anchorText2.replace(/\s+/g, "").toUpperCase();
                        if (cleaned2.length > 0) {
                            codeSet.add(cleaned2);
                        }
                    }
                } else {
                    var tds2 = tr2.querySelectorAll("td");
                    if (tds2 && tds2.length > 0) {
                        var firstText2 = (tds2[0].textContent + "").trim().replace(/\s+/g, " ");
                        var code4 = extractIECodeStrict(firstText2);
                        if (code4.length > 0) {
                            codeSet.add(code4.toUpperCase());
                        }
                    }
                }
                ri2 = ri2 + 1;
            }
        }
        log("ImportIE: collectAllTableCodes done count=" + String(codeSet.size));
        log("ImportIE: collectAllTableCodes codes=" + JSON.stringify(Array.from(codeSet)));
        return codeSet;
    }

    async function collectMappingsFromModal(existingCodeSet) {
        log("ImportIE: collectMappingsFromModal start");
        var mappings = [];
        var planSel = document.querySelector("select#activityPlan");
        if (!planSel) {
            planSel = await waitForElement("select#activityPlan", 10000);
        }
        if (!planSel) {
            log("ImportIE: collectMappingsFromModal planSel not found");
            return mappings;
        }
        var planOpts = planSel.querySelectorAll("option");
        var pi = 0;
        while (pi < planOpts.length) {
            var pVal = (planOpts[pi].value + "").trim();
            var pTxt = (planOpts[pi].textContent + "").trim().replace(/\s+/g, " ");
            if (pVal.length === 0) {
                pi = pi + 1;
                continue;
            }
            log("ImportIE: collectMappingsFromModal selecting plan='" + String(pTxt) + "' value='" + String(pVal) + "'");
            planSel.value = pVal;
            select2TriggerChange(planSel);
            await sleep(200);
            var schedSel = document.querySelector("select#scheduledActivity");
            if (!schedSel) {
                schedSel = await waitForElement("select#scheduledActivity", 8000);
            }
            if (!schedSel) {
                log("ImportIE: collectMappingsFromModal schedSel not found for plan='" + String(pTxt) + "'");
                pi = pi + 1;
                continue;
            }
            var hasSchedOpts = await waitForSelectOptions(schedSel, 1, 3000);
            if (!hasSchedOpts) {
                log("ImportIE: collectMappingsFromModal no SA options for plan='" + String(pTxt) + "'");
                pi = pi + 1;
                continue;
            }
            var schedOpts = schedSel.querySelectorAll("option");
            var si = 0;
            while (si < schedOpts.length) {
                var sVal = (schedOpts[si].value + "").trim();
                var sTxt = (schedOpts[si].textContent + "").trim().replace(/\s+/g, " ");
                if (sVal.length === 0) {
                    si = si + 1;
                    continue;
                }
                log("ImportIE: collectMappingsFromModal selecting SA='" + String(sTxt) + "' value='" + String(sVal) + "'");
                var prevItemSig = getItemRefOptionsSignature();
                schedSel.value = sVal;
                select2TriggerChange(schedSel);
                await sleep(100);
                var itemRefSel = document.querySelector("select#itemRef");
                if (!itemRefSel) {
                    itemRefSel = await waitForElement("select#itemRef", 8000);
                }
                if (!itemRefSel) {
                    log("ImportIE: collectMappingsFromModal itemRefSel not found for SA='" + String(sTxt) + "'");
                    si = si + 1;
                    continue;
                }
                var reloaded = await waitForItemRefReload(prevItemSig, 2000);
                if (!reloaded) {
                    log("ImportIE: collectMappingsFromModal itemRef did not reload for SA='" + String(sTxt) + "', checking current options");
                }
                await sleep(100);
                itemRefSel = document.querySelector("select#itemRef");
                if (!itemRefSel) {
                    log("ImportIE: collectMappingsFromModal itemRefSel lost after wait");
                    si = si + 1;
                    continue;
                }
                var itemOpts = itemRefSel.querySelectorAll("option");
                var ii = 0;
                while (ii < itemOpts.length) {
                    var iVal = (itemOpts[ii].value + "").trim();
                    var iTxt = (itemOpts[ii].textContent + "").trim().replace(/\s+/g, " ");
                    if (iVal.length === 0) {
                        ii = ii + 1;
                        continue;
                    }
                    var ieCode = extractIECodeStrict(iTxt);
                    if (ieCode.length === 0) {
                        ieCode = extractIECode(iTxt);
                    }
                    if (ieCode.length > 0) {
                        var record = {
                            activityPlanText: pTxt,
                            scheduledActivityText: sTxt,
                            checkItemText: iTxt,
                            code: ieCode.toUpperCase(),
                            ids: {
                                activityPlanValue: pVal,
                                scheduledActivityValue: sVal,
                                checkItemValue: iVal
                            }
                        };
                        log("ImportIE: collectMappingsFromModal found mapping code=" + String(record.code) + " plan='" + String(pTxt) + "' sa='" + String(sTxt) + "' item='" + String(iTxt) + "'");
                        mappings.push(record);
                    }
                    ii = ii + 1;
                }
                si = si + 1;
            }
            pi = pi + 1;
        }
        log("ImportIE: collectMappingsFromModal done total=" + String(mappings.length));
        return mappings;
    }

    function deduplicateMappings(mappings) {
        log("ImportIE: deduplicateMappings start count=" + String(mappings.length));
        var seen = {};
        var result = [];
        var mi = 0;
        while (mi < mappings.length) {
            var m = mappings[mi];
            var key = m.code + "|" + m.ids.activityPlanValue + "|" + m.ids.scheduledActivityValue + "|" + m.ids.checkItemValue;
            if (!seen.hasOwnProperty(key)) {
                seen[key] = true;
                result.push(m);
            }
            mi = mi + 1;
        }
        log("ImportIE: deduplicateMappings done count=" + String(result.length));
        return result;
    }

    function buildHierarchy(mappings) {
        log("ImportIE: buildHierarchy start");
        var plans = [];
        var planMap = {};
        var mi = 0;
        while (mi < mappings.length) {
            var m = mappings[mi];
            var pKey = m.ids.activityPlanValue;
            if (!planMap.hasOwnProperty(pKey)) {
                planMap[pKey] = {
                    text: m.activityPlanText,
                    value: pKey,
                    activities: [],
                    activityMap: {}
                };
                plans.push(planMap[pKey]);
            }
            var plan = planMap[pKey];
            var sKey = m.ids.scheduledActivityValue;
            if (!plan.activityMap.hasOwnProperty(sKey)) {
                plan.activityMap[sKey] = {
                    text: m.scheduledActivityText,
                    value: sKey,
                    items: []
                };
                plan.activities.push(plan.activityMap[sKey]);
            }
            plan.activityMap[sKey].items.push(m);
            mi = mi + 1;
        }
        log("ImportIE: buildHierarchy done plans=" + String(plans.length));
        return plans;
    }

    function buildImportIEReviewPanel(existingCodeSet, mappings, onConfirm) {
        log("ImportIE: buildImportIEReviewPanel start existingCodes=" + String(existingCodeSet.size) + " mappings=" + String(mappings.length));
        var hierarchy = buildHierarchy(mappings);
        var existingArr = Array.from(existingCodeSet);
        existingArr.sort();

        var overlay = document.createElement("div");
        overlay.id = "importIEReviewOverlay";
        overlay.style.position = "fixed";
        overlay.style.top = "0";
        overlay.style.left = "0";
        overlay.style.width = "100%";
        overlay.style.height = "100%";
        overlay.style.zIndex = "999997";
        overlay.style.background = "rgba(0,0,0,0.7)";
        overlay.style.backdropFilter = "blur(3px)";
        overlay.style.display = "flex";
        overlay.style.alignItems = "center";
        overlay.style.justifyContent = "center";
        overlay.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        overlay.style.fontSize = "13px";
        overlay.style.color = "#fff";

        var container = document.createElement("div");
        container.style.background = "#111";
        container.style.border = "1px solid #444";
        container.style.borderRadius = "8px";
        container.style.width = "90%";
        container.style.maxWidth = "1100px";
        container.style.height = "85vh";
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.boxShadow = "0 8px 32px rgba(0,0,0,0.6)";

        var headerBar = document.createElement("div");
        headerBar.style.display = "flex";
        headerBar.style.alignItems = "center";
        headerBar.style.justifyContent = "space-between";
        headerBar.style.padding = "14px 18px";
        headerBar.style.borderBottom = "1px solid #333";
        headerBar.style.flexShrink = "0";
        headerBar.style.background = "linear-gradient(180deg, #1a1a1a, #111)";

        var headerTitle = document.createElement("div");
        headerTitle.textContent = "Import I/E - Review Mappings";
        headerTitle.style.fontWeight = "600";
        headerTitle.style.fontSize = "16px";

        var headerCloseBtn = document.createElement("button");
        headerCloseBtn.textContent = "\u2715";
        headerCloseBtn.style.background = "transparent";
        headerCloseBtn.style.color = "#fff";
        headerCloseBtn.style.border = "none";
        headerCloseBtn.style.cursor = "pointer";
        headerCloseBtn.style.fontSize = "18px";
        headerCloseBtn.style.padding = "4px 8px";
        headerCloseBtn.style.borderRadius = "4px";
        headerCloseBtn.style.transition = "background 0.15s";
        headerCloseBtn.addEventListener("mouseenter", function () {
            headerCloseBtn.style.background = "#333";
        });
        headerCloseBtn.addEventListener("mouseleave", function () {
            headerCloseBtn.style.background = "transparent";
        });
        headerCloseBtn.addEventListener("click", function () {
            log("ImportIE: review panel closed by user");
            overlay.remove();
        });

        headerBar.appendChild(headerTitle);
        headerBar.appendChild(headerCloseBtn);
        container.appendChild(headerBar);

        var panelBody = document.createElement("div");
        panelBody.style.display = "flex";
        panelBody.style.flex = "1";
        panelBody.style.overflow = "hidden";

        var leftPanel = document.createElement("div");
        leftPanel.style.width = "280px";
        leftPanel.style.flexShrink = "0";
        leftPanel.style.borderRight = "1px solid #444";
        leftPanel.style.display = "flex";
        leftPanel.style.flexDirection = "column";
        leftPanel.style.padding = "12px";

        var leftTitle = document.createElement("div");
        leftTitle.textContent = "Existing Table Codes (" + String(existingArr.length) + ")";
        leftTitle.style.fontWeight = "600";
        leftTitle.style.marginBottom = "8px";
        leftTitle.style.fontSize = "13px";
        leftPanel.appendChild(leftTitle);

        var leftSearch = document.createElement("input");
        leftSearch.type = "text";
        leftSearch.placeholder = "Search codes...";
        leftSearch.style.width = "100%";
        leftSearch.style.padding = "6px 8px";
        leftSearch.style.marginBottom = "8px";
        leftSearch.style.background = "#222";
        leftSearch.style.color = "#fff";
        leftSearch.style.border = "1px solid #555";
        leftSearch.style.borderRadius = "4px";
        leftSearch.style.boxSizing = "border-box";
        leftSearch.style.fontSize = "12px";
        leftPanel.appendChild(leftSearch);

        var leftList = document.createElement("div");
        leftList.style.flex = "1";
        leftList.style.overflowY = "auto";

        function renderLeftList(filter) {
            leftList.innerHTML = "";
            var fi = 0;
            while (fi < existingArr.length) {
                var c = existingArr[fi];
                if (filter && filter.length > 0) {
                    if (c.toLowerCase().indexOf(filter.toLowerCase()) < 0) {
                        fi = fi + 1;
                        continue;
                    }
                }
                var row = document.createElement("div");
                row.style.padding = "3px 6px";
                row.style.fontSize = "12px";
                row.style.color = "#aaa";
                row.textContent = c;
                leftList.appendChild(row);
                fi = fi + 1;
            }
        }
        renderLeftList("");
        leftSearch.addEventListener("input", function () {
            renderLeftList(leftSearch.value);
        });
        leftPanel.appendChild(leftList);
        panelBody.appendChild(leftPanel);

        var rightPanel = document.createElement("div");
        rightPanel.style.flex = "1";
        rightPanel.style.display = "flex";
        rightPanel.style.flexDirection = "column";
        rightPanel.style.padding = "12px";
        rightPanel.style.overflow = "hidden";

        var rightTitle = document.createElement("div");
        rightTitle.textContent = "Mapped Check Items";
        rightTitle.style.fontWeight = "600";
        rightTitle.style.marginBottom = "8px";
        rightTitle.style.fontSize = "13px";
        rightPanel.appendChild(rightTitle);

        var rightSearch = document.createElement("input");
        rightSearch.type = "text";
        rightSearch.placeholder = "Search mappings...";
        rightSearch.style.width = "100%";
        rightSearch.style.padding = "6px 8px";
        rightSearch.style.marginBottom = "8px";
        rightSearch.style.background = "#222";
        rightSearch.style.color = "#fff";
        rightSearch.style.border = "1px solid #555";
        rightSearch.style.borderRadius = "4px";
        rightSearch.style.boxSizing = "border-box";
        rightSearch.style.fontSize = "12px";
        rightPanel.appendChild(rightSearch);

        var rightList = document.createElement("div");
        rightList.style.flex = "1";
        rightList.style.overflowY = "auto";
        rightList.style.paddingRight = "4px";

        var allCheckboxes = [];
        var planCheckboxes = [];
        var saCheckboxes = [];
        var itemCheckboxes = [];

        function updateCounter() {
            var count = 0;
            var ci = 0;
            while (ci < itemCheckboxes.length) {
                if (itemCheckboxes[ci].cb.checked && !itemCheckboxes[ci].cb.disabled) {
                    count = count + 1;
                }
                ci = ci + 1;
            }
            counterEl.textContent = "Selected: " + String(count);
            if (count > 0) {
                confirmBtn.disabled = false;
                confirmBtn.style.opacity = "1";
            } else {
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = "0.5";
            }
        }

        function updateParentState(planIdx) {
            updateCounter();
        }

        // State tracking
        var selectedItems = {};
        var selectedSAs = {};
        var selectedPlans = {};

        // Helper functions
        function toggleItem(itemIdx, itemEntry) {
            var key = "item_" + itemIdx;
            if (selectedItems[key]) {
                delete selectedItems[key];
                itemEntry.cb.checked = false;
            } else {
                selectedItems[key] = itemEntry;
                itemEntry.cb.checked = true;
            }
            updateCounter();
        }

        function toggleSA(saIdx, saEntry) {
            var key = "sa_" + saIdx;
            if (selectedSAs[key]) {
                delete selectedSAs[key];
                saEntry.cb.checked = false;
                // Uncheck all items under this SA
                for (var i = 0; i < saEntry.itemIndices.length; i++) {
                    var itIdx = saEntry.itemIndices[i];
                    var itKey = "item_" + itIdx;
                    delete selectedItems[itKey];
                    itemCheckboxes[itIdx].cb.checked = false;
                }
            } else {
                selectedSAs[key] = saEntry;
                saEntry.cb.checked = true;
                // Check all items under this SA
                for (var j = 0; j < saEntry.itemIndices.length; j++) {
                    var itIdx = saEntry.itemIndices[j];
                    var itEntry = itemCheckboxes[itIdx];
                    if (!itEntry.cb.disabled) {
                        var itKey = "item_" + itIdx;
                        selectedItems[itKey] = itEntry;
                        itEntry.cb.checked = true;
                    }
                }
            }
            updateCounter();
        }

        function togglePlan(planIdx, planEntry) {
            var key = "plan_" + planIdx;
            if (selectedPlans[key]) {
                delete selectedPlans[key];
                planEntry.cb.checked = false;
                // Uncheck all SAs and items under this plan
                for (var i = 0; i < planEntry.saIndices.length; i++) {
                    var saIdx = planEntry.saIndices[i];
                    var saKey = "sa_" + saIdx;
                    delete selectedSAs[saKey];
                    saCheckboxes[saIdx].cb.checked = false;
                    // Uncheck all items under this SA
                    var saEntry = saCheckboxes[saIdx];
                    for (var j = 0; j < saEntry.itemIndices.length; j++) {
                        var itIdx = saEntry.itemIndices[j];
                        var itKey = "item_" + itIdx;
                        delete selectedItems[itKey];
                        itemCheckboxes[itIdx].cb.checked = false;
                    }
                }
            } else {
                selectedPlans[key] = planEntry;
                planEntry.cb.checked = true;
                // Check all SAs and items under this plan
                for (var k = 0; k < planEntry.saIndices.length; k++) {
                    var saIdx = planEntry.saIndices[k];
                    var saEntry = saCheckboxes[saIdx];
                    if (!saEntry.cb.disabled) {
                        var saKey = "sa_" + saIdx;
                        selectedSAs[saKey] = saEntry;
                        saEntry.cb.checked = true;
                        // Check all items under this SA
                        for (var l = 0; l < saEntry.itemIndices.length; l++) {
                            var itIdx = saEntry.itemIndices[l];
                            var itEntry = itemCheckboxes[itIdx];
                            if (!itEntry.cb.disabled) {
                                var itKey = "item_" + itIdx;
                                selectedItems[itKey] = itEntry;
                                itEntry.cb.checked = true;
                            }
                        }
                    }
                }
            }
            updateCounter();
        }

        function setDescendants(planIdx, checked) {
            log("ImportIE: setDescendants planIdx=" + planIdx + " checked=" + checked);
            var pEntry = planCheckboxes[planIdx];
            if (!pEntry) {
                log("ImportIE: setDescendants - plan entry not found");
                return;
            }
            // Only affect SA checkboxes, don't force their state permanently
            var sai = 0;
            while (sai < pEntry.saIndices.length) {
                var saIdx = pEntry.saIndices[sai];
                var saEntry = saCheckboxes[saIdx];
                if (saEntry && saEntry.cb) {
                    // Only set SA checkbox if it's not disabled
                    if (!saEntry.cb.disabled) {
                        saEntry.cb.checked = checked;
                        saEntry.cb.indeterminate = false;
                        log("ImportIE: setDescendants - set SA " + saIdx + " to " + checked);
                        // Now affect the items under this SA
                        var iti = 0;
                        while (iti < saEntry.itemIndices.length) {
                            var itIdx = saEntry.itemIndices[iti];
                            var itEntry = itemCheckboxes[itIdx];
                            if (itEntry && itEntry.cb && !itEntry.cb.disabled) {
                                itEntry.cb.checked = checked;
                                log("ImportIE: setDescendants - set item " + itIdx + " to " + checked);
                            }
                            iti = iti + 1;
                        }
                    }
                }
                sai = sai + 1;
            }
            updateCounter();
        }

        function setSADescendants(saIdx, checked) {
            log("ImportIE: setSADescendants saIdx=" + saIdx + " checked=" + checked);
            var saEntry = saCheckboxes[saIdx];
            if (!saEntry) {
                log("ImportIE: setSADescendants - SA entry not found");
                return;
            }
            var iti = 0;
            while (iti < saEntry.itemIndices.length) {
                var itIdx = saEntry.itemIndices[iti];
                var itEntry = itemCheckboxes[itIdx];
                if (itEntry && itEntry.cb && !itEntry.cb.disabled) {
                    itEntry.cb.checked = checked;
                    log("ImportIE: setSADescendants - set item " + itIdx + " to " + checked);
                }
                iti = iti + 1;
            }
            updateCounter();
        }

        var counterEl = document.createElement("div");
        counterEl.style.padding = "8px 0";
        counterEl.style.fontWeight = "500";
        counterEl.style.fontSize = "13px";
        counterEl.textContent = "Selected: 0";

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#1a7a1a";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "1px solid #2a9a2a";
        confirmBtn.style.padding = "8px 24px";
        confirmBtn.style.borderRadius = "4px";
        confirmBtn.style.cursor = "pointer";
        confirmBtn.style.fontWeight = "600";
        confirmBtn.style.transition = "background 0.15s, opacity 0.15s";
        confirmBtn.disabled = true;
        confirmBtn.style.opacity = "0.5";
        confirmBtn.addEventListener("mouseenter", function () {
            if (!confirmBtn.disabled) {
                confirmBtn.style.background = "#228a22";
            }
        });
        confirmBtn.addEventListener("mouseleave", function () {
            if (!confirmBtn.disabled) {
                confirmBtn.style.background = "#1a7a1a";
            }
        });

        function renderRightList(filter) {
            rightList.innerHTML = "";
            allCheckboxes = [];
            planCheckboxes = [];
            saCheckboxes = [];
            itemCheckboxes = [];
            var filterLower = (filter + "").toLowerCase();

            var hi = 0;
            while (hi < hierarchy.length) {
                var plan = hierarchy[hi];
                var planIdx = planCheckboxes.length;

                var planRow = document.createElement("div");
                planRow.style.display = "flex";
                planRow.style.alignItems = "center";
                planRow.style.gap = "6px";
                planRow.style.padding = "6px 4px";
                planRow.style.fontWeight = "600";
                planRow.style.fontSize = "13px";
                planRow.style.borderBottom = "1px solid #333";
                planRow.style.marginTop = hi > 0 ? "10px" : "0";

                var planCb = document.createElement("input");
                planCb.type = "checkbox";
                planCb.checked = true;
                planCb.style.cursor = "pointer";
                planCb.style.flexShrink = "0";

                var planLabel = document.createElement("span");
                planLabel.textContent = plan.text;
                planLabel.style.cursor = "pointer";

                planRow.appendChild(planCb);
                planRow.appendChild(planLabel);

                var planEntry = { cb: planCb, saIndices: [], row: planRow };
                planCheckboxes.push(planEntry);

                var planVisible = false;

                var sai = 0;
                while (sai < plan.activities.length) {
                    var sa = plan.activities[sai];
                    var saIdx = saCheckboxes.length;
                    planEntry.saIndices.push(saIdx);

                    var saRow = document.createElement("div");
                    saRow.style.display = "flex";
                    saRow.style.alignItems = "center";
                    saRow.style.gap = "6px";
                    saRow.style.padding = "4px 4px 4px 24px";
                    saRow.style.fontSize = "12px";
                    saRow.style.color = "#ccc";

                    var saCb = document.createElement("input");
                    saCb.type = "checkbox";
                    saCb.checked = true;
                    saCb.style.cursor = "pointer";
                    saCb.style.flexShrink = "0";

                    var saLabel = document.createElement("span");
                    saLabel.textContent = "- " + sa.text;
                    saLabel.style.cursor = "pointer";

                    saRow.appendChild(saCb);
                    saRow.appendChild(saLabel);

                    var saEntry = { cb: saCb, itemIndices: [], planIdx: planIdx, row: saRow };
                    saCheckboxes.push(saEntry);

                    var saVisible = false;

                    var iti = 0;
                    while (iti < sa.items.length) {
                        var item = sa.items[iti];
                        var itemIdx = itemCheckboxes.length;
                        saEntry.itemIndices.push(itemIdx);

                        var alreadyExists = existingCodeSet.has(item.code.toUpperCase());
                        var statusText = alreadyExists ? "Already Exist" : "Not Added";

                        var itemRow = document.createElement("div");
                        itemRow.style.display = "flex";
                        itemRow.style.alignItems = "center";
                        itemRow.style.gap = "6px";
                        itemRow.style.padding = "3px 4px 3px 48px";
                        itemRow.style.fontSize = "12px";

                        var itemCb = document.createElement("input");
                        itemCb.type = "checkbox";
                        itemCb.style.cursor = "pointer";
                        itemCb.style.flexShrink = "0";
                        if (alreadyExists) {
                            itemCb.checked = false;
                            itemCb.disabled = true;
                            itemCb.style.cursor = "not-allowed";
                        } else {
                            itemCb.checked = true;
                        }

                        var itemLabel = document.createElement("span");
                        itemLabel.textContent = "-- " + item.checkItemText;
                        itemLabel.style.flex = "1";
                        itemLabel.style.wordBreak = "break-word";

                        var statusBadge = document.createElement("span");
                        statusBadge.style.fontSize = "10px";
                        statusBadge.style.padding = "2px 6px";
                        statusBadge.style.borderRadius = "3px";
                        statusBadge.style.flexShrink = "0";
                        if (alreadyExists) {
                            statusBadge.textContent = "Already Exist";
                            statusBadge.style.background = "#555";
                            statusBadge.style.color = "#aaa";
                        } else {
                            statusBadge.textContent = "Not Added";
                            statusBadge.style.background = "#1a4a1a";
                            statusBadge.style.color = "#6f6";
                        }

                        itemRow.appendChild(itemCb);
                        itemRow.appendChild(itemLabel);
                        itemRow.appendChild(statusBadge);

                        var itemEntry = { cb: itemCb, mapping: item, row: itemRow, planIdx: planIdx, saIdx: saIdx };
                        itemCheckboxes.push(itemEntry);

                        var matchesFilter = true;
                        if (filterLower.length > 0) {
                            var combined = (item.code + " " + item.checkItemText + " " + item.activityPlanText + " " + item.scheduledActivityText).toLowerCase();
                            if (combined.indexOf(filterLower) < 0) {
                                matchesFilter = false;
                            }
                        }
                        if (matchesFilter) {
                            saVisible = true;
                        }
                        itemRow.style.display = matchesFilter ? "flex" : "none";

                        (function (capturedItemIdx, capturedItemEntry) {
                            itemCb.addEventListener("change", function(e) {
                                e.stopPropagation();
                                toggleItem(capturedItemIdx, capturedItemEntry);
                            });
                        })(itemIdx, itemEntry);

                        allCheckboxes.push({ type: "item", idx: itemIdx, row: itemRow });
                        iti = iti + 1;
                    }

                    if (saVisible) {
                        planVisible = true;
                    }
                    saRow.style.display = saVisible ? "flex" : "none";

                    (function (capturedSaIdx, capturedSaEntry) {
                        saCb.addEventListener("change", function(e) {
                            e.stopPropagation();
                            toggleSA(capturedSaIdx, capturedSaEntry);
                        });
                    })(saIdx, saEntry);

                    allCheckboxes.push({ type: "sa", idx: saIdx, row: saRow });
                    sai = sai + 1;
                }

                planRow.style.display = planVisible ? "flex" : "none";

                (function (capturedPlanIdx, capturedPlanEntry) {
                    planCb.addEventListener("change", function(e) {
                        e.stopPropagation();
                        togglePlan(capturedPlanIdx, capturedPlanEntry);
                    });
                })(planIdx, planEntry);

                allCheckboxes.push({ type: "plan", idx: planIdx, row: planRow });

                rightList.appendChild(planRow);

                var sa2i = 0;
                while (sa2i < plan.activities.length) {
                    var saE = saCheckboxes[planEntry.saIndices[sa2i]];
                    rightList.appendChild(saE.row);
                    var it2i = 0;
                    while (it2i < saE.itemIndices.length) {
                        var itE = itemCheckboxes[saE.itemIndices[it2i]];
                        rightList.appendChild(itE.row);
                        it2i = it2i + 1;
                    }
                    sa2i = sa2i + 1;
                }

                hi = hi + 1;
            }

            var pci = 0;
            while (pci < planCheckboxes.length) {
                updateParentState(pci);
                pci = pci + 1;
            }
        }

        renderRightList("");
        rightSearch.addEventListener("input", function () {
            renderRightList(rightSearch.value);
        });
        rightPanel.appendChild(rightList);

        rightPanel.appendChild(counterEl);

        panelBody.appendChild(rightPanel);
        container.appendChild(panelBody);

        var footerBar = document.createElement("div");
        footerBar.style.display = "flex";
        footerBar.style.alignItems = "center";
        footerBar.style.justifyContent = "flex-end";
        footerBar.style.gap = "10px";
        footerBar.style.padding = "12px 16px";
        footerBar.style.borderTop = "1px solid #333";
        footerBar.style.flexShrink = "0";
        footerBar.style.background = "#0d0d0d";

        var selectAllBtn = document.createElement("button");
        selectAllBtn.textContent = "Select All";
        selectAllBtn.style.background = "#2a2a2a";
        selectAllBtn.style.color = "#fff";
        selectAllBtn.style.border = "1px solid #444";
        selectAllBtn.style.padding = "8px 16px";
        selectAllBtn.style.borderRadius = "4px";
        selectAllBtn.style.cursor = "pointer";
        selectAllBtn.style.transition = "background 0.15s";
        selectAllBtn.addEventListener("mouseenter", function () {
            selectAllBtn.style.background = "#3a3a3a";
        });
        selectAllBtn.addEventListener("mouseleave", function () {
            selectAllBtn.style.background = "#2a2a2a";
        });
        selectAllBtn.addEventListener("click", function () {
            log("ImportIE: Select All clicked");
            var xi = 0;
            while (xi < itemCheckboxes.length) {
                if (!itemCheckboxes[xi].cb.disabled) {
                    itemCheckboxes[xi].cb.checked = true;
                }
                xi = xi + 1;
            }
            var pci2 = 0;
            while (pci2 < planCheckboxes.length) {
                updateParentState(pci2);
                pci2 = pci2 + 1;
            }
        });

        var clearAllBtn = document.createElement("button");
        clearAllBtn.textContent = "Clear All";
        clearAllBtn.style.background = "#2a2a2a";
        clearAllBtn.style.color = "#fff";
        clearAllBtn.style.border = "1px solid #444";
        clearAllBtn.style.padding = "8px 16px";
        clearAllBtn.style.borderRadius = "4px";
        clearAllBtn.style.cursor = "pointer";
        clearAllBtn.style.transition = "background 0.15s";
        clearAllBtn.addEventListener("mouseenter", function () {
            clearAllBtn.style.background = "#3a3a3a";
        });
        clearAllBtn.addEventListener("mouseleave", function () {
            clearAllBtn.style.background = "#2a2a2a";
        });
        clearAllBtn.addEventListener("click", function () {
                selectedItems = {};
                selectedSAs = {};
                selectedPlans = {};
                
                var ci = 0;
                while (ci < itemCheckboxes.length) {
                    if (!itemCheckboxes[ci].cb.disabled) {
                        itemCheckboxes[ci].cb.checked = false;
                    }
                    ci = ci + 1;
                }
                var sai = 0;
                while (sai < saCheckboxes.length) {
                    saCheckboxes[sai].cb.checked = false;
                    saCheckboxes[sai].cb.indeterminate = false;
                    sai = sai + 1;
                }
                var pi = 0;
                while (pi < planCheckboxes.length) {
                    planCheckboxes[pi].cb.checked = false;
                    planCheckboxes[pi].cb.indeterminate = false;
                    pi = pi + 1;
                }
                updateCounter();
        });

        confirmBtn.addEventListener("click", function () {
            log("ImportIE: Confirm clicked in review panel");
            var selected = [];
            var gi = 0;
            while (gi < itemCheckboxes.length) {
                if (itemCheckboxes[gi].cb.checked && !itemCheckboxes[gi].cb.disabled) {
                    selected.push(itemCheckboxes[gi].mapping);
                }
                gi = gi + 1;
            }
            log("ImportIE: confirmed " + String(selected.length) + " items");
            overlay.remove();
            onConfirm(selected);
        });

        footerBar.appendChild(selectAllBtn);
        footerBar.appendChild(clearAllBtn);
        footerBar.appendChild(confirmBtn);
        container.appendChild(footerBar);

        overlay.appendChild(container);
        document.body.appendChild(overlay);

        updateCounter();
        log("ImportIE: buildImportIEReviewPanel rendered");
    }

    function buildProgressPopup(selectedMappings) {
        log("ImportIE: buildProgressPopup items=" + String(selectedMappings.length));
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "10px";

        var summaryRow = document.createElement("div");
        summaryRow.style.display = "flex";
        summaryRow.style.justifyContent = "space-between";
        summaryRow.style.fontSize = "13px";
        summaryRow.style.padding = "8px 10px";
        summaryRow.style.borderBottom = "1px solid #333";
        summaryRow.style.background = "#1a1a1a";
        summaryRow.style.borderRadius = "4px";

        var totalEl = document.createElement("span");
        totalEl.textContent = "Total: " + String(selectedMappings.length);
        var successEl = document.createElement("span");
        successEl.textContent = "Success: 0";
        successEl.style.color = "#5cb85c";
        var failEl = document.createElement("span");
        failEl.textContent = "Failed: 0";
        failEl.style.color = "#d9534f";
        var statusEl = document.createElement("span");
        statusEl.textContent = "In Progress";
        statusEl.style.color = "#5bc0de";
        statusEl.style.fontWeight = "600";

        summaryRow.appendChild(totalEl);
        summaryRow.appendChild(successEl);
        summaryRow.appendChild(failEl);
        summaryRow.appendChild(statusEl);
        container.appendChild(summaryRow);

        var listEl = document.createElement("div");
        listEl.style.maxHeight = "400px";
        listEl.style.overflowY = "auto";
        listEl.style.fontSize = "12px";

        var rows = [];
        var ri = 0;
        while (ri < selectedMappings.length) {
            var m = selectedMappings[ri];
            var row = document.createElement("div");
            row.style.display = "flex";
            row.style.justifyContent = "space-between";
            row.style.alignItems = "center";
            row.style.padding = "4px 6px";
            row.style.borderBottom = "1px solid #222";

            var labelSpan = document.createElement("span");
            labelSpan.textContent = m.code + " - " + m.checkItemText;
            labelSpan.style.flex = "1";
            labelSpan.style.overflow = "hidden";
            labelSpan.style.textOverflow = "ellipsis";
            labelSpan.style.whiteSpace = "nowrap";
            labelSpan.style.marginRight = "8px";

            var statusSpan = document.createElement("span");
            statusSpan.textContent = "Pending";
            statusSpan.style.color = "#888";
            statusSpan.style.flexShrink = "0";
            statusSpan.style.fontSize = "11px";
            statusSpan.style.padding = "2px 8px";
            statusSpan.style.borderRadius = "3px";
            statusSpan.style.background = "#222";

            row.appendChild(labelSpan);
            row.appendChild(statusSpan);
            listEl.appendChild(row);
            rows.push({ row: row, statusSpan: statusSpan });
            ri = ri + 1;
        }
        container.appendChild(listEl);

        var popup = createPopup({
            title: "Import I/E - Progress",
            content: container,
            width: "600px",
            height: "auto",
            maxHeight: "80vh",
            onClose: function() {
                log("ImportIE: Progress popup closed by user, cancelling operation");
                IMPORT_IE_CANCELED = true;
            }
        });

        return {
            popup: popup,
            rows: rows,
            successEl: successEl,
            failEl: failEl,
            statusEl: statusEl,
            updateItem: function (index, status, msg) {
                if (index >= 0 && index < rows.length) {
                    rows[index].statusSpan.textContent = status;
                    if (status === "Success") {
                        rows[index].statusSpan.style.color = "#5cb85c";
                    } else if (status === "Failed") {
                        rows[index].statusSpan.style.color = "#d9534f";
                        if (msg) {
                            rows[index].statusSpan.title = msg;
                        }
                    }
                    listEl.scrollTop = listEl.scrollHeight;
                }
            },
            updateSummary: function (successes, failures) {
                successEl.textContent = "Success: " + String(successes);
                failEl.textContent = "Failed: " + String(failures);
            },
            setCompleted: function () {
                statusEl.textContent = "Completed";
                statusEl.style.color = "#5cb85c";
            }
        };
    }

    async function executeSelectedMappings(selectedMappings, existingCodeSet) {
        log("ImportIE: executeSelectedMappings start count=" + String(selectedMappings.length));
        IMPORT_IE_CANCELED = false; 
        var progress = buildProgressPopup(selectedMappings);
        var successes = 0;
        var failures = 0;

        var mi = 0;
        while (mi < selectedMappings.length) {
            var mapping = selectedMappings[mi];
            log("ImportIE: processing item " + String(mi + 1) + "/" + String(selectedMappings.length) + " code=" + String(mapping.code));

            var modalOpen = document.querySelector("#ajaxModal .modal-content");
            if (!modalOpen) {
                log("ImportIE: modal not open, clicking #addEligButton");
                var addBtn = document.querySelector("a#addEligButton");
                if (!addBtn) {
                    addBtn = document.querySelector("#addEligButton");
                }
                if (!addBtn) {
                    log("ImportIE: #addEligButton not found");
                    progress.updateItem(mi, "Failed", "Add button not found");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
                if (addBtn.hasAttribute("disabled")) {
                    log("ImportIE: #addEligButton is disabled");
                    progress.updateItem(mi, "Failed", "Add button disabled");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
                addBtn.click();
                var opened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                if (!opened) {
                    log("ImportIE: modal did not open on first try, retrying");
                    addBtn.click();
                    opened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                }
                if (!opened) {
                    log("ImportIE: modal failed to open after retry");
                    progress.updateItem(mi, "Failed", "Modal did not open");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
                if (IMPORT_IE_CANCELED) {
                    log("ImportIE: Operation cancelled by user");
                    progress.statusEl.textContent = "Cancelled";
                    progress.statusEl.style.color = "#d9534f";
                    return;
                }
            }
            await sleep(500);

            log("ImportIE: step b - selecting eligibility item for code=" + String(mapping.code));
            var eligSel = document.querySelector("select#eligibilityItemRef");
            if (!eligSel) {
                eligSel = await waitForElement("select#eligibilityItemRef", 10000);
            }
            if (!eligSel) {
                log("ImportIE: eligibilityItemRef not found");
                progress.updateItem(mi, "Failed", "Eligibility select not found");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb1 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb1) {
                    cb1.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            var eligMatched = false;
            var eligOpts = eligSel.querySelectorAll("option");
            var bestMatch = null;
            var bestMatchVal = "";
            var eoi = 0;
            while (eoi < eligOpts.length) {
                var eoVal = (eligOpts[eoi].value + "").trim();
                var eoTxt = (eligOpts[eoi].textContent + "").trim();
                if (eoVal.length > 0) {
                    var eoCode = extractIECodeStrict(eoTxt);
                    if (eoCode.length === 0) {
                        eoCode = parseItemCodeFromEligibilityOptionText(eoTxt);
                    }
                    if (eoCode.toUpperCase() === mapping.code.toUpperCase()) {
                        bestMatch = eligOpts[eoi];
                        bestMatchVal = eoVal;
                        break;
                    }
                    if (!bestMatch) {
                        var eoTxtUpper = eoTxt.toUpperCase();
                        var codeUpper = mapping.code.toUpperCase();
                        var re = new RegExp("\\b" + codeUpper.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i");
                        if (re.test(eoTxtUpper)) {
                            bestMatch = eligOpts[eoi];
                            bestMatchVal = eoVal;
                        }
                    }
                }
                eoi = eoi + 1;
            }

            if (!bestMatch) {
                log("ImportIE: no eligibility item match for code=" + String(mapping.code));
                progress.updateItem(mi, "Failed", "No eligibility item match");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb2 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb2) {
                    cb2.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            eligSel.value = bestMatchVal;
            select2TriggerChange(eligSel);
            log("ImportIE: eligibility item selected value='" + String(bestMatchVal) + "'");
            await sleep(importIERandomDelay() + 200);
            eligMatched = true;

            if (IMPORT_IE_CANCELED) {
                log("ImportIE: Operation cancelled by user after sleep");
                progress.statusEl.textContent = "Cancelled";
                progress.statusEl.style.color = "#d9534f";
                return;
            }

            log("ImportIE: step c - setting sex option");
            var eligItemText = (bestMatch.textContent + "").trim();
            var sex = parseSexFromEligibilityText(eligItemText);
            var sexSel = document.querySelector("select#sexOption");
            if (sexSel) {
                sexSel.value = sex;
                select2TriggerChange(sexSel);
                log("ImportIE: sex set to '" + String(sex) + "'");
                await sleep(importIERandomDelay());
            } else {
                log("ImportIE: sexOption select not found, skipping sex step");
            }

            log("ImportIE: step d - selecting activity plan value='" + String(mapping.ids.activityPlanValue) + "'");
            var planOk = await select2SelectByValue("select#activityPlan", mapping.ids.activityPlanValue);
            if (!planOk) {
                planOk = await select2SelectByText("select#activityPlan", mapping.activityPlanText);
            }
            if (!planOk) {
                log("ImportIE: activity plan selection failed");
                progress.updateItem(mi, "Failed", "Activity plan selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb3 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb3) {
                    cb3.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(importIERandomDelay() + 300);

            log("ImportIE: step e - selecting scheduled activity value='" + String(mapping.ids.scheduledActivityValue) + "'");
            var schedSelE = document.querySelector("select#scheduledActivity");
            if (schedSelE) {
                await waitForSelectOptions(schedSelE, 1, 8000);
            }
            var saOk = await select2SelectByValue("select#scheduledActivity", mapping.ids.scheduledActivityValue);
            if (!saOk) {
                saOk = await select2SelectByText("select#scheduledActivity", mapping.scheduledActivityText);
            }
            if (!saOk) {
                log("ImportIE: scheduled activity selection failed");
                progress.updateItem(mi, "Failed", "Scheduled activity selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb4 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb4) {
                    cb4.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(importIERandomDelay() + 300);
            if (IMPORT_IE_CANCELED) {
                log("ImportIE: Operation cancelled by user after sleep");
                progress.statusEl.textContent = "Cancelled";
                progress.statusEl.style.color = "#d9534f";
                return;
            }
            log("ImportIE: step f - selecting check item value='" + String(mapping.ids.checkItemValue) + "'");
            var itemRefSelF = document.querySelector("select#itemRef");
            if (itemRefSelF) {
                await waitForSelectOptions(itemRefSelF, 1, 8000);
            }
            var prevSigF = getItemRefOptionsSignature();
            await waitForItemRefReload(prevSigF, 2000);
            var itemOk = await select2SelectByValue("select#itemRef", mapping.ids.checkItemValue);
            if (!itemOk) {
                itemOk = await select2SelectByText("select#itemRef", mapping.checkItemText);
            }
            if (!itemOk) {
                log("ImportIE: check item selection failed");
                progress.updateItem(mi, "Failed", "Check item selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb5 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb5) {
                    cb5.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(100);

            log("ImportIE: step g - setting comparator to EQ");
            var compOk = await setComparatorEQ();
            if (!compOk) {
                await stabilizeSelectionBeforeComparator(4000);
                compOk = await setComparatorEQ();
            }
            if (!compOk) {
                log("ImportIE: comparator EQ failed");
                progress.updateItem(mi, "Failed", "Comparator EQ failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb6 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb6) {
                    cb6.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            log("ImportIE: step h - selecting code list item containing SF");
            var sfOk = await selectCodeListValueContainingSF();
            if (!sfOk) {
                log("ImportIE: SF selection failed");
                progress.updateItem(mi, "Failed", "Code list SF selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb7 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb7) {
                    cb7.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            if (IMPORT_IE_CANCELED) {
                log("ImportIE: Operation cancelled by user after sleep");
                progress.statusEl.textContent = "Cancelled";
                progress.statusEl.style.color = "#d9534f";
                return;
            }
            log("ImportIE: step i - clicking Save");
            var saveBtn = document.querySelector("button#actionButton");
            if (!saveBtn) {
                saveBtn = await waitForElement("button#actionButton", 5000);
            }
            if (!saveBtn) {
                log("ImportIE: Save button not found");
                progress.updateItem(mi, "Failed", "Save button not found");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                mi = mi + 1;
                continue;
            }
            saveBtn.click();
            log("ImportIE: Save clicked, waiting for modal close");
            var modalClosed = await waitForModalClose(15000);
            if (!modalClosed) {
                log("ImportIE: modal did not close after save, checking for errors");
                var errorEl = document.querySelector("#ajaxModal .modal-content .alert-danger, #ajaxModal .modal-content .error-message, #ajaxModal .modal-content .has-error");
                if (errorEl) {
                    var errorText = (errorEl.textContent + "").trim().replace(/\s+/g, " ");
                    log("ImportIE: validation error detected: " + String(errorText));
                    progress.updateItem(mi, "Failed", errorText);
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    var cb8 = document.querySelector("#ajaxModal .modal-content button.close");
                    if (cb8) {
                        cb8.click();
                        await waitForModalClose(5000);
                    }
                    mi = mi + 1;
                    continue;
                }
                log("ImportIE: retrying save");
                saveBtn = document.querySelector("button#actionButton");
                if (saveBtn) {
                    saveBtn.click();
                    modalClosed = await waitForModalClose(10000);
                }
                if (!modalClosed) {
                    log("ImportIE: save retry failed");
                    progress.updateItem(mi, "Failed", "Modal did not close after save");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    var cb9 = document.querySelector("#ajaxModal .modal-content button.close");
                    if (cb9) {
                        cb9.click();
                        await waitForModalClose(5000);
                    }
                    mi = mi + 1;
                    continue;
                }
            }

            log("ImportIE: item saved successfully code=" + String(mapping.code));
            progress.updateItem(mi, "Success");
            successes = successes + 1;
            progress.updateSummary(successes, failures);
            existingCodeSet.add(mapping.code.toUpperCase());
            await sleep(importIERandomDelay() + 500);

            mi = mi + 1;
        }

        progress.setCompleted();
        log("ImportIE: executeSelectedMappings done successes=" + String(successes) + " failures=" + String(failures));
    }

    function startImportEligibilityMapping() {
        log("ImportIE: startImportEligibilityMapping invoked");

        if (!isValidEligibilityPage()) {
            createPopup({
                title: "Import I/E - Error",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#ff6b6b;font-size:16px;margin-bottom:12px;">⚠️ Wrong Page</p><p>You must navigate to the Eligibility list page before using Import I/E.</p><p style="margin-top:12px;font-size:12px;color:#888;word-wrap:break-word;word-break:break-all;">Required URL: ' + getBaseUrl() + ELIGIBILITY_LIST_PATH + '</p></div>',
                width: "450px",
                height: "auto"
            });
            log("ImportIE: invalid page, stopping");
            return;
        }

        log("ImportIE: page validated, starting collection flow");

        var loadingEl = document.createElement("div");
        loadingEl.style.textAlign = "center";
        loadingEl.style.fontSize = "15px";
        loadingEl.style.color = "#fff";
        loadingEl.style.padding = "20px";
        loadingEl.textContent = "Collecting existing table codes and mappings...";

        var loadingPopup = createPopup({
            title: "Import I/E - Scanning",
            content: loadingEl,
            width: "500px",
            height: "auto"
        });

        var dots = 1;
        var loadingInterval = setInterval(function () {
            dots = dots + 1;
            if (dots > 3) {
                dots = 1;
            }
            var t = "Scanning";
            var di = 0;
            while (di < dots) {
                t = t + ".";
                di = di + 1;
            }
            loadingEl.textContent = t;
        }, 400);

        setTimeout(async function () {
            try {
                log("ImportIE: step 1 - collecting existing codes from table");
                loadingEl.textContent = "Step 1: Collecting existing table codes...";
                var existingCodeSet = await collectAllTableCodes();
                log("ImportIE: existing codes collected count=" + String(existingCodeSet.size));

                log("ImportIE: step 2 - checking Add button");
                loadingEl.textContent = "Step 2: Checking Add button...";
                var addBtn = document.querySelector("a#addEligButton");
                if (!addBtn) {
                    addBtn = document.querySelector("#addEligButton");
                }
                if (!addBtn) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Error", "The Add button (#addEligButton) was not found on this page.");
                    log("ImportIE: add button not found, stopping");
                    return;
                }
                if (addBtn.hasAttribute("disabled")) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Add Button Disabled", "The Add button is currently disabled. Please ensure you have permission to add eligibility items and that the mapping is unlocked.");
                    log("ImportIE: add button disabled, stopping");
                    return;
                }

                log("ImportIE: step 3 - clicking Add to open modal");
                loadingEl.textContent = "Step 3: Opening modal...";
                addBtn.click();
                var modalOpened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                if (!modalOpened) {
                    log("ImportIE: modal did not open on first try, retrying");
                    addBtn.click();
                    modalOpened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                }
                if (!modalOpened) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Error", "The Eligibility Management modal did not open. Please try again.");
                    log("ImportIE: modal failed to open, stopping");
                    return;
                }
                await sleep(800);

                log("ImportIE: step 4 - collecting mappings from modal");
                loadingEl.textContent = "Step 4: Scanning all Activity Plans, Scheduled Activities, and Check Items...";
                var rawMappings = await collectMappingsFromModal(existingCodeSet);
                log("ImportIE: raw mappings collected count=" + String(rawMappings.length));

                var mappings = deduplicateMappings(rawMappings);
                log("ImportIE: deduplicated mappings count=" + String(mappings.length));

                log("ImportIE: closing modal after collection");
                var closeModalBtn = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeModalBtn) {
                    closeModalBtn.click();
                    await waitForModalClose(5000);
                }

                clearInterval(loadingInterval);
                loadingPopup.close();

                if (mappings.length === 0) {
                    showWarningPopup("Import I/E - No Mappings", "No INC/EXC check items were found across any Activity Plans and Scheduled Activities.");
                    log("ImportIE: no mappings found, stopping");
                    return;
                }

                log("ImportIE: step 5 - showing review panel");
                buildImportIEReviewPanel(existingCodeSet, mappings, function (selectedMappings) {
                    log("ImportIE: user confirmed " + String(selectedMappings.length) + " mappings");
                    if (selectedMappings.length === 0) {
                        log("ImportIE: no items selected, stopping");
                        return;
                    }
                    executeSelectedMappings(selectedMappings, existingCodeSet);
                });

            } catch (err) {
                clearInterval(loadingInterval);
                if (loadingPopup && loadingPopup.close) {
                    loadingPopup.close();
                }
                log("ImportIE: error in startImportEligibilityMapping: " + String(err));
                showWarningPopup("Import I/E - Error", "An unexpected error occurred: " + String(err));
            }
        }, 300);
    }

    // ======================
    // ======================
    function setPanelHidden(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_HIDDEN, flag ? "1" : "0");
            log("Panel hidden state set to " + String(flag));
        } catch (e) {
        }
    }

    function getPanelHidden() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_PANEL_HIDDEN);
        } catch (e) {
        }
        if (raw === "1") {
            return true;
        }
        return false;
    }

    function showAeSubjectInputPopup(onDone) {
        log("Find AE: showing subject input popup");
        var container = document.createElement("div");
        container.style.display = "grid";
        container.style.gridTemplateRows = "auto auto";
        container.style.gap = "10px";

        var fieldRow = document.createElement("div");
        fieldRow.style.display = "grid";
        fieldRow.style.gridTemplateColumns = "120px 1fr";
        fieldRow.style.alignItems = "center";
        fieldRow.style.gap = "8px";
        var label = document.createElement("div");
        label.textContent = AE_POPUP_LABEL;
        label.style.fontWeight = "600";
        var input = document.createElement("input");
        input.type = "text";
        input.placeholder = "e.g., 1001, S123, or subject label text";
        input.style.width = "100%";
        input.style.boxSizing = "border-box";
        input.style.padding = "8px";
        input.style.borderRadius = "6px";
        input.style.border = "1px solid #444";
        input.style.background = "#1a1a1a";
        input.style.color = "#fff";
        fieldRow.appendChild(label);
        fieldRow.appendChild(input);
        container.appendChild(fieldRow);

        var btnRow = document.createElement("div");
        btnRow.style.display = "inline-flex";
        btnRow.style.justifyContent = "flex-end";
        btnRow.style.gap = "8px";
        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = AE_POPUP_CANCEL_TEXT;
        cancelBtn.style.background = "#333";
        cancelBtn.style.color = "#fff";
        cancelBtn.style.border = "none";
        cancelBtn.style.borderRadius = "6px";
        cancelBtn.style.padding = "8px 12px";
        cancelBtn.style.cursor = "pointer";
        var okBtn = document.createElement("button");
        okBtn.textContent = AE_POPUP_OK_TEXT;
        okBtn.style.background = "#0b82ff";
        okBtn.style.color = "#fff";
        okBtn.style.border = "none";
        okBtn.style.borderRadius = "6px";
        okBtn.style.padding = "8px 12px";
        okBtn.style.cursor = "pointer";
        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(okBtn);
        container.appendChild(btnRow);

        var popup = createPopup({ title: AE_POPUP_TITLE, content: container, width: "380px", height: "auto" });

        window.setTimeout(function () {
            try {
                input.focus();
                input.select();
                log("Find AE: subject input focused");
            } catch (e) {
                log("Find AE: failed to focus subject input");
            }
        }, 50);

        function doContinue() {
            var v = input.value + "";
            var t = v.trim();
            log("Find AE: popup input='" + String(t) + "'");
            if (popup && popup.close) {
                popup.close();
            }
            document.removeEventListener("keydown", keyHandler, true);
            if (typeof onDone === "function") {
                onDone(t);
            }
        }

        function keyHandler(e) {
            var code = e.key || e.code || "";
            if (code === "Enter") {
                log("Find AE: Enter pressed; continuing");
                doContinue();
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            if (code === "Escape" || code === "Esc") {
                log("Find AE: Esc pressed; closing");
                if (popup && popup.close) {
                    popup.close();
                }
                document.removeEventListener("keydown", keyHandler, true);
                if (typeof onDone === "function") {
                    onDone(null);
                }
                e.preventDefault();
                e.stopPropagation();
                return;
            }
        }

        document.addEventListener("keydown", keyHandler, true);

        cancelBtn.addEventListener("click", function () {
            log("Find AE: popup canceled by user");
            if (popup && popup.close) {
                popup.close();
            }
            document.removeEventListener("keydown", keyHandler, true);
            if (typeof onDone === "function") {
                onDone(null);
            }
        });

        okBtn.addEventListener("click", function () {
            doContinue();
        });
    }

    function continueFindAdverseEvent(b, subjId) {
        var urls = buildAeListUrl();
        var u = urls.baseUrl;
        var h = urls.searchUrl;
        log("Find AE: opening list search url='" + String(h) + "'");
        var w = window.open(h, "_blank");
        var t = 0;
        var r = setInterval(function () {
            t = t + 1;
            if (!w || w.closed) {
                clearInterval(r);
                log("Find AE: child window closed");
                return;
            }
            if (t > 120) {
                clearInterval(r);
                log("Find AE: timeout waiting for child page");
                return;
            }
            var d = w.document;
            if (!d) {
                log("Find AE: child document not ready at t=" + String(t));
                return;
            }

            var os = d.querySelectorAll('#subjectIds option,select[name="subjectIds"] option');
            var y = "";
            if (os && os.length > 0) {
                var j = 0;
                while (j < os.length) {
                    var op = os[j];
                    var txt = "";
                    var val = "";
                    if (op) {
                        txt = op.textContent + "";
                        val = op.value + "";
                    }
                    var nt = aeNormalize(txt);
                    var matchByTextExact = b && b.length > 0 && nt === b;
                    var matchByTextContains = b && b.length > 0 && nt.indexOf(b) >= 0;
                    var matchById = subjId && subjId.length > 0 && val === subjId;
                    var matched = false;
                    if (matchByTextExact) {
                        matched = true;
                    } else {
                        if (matchByTextContains) {
                            matched = true;
                        } else {
                            if (matchById) {
                                matched = true;
                            }
                        }
                    }
                    if (matched) {
                        y = val;
                        log("Find AE: subject matched optionIndex=" + String(j) + " text='" + String(txt) + "' value='" + String(val) + "'");
                        break;
                    }
                    j = j + 1;
                }
            } else {
                log("Find AE: subject options not found");
            }

            var oe = d.querySelectorAll('#studyEventIds option,select[name="studyEventIds"] option');
            var z = "";
            var keywordsNorm = [];
            var ki = 0;
            while (ki < AE_EVENT_KEYWORDS.length) {
                keywordsNorm.push(aeNormalize(AE_EVENT_KEYWORDS[ki]));
                ki = ki + 1;
            }
            if (oe && oe.length > 0) {
                var k = 0;
                while (k < oe.length) {
                    var eop = oe[k];
                    var etxt = "";
                    var eval0 = "";
                    if (eop) {
                        etxt = eop.textContent + "";
                        eval0 = eop.value + "";
                    }
                    var ent = aeNormalize(etxt);
                    var matchedEvent = false;
                    var kk = 0;
                    while (kk < keywordsNorm.length) {
                        var kw = keywordsNorm[kk];
                        if (ent.indexOf(kw) >= 0) {
                            matchedEvent = true;
                            break;
                        }
                        kk = kk + 1;
                    }
                    if (matchedEvent) {
                        z = eval0;
                        log("Find AE: event matched optionIndex=" + String(k) + " text='" + String(etxt) + "' value='" + String(eval0) + "'");
                        break;
                    }
                    k = k + 1;
                }
            } else {
                log("Find AE: study event options not found");
            }

            var shouldNavigate = false;
            var finalUrl = u + "&statusValues=Nonconformant&statusValues=notCanceled&statusValues=formDataNotCanceled";

            if (y && y.length > 0) {
                finalUrl = finalUrl + "&subjectIds=" + encodeURIComponent(y);
                shouldNavigate = true;
                log("Find AE: subject will be applied to final url");
            } else {
                log("Find AE: subject not matched; proceeding without subject");
            }

            if (z && z.length > 0) {
                finalUrl = finalUrl + "&studyEventIds=" + encodeURIComponent(z);
                shouldNavigate = true;
                log("Find AE: AE event will be applied to final url");
            } else {
                log("Find AE: AE event not matched; proceeding without event");
            }

            if (shouldNavigate) {
                clearInterval(r);
                log("Find AE: navigating to final url");
                w.location.href = finalUrl;
                return;
            }

            log("Find AE: waiting for dropdowns t=" + String(t));
        }, 200);
    }

    function buildAeListUrl() {
        var g = window.location.href + "";
        var isTest = g.indexOf("cenexeltest") >= 0;
        var base = isTest ? AE_LIST_TEST_BASE_URL : AE_LIST_BASE_URL;
        var h = base + "&statusValues=Nonconformant&statusValues=notCanceled&statusValues=formDataNotCanceled";
        return { baseUrl: base, searchUrl: h };
    }

    function openAndLocateAdverseEvent() {
        log("Find AE: starting");
        var info = getSubjectIdentifierForAE();
        var b = info.normalizedIdentifier;
        var subjId = info.subjectId;
        var confident = info.confident;
        var raw = info.raw;

        var needPopup = false;
        if (!subjId || subjId.length === 0) {
            needPopup = true;
        }

        if (needPopup) {
            log("Find AE: no subjectId found; showing popup");
            showAeSubjectInputPopup(function (userInput) {
                if (userInput === null) {
                    log("Find AE: canceled by user; stopping with no navigation");
                    return;
                }
                var entered = userInput || "";
                var normalized = aeNormalize(entered);
                if (!normalized || normalized.length === 0) {
                    if (b && b.length > 0) {
                        normalized = b;
                        log("Find AE: using page-derived normalized fallback='" + String(normalized) + "'");
                    } else {
                        log("Find AE: proceeding without subject filter");
                    }
                } else {
                    log("Find AE: using user-entered normalized='" + String(normalized) + "'");
                }
                continueFindAdverseEvent(normalized, "");
            });

            var prefill = "";
            if (raw && raw.length > 0) {
                prefill = raw;
            } else {
                prefill = "";
            }
            window.setTimeout(function () {
                var lastPopup = document.querySelector('[id^="clinsparkPopup_"] input[type="text"]');
                if (lastPopup) {
                    lastPopup.value = prefill;
                    var ev = new Event("input", { bubbles: true });
                    lastPopup.dispatchEvent(ev);
                    log("Find AE: popup prefilled='" + String(prefill) + "'");
                }
            }, 120);
            return;
        }

        log("Find AE: proceeding with page-derived subjectId");
        continueFindAdverseEvent(b, subjId);
    }

    //==========================
    // RUN BARCODE FEATURE
    //==========================
    // This section contains all functions related to barcode lookup and data entry.
    // This feature automates finding subject barcodes by searching through epochs,
    // cohorts, and subjects. Functions handle subject identification, barcode retrieval,
    // and populating barcode input modals.
    //==========================


    // Fetch barcode completely in the background without opening a new tab.
    // Uses the subject ID directly if available, or uses a hidden iframe to search.
    async function fetchBarcodeInBackground(subjectText, subjectId) {
        log("Background Barcode: starting search subjectText='" + String(subjectText) + "' subjectId='" + String(subjectId) + "'");

        // If we have a subject ID, the barcode is simply "S" + subjectId
        if (subjectId && subjectId.length > 0) {
            var result = "S" + String(subjectId);
            log("Background Barcode: using direct ID, result=" + result);
            return result;
        }

        // If we only have subject text, we need to search through the barcode printing page
        if (!subjectText || subjectText.length === 0) {
            log("Background Barcode: no subjectText or subjectId provided");
            return null;
        }

        // Use a hidden iframe to load the page and interact with it
        // This allows JavaScript to execute and populate the dynamic dropdowns
        return await searchBarcodeViaIframe(subjectText);
    }

    // Search for barcode using a hidden iframe that can execute JavaScript
    async function searchBarcodeViaIframe(subjectText) {
        log("Background Barcode: using hidden iframe approach");

        return new Promise(function(resolve) {
            var iframe = document.createElement("iframe");
            iframe.style.position = "fixed";
            iframe.style.top = "-9999px";
            iframe.style.left = "-9999px";
            iframe.style.width = "1px";
            iframe.style.height = "1px";
            iframe.style.visibility = "hidden";
            iframe.style.pointerEvents = "none";

            var timeoutId = null;
            var resolved = false;

            function cleanup() {
                if (timeoutId) {
                    clearTimeout(timeoutId);
                    timeoutId = null;
                }
                if (iframe && iframe.parentNode) {
                    iframe.parentNode.removeChild(iframe);
                }
            }

            function finishWithResult(result) {
                if (resolved) return;
                resolved = true;
                cleanup();
                resolve(result);
            }

            // Set a maximum timeout
            timeoutId = setTimeout(function() {
                log("Background Barcode: iframe timeout reached");
                finishWithResult(null);
            }, 30000);

            iframe.onload = async function() {
                try {
                    var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;

                    // Wait for epoch select to appear and be populated (Edge needs more time than Chrome)
                    var epochSel = null;
                    var retryCount = 0;
                    var maxRetries = 80;
                    while (retryCount < maxRetries && !epochSel) {
                        await sleep(300);
                        epochSel = iframeDoc.querySelector("select#epoch");
                        if (epochSel) {
                            var opts = epochSel.querySelectorAll("option");
                            if (opts.length === 0) {
                                epochSel = null;
                            } else {
                                log("Background Barcode: iframe - epoch select ready with " + opts.length + " options (attempt " + (retryCount + 1) + ")");
                            }
                        } else {
                        }
                        retryCount++;
                    }

                    if (!epochSel) {
                        finishWithResult(null);
                        return;
                    }

                    var epochOpts = epochSel.querySelectorAll("option");
                    log("Background Barcode: iframe - found " + String(epochOpts.length) + " epoch options");

                    // Iterate through epochs
                    for (var eIdx = 0; eIdx < epochOpts.length; eIdx++) {
                        if (resolved) return;

                        var epochVal = (epochOpts[eIdx].value + "").trim();
                        if (epochVal.length === 0) continue;

                        log("Background Barcode: iframe - selecting epoch value='" + epochVal + "'");
                        epochSel.value = epochVal;
                        var epochEvt = new Event("change", { bubbles: true });
                        epochSel.dispatchEvent(epochEvt);

                        // Wait for cohorts to load (Edge needs more time)
                        await sleep(1200);

                        var cohortSel = iframeDoc.querySelector("select#cohort");
                        if (!cohortSel) {
                            log("Background Barcode: iframe - cohort select not found");
                            continue;
                        }

                        var cohortOpts = cohortSel.querySelectorAll("option");
                        log("Background Barcode: iframe - found " + String(cohortOpts.length) + " cohort options");

                        for (var cIdx = 0; cIdx < cohortOpts.length; cIdx++) {
                            if (resolved) return;

                            var cohortVal = (cohortOpts[cIdx].value + "").trim();
                            if (cohortVal.length === 0) continue;

                            log("Background Barcode: iframe - selecting cohort value='" + cohortVal + "'");
                            cohortSel.value = cohortVal;
                            var cohortEvt = new Event("change", { bubbles: true });
                            cohortSel.dispatchEvent(cohortEvt);

                            // Wait for subjects to load (Edge needs more time)
                            await sleep(1200);

                            var subjectsSel = iframeDoc.querySelector("select#subjects");
                            if (!subjectsSel) {
                                log("Background Barcode: iframe - subjects select not found");
                                continue;
                            }

                            var subjectOpts = subjectsSel.querySelectorAll("option");
                            log("Background Barcode: iframe - found " + String(subjectOpts.length) + " subject options");

                            for (var sIdx = 0; sIdx < subjectOpts.length; sIdx++) {
                                var sVal = (subjectOpts[sIdx].value + "").trim();
                                var sTxt = (subjectOpts[sIdx].textContent + "").trim();

                                if (sVal.length > 0) {
                                    var textMatch = normalizeSubjectString(sTxt) === normalizeSubjectString(subjectText);
                                    if (textMatch) {
                                        var result = "S" + String(sVal);
                                        log("Background Barcode: iframe - found match! text='" + sTxt + "' value='" + sVal + "' result=" + result);
                                        finishWithResult(result);
                                        return;
                                    }
                                }
                            }
                        }
                    }

                    log("Background Barcode: iframe - subject not found after searching all epochs/cohorts");
                    finishWithResult(null);

                } catch (e) {
                    log("Background Barcode: iframe error - " + String(e));
                    finishWithResult(null);
                }
            };

            iframe.onerror = function() {
                log("Background Barcode: iframe failed to load");
                finishWithResult(null);
            };

            document.body.appendChild(iframe);
            iframe.src = location.origin + "/secure/barcodeprinting/subjects";
        });
    }

    // Run Barcode feature once from the current page context.
    // Runs completely in the background without opening a new tab.
    async function APS_RunBarcode() {
        BARCODE_START_TS = Date.now();

        var info = getSubjectFromBreadcrumbOrTooltip();
        var hasText = !!(info.subjectText && info.subjectText.length > 0);
        var hasId = !!(info.subjectId && info.subjectId.length > 0);

        if (!hasText && !hasId) {
            log("Run Barcode: subject breadcrumb or tooltip not found (APS_RunBarcode)");
            return;
        }

        log("APS_RunBarcode: Fetching barcode in background…");

        var loadingText = document.createElement("div");
        loadingText.style.textAlign = "center";
        loadingText.style.fontSize = "16px";
        loadingText.style.color = "#fff";
        loadingText.style.padding = "20px";
        loadingText.textContent = "Locating barcode.";

        var popup = createPopup({
            title: "Locating Barcode",
            content: loadingText,
            width: "300px",
            height: "auto"
        });

        var dots = 1;
        var loadingInterval = setInterval(function () {
            dots = dots + 1;
            if (dots > 3) {
                dots = 1;
            }
            var text = "Locating barcode";
            var i = 0;
            while (i < dots) {
                text = text + ".";
                i = i + 1;
            }
            loadingText.textContent = text;
        }, 500);

        // Fetch barcode in background using the new function
        var r = await fetchBarcodeInBackground(info.subjectText, info.subjectId);

        try {
            clearInterval(loadingInterval);
        } catch (e1) {}
        try {
            if (popup && popup.close) {
                popup.close();
            }
        } catch (e2) {}

        if (r && r.length > 0) {
            log("APS_RunBarcode: got result '" + r + "'; attempting to populate input");

            var inputBox = document.querySelector("input.bootbox-input.bootbox-input-text.form-control");
            if (!inputBox) {
                inputBox = await openBarcodeDataEntryModalIfNeeded(6000);
            }

            if (inputBox) {
                inputBox.value = r;
                var evt = new Event("input", { bubbles: true });
                inputBox.dispatchEvent(evt);
                log("APS_RunBarcode: populated barcode value in modal");

                var okBtn = document.querySelector("button[data-bb-handler=\"confirm\"].btn.btn-primary");
                if (!okBtn) {
                    okBtn = document.querySelector("button.btn.btn-primary[data-bb-handler=\"confirm\"]");
                }
                if (okBtn) {
                    okBtn.click();
                    log("APS_RunBarcode: confirmed barcode modal");
                } else {
                    log("APS_RunBarcode: OK button not found after populating input");
                }
            } else {
                log("APS_RunBarcode: unable to find or open barcode input modal");
            }
        } else {
            log("APS_RunBarcode: no barcode result found");
        }

        try {
            var secs = (Date.now() - BARCODE_START_TS) / 1000;
            var s = secs.toFixed(2);
            log("APS_RunBarcode: elapsed " + String(s) + " s");
        } catch (e3) {}
        BARCODE_START_TS = 0;
    }


    async function openBarcodeDataEntryModalIfNeeded(timeoutMs) {
        var inputBox = document.querySelector("input.bootbox-input.bootbox-input-text.form-control");
        if (inputBox) {
            return inputBox;
        }
        var link = document.querySelector("a[href='#'][onclick*='inputBarcodeInModal']");
        if (!link) {
            log("Run Barcode: 'Barcode Data Entry' link not found");
            return null;
        }
        link.click();
        log("Run Barcode: 'Barcode Data Entry' link clicked to open modal");
        var inputEl = await waitForSelector("input.bootbox-input.bootbox-input-text.form-control", typeof timeoutMs === "number" ? timeoutMs : 6000);
        if (!inputEl) {
            log("Run Barcode: bootbox input did not appear after clicking 'Barcode Data Entry'");
            return null;
        }
        return inputEl;
    }

    function getText(el) {
        if (!el) {
            return "";
        }
        var s = (el.textContent + "").replace(/\u00A0/g, " ");
        s = s.trim();
        return s;
    }

    function setPanelCollapsed(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_COLLAPSED, flag ? "1" : "0");
        } catch (e) {}
    }

    function getPanelCollapsed() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_PANEL_COLLAPSED);
        } catch (e) {}
        if (raw === "1") {
            return true;
        }
        return false;
    }

    function setBarcodeSubjectText(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_SUBJECT_TEXT, String(t));
        } catch (e) {}
    }

    function getBarcodeSubjectText() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_SUBJECT_TEXT);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    function clearBarcodeSubjectText() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_SUBJECT_TEXT);
        } catch (e) {}
    }

    function setBarcodeSubjectId(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_SUBJECT_ID, String(t));
        } catch (e) {}
    }

    function getBarcodeSubjectId() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_SUBJECT_ID);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    function clearBarcodeSubjectId() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_SUBJECT_ID);
        } catch (e) {}
    }

    function setBarcodeResult(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_RESULT, String(t));
        } catch (e) {}
    }

    function getBarcodeResult() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_RESULT);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    function clearBarcodeResult() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_RESULT);
        } catch (e) {}
    }

    function getSubjectFromBreadcrumbOrTooltip() {
        var text = "";
        var id = "";
        var items = document.querySelectorAll("ul.page-breadcrumb.breadcrumb > li");
        if (items && items.length >= 3) {
            var li = items[2];
            if (li) {
                var t = (li.textContent + "").trim();
                if (t.length > 0) {
                    text = t;
                }
            }
        }
        if (!text || text.length === 0) {
            var a = document.querySelector('span.tooltips[data-original-title="Subject"] a[data-target="#ajaxModal"][data-toggle="modal"]');
            if (a) {
                var href = a.getAttribute("href") + "";
                var m = href.match(/\/show\/subject\/(\d+)\//);
                if (m && m[1]) {
                    id = m[1];
                }
            }
        }
        if ((!text || text.length === 0) && (!id || id.length === 0)) {
            var cap = document.querySelector("div.portlet-title div.caption");
            if (cap) {
                var c = (cap.textContent + "").trim();
                if (c.length > 0) {
                    text = c;
                }
            }
        }
        return { subjectText: text, subjectId: id };
    }

    function normalizeSubjectString(t) {
        if (typeof t !== "string") {
            return "";
        }
        var s = t.replace(/\u00A0/g, " ");
        s = s.trim();
        s = s.replace(/\s+/g, " ");
        return s;
    }

    function isBarcodeSubjectsPage() {
        var path = location.pathname;
        if (path === "/secure/barcodeprinting/subjects") {
            return true;
        }
        return false;
    }

    function firstNonEmptyOptionIndex(selectEl) {
        if (!selectEl) {
            return -1;
        }
        var opts = selectEl.querySelectorAll("option");
        var i = 0;
        while (i < opts.length) {
            var v = (opts[i].value + "").trim();
            if (v.length > 0) {
                return i;
            }
            i = i + 1;
        }
        return -1;
    }

    async function processBarcodeSubjectsPage() {
        var wantedText = getBarcodeSubjectText();
        var wantedId = getBarcodeSubjectId();
        if ((!wantedText || wantedText.length === 0) && (!wantedId || wantedId.length === 0)) {
            log("Run Barcode: no wantedText or wantedId present");
            return;
        }
        log("Run Barcode: starting search wantedText='" + String(wantedText) + "' wantedId='" + String(wantedId) + "'");
        var epochSel = await waitForSelector("select#epoch", 10000);
        if (!epochSel) {
            log("Run Barcode: epoch select not found");
            return;
        }
        var cohortSel = await waitForSelector("select#cohort", 10000);
        if (!cohortSel) {
            log("Run Barcode: cohort select not found");
            return;
        }
        var eOpts = epochSel.querySelectorAll("option");
        log("Run Barcode: epoch options count=" + String(eOpts.length));
        var eStart = firstNonEmptyOptionIndex(epochSel);
        if (eStart < 0) {
            log("Run Barcode: no epoch option with non-empty value");
            return;
        }
        var eIdx = eStart;
        var found = false;
        var foundVal = "";
        var foundText = "";
        while (eIdx < eOpts.length) {
            var eVal = (eOpts[eIdx].value + "").trim();
            log("Run Barcode: selecting epoch index=" + String(eIdx) + " value='" + String(eVal) + "'");
            if (eVal.length > 0) {
                epochSel.value = eVal;
                var eEvt = new Event("change", { bubbles: true });
                epochSel.dispatchEvent(eEvt);
                await sleep(600);
            }
            var cOpts = cohortSel.querySelectorAll("option");
            log("Run Barcode: cohort options count after epoch change=" + String(cOpts.length));
            var cStart = firstNonEmptyOptionIndex(cohortSel);
            if (cStart < 0) {
                log("Run Barcode: no cohort option with non-empty value under epoch value='" + String(eVal) + "'");
                eIdx = eIdx + 1;
                continue;
            }
            var cIdx = cStart;
            while (cIdx < cOpts.length) {
                var cVal = (cOpts[cIdx].value + "").trim();
                log("Run Barcode: selecting cohort index=" + String(cIdx) + " value='" + String(cVal) + "'");
                if (cVal.length > 0) {
                    cohortSel.value = cVal;
                    var cEvt = new Event("change", { bubbles: true });
                    cohortSel.dispatchEvent(cEvt);
                    await sleep(1000);
                }
                var subjectsSel = document.querySelector("#subjectsDiv select#subjects");
                if (!subjectsSel) {
                    log("Run Barcode: subjects select not present immediately, entering retry window");
                }
                var retryElapsed = 0;
                var retryStep = 500;
                var retryMax = 3000;
                var lastCount = 0;
                while (retryElapsed <= retryMax) {
                    subjectsSel = document.querySelector("#subjectsDiv select#subjects");
                    if (subjectsSel) {
                        lastCount = subjectsSel.querySelectorAll("option").length;
                        log("Run Barcode: subjects options count=" + String(lastCount) + " at t=" + String(retryElapsed) + "ms");
                        if (lastCount > 0) {
                            break;
                        }
                    } else {
                        log("Run Barcode: subjects select not found at t=" + String(retryElapsed) + "ms");
                    }
                    await sleep(retryStep);
                    retryElapsed = retryElapsed + retryStep;
                }
                var optsCount = 0;
                if (subjectsSel) {
                    optsCount = subjectsSel.querySelectorAll("option").length;
                }
                log("Run Barcode: final subjects options count=" + String(optsCount));
                var matchOp = null;
                if (subjectsSel && optsCount > 0) {
                    var opts = subjectsSel.querySelectorAll("option");
                    var i = 0;
                    while (i < opts.length) {
                        var op = opts[i];
                        var txt = (op.textContent + "").trim();
                        var val = (op.value + "").trim();
                        var byText = wantedText && wantedText.length > 0 && normalizeSubjectString(txt) === normalizeSubjectString(wantedText);
                        var byId = (!wantedText || wantedText.length === 0) && wantedId && wantedId.length > 0 && val === wantedId;
                        log("Run Barcode: checking option index=" + String(i) + " text='" + String(txt) + "' value='" + String(val) + "' byText=" + String(byText) + " byId=" + String(byId));
                        if (byText || byId) {
                            matchOp = op;
                            break;
                        }
                        i = i + 1;
                    }
                } else {
                    log("Run Barcode: subjects list empty under cohort value='" + String(cVal) + "'");
                }
                if (matchOp) {
                    found = true;
                    foundVal = (matchOp.value + "").trim();
                    foundText = (matchOp.textContent + "").trim();
                    log("Run Barcode: match found text='" + String(foundText) + "' value='" + String(foundVal) + "'");
                    break;
                }
                log("Run Barcode: no match in cohort value='" + String(cVal) + "', trying next cohort");
                cIdx = cIdx + 1;
            }
            if (found) {
                break;
            }
            log("Run Barcode: no match in epoch value='" + String(eVal) + "', trying next epoch");
            eIdx = eIdx + 1;
        }
        if (found) {
            var result = "S" + String(foundVal);
            setBarcodeResult(result);
            log("Run Barcode: result " + result + " from option text='" + String(foundText) + "'");
            var bootInput = document.querySelector("input.bootbox-input.bootbox-input-text.form-control");
            if (bootInput) {
                bootInput.value = result;
                var evtInput = new Event("input", { bubbles: true });
                bootInput.dispatchEvent(evtInput);
                log("Run Barcode: bootbox input autopopulated in child tab");
                var okBtn = document.querySelector('button[data-bb-handler="confirm"].btn.btn-primary');
                if (!okBtn) {
                    okBtn = document.querySelector('button.btn.btn-primary[data-bb-handler="confirm"]');
                }
                if (okBtn) {
                    okBtn.click();
                    log("Run Barcode: bootbox OK clicked in child tab");
                } else {
                    log("Run Barcode: bootbox OK button not found in child tab");
                }
            }
            clearBarcodeSubjectText();
            clearBarcodeSubjectId();
            await sleep(300);
            window.close();
            return;
        }
        setBarcodeResult("Subject not found");
        log("Run Barcode: Subject not found after scanning all epochs and cohorts");
        clearBarcodeSubjectText();
        clearBarcodeSubjectId();
    }

    function isPaused() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_PAUSED);
        } catch (e) {}
        if (raw === "1") {
            return true;
        }
        return false;
    }

    function setPaused(flag) {
        try {
            localStorage.setItem(STORAGE_PAUSED, flag ? "1" : "0");
            log("Paused=" + String(flag));
        } catch (e) {}
    }


    function log(msg) {
        try {
            var ts = new Date().toISOString();
            var line = "[" + ts + "] " + String(msg);
            console.log("[APS] " + line);
        } catch (e) {}
        try {
            var key = "activityPlanState.logs";
            var logs = [];
            var raw = localStorage.getItem(key);
            if (raw) {
                try {
                    logs = JSON.parse(raw);
                } catch (e2) {
                    logs = [];
                }
            }
            logs.push(String(line));
            var limit = 500;
            if (logs.length > limit) {
                var start = logs.length - limit;
                logs = logs.slice(start);
            }
            localStorage.setItem(key, JSON.stringify(logs));
            var box = document.getElementById(LOG_ID);
            if (box) {
                var div = document.createElement("div");
                div.textContent = String(line);
                box.appendChild(div);
                var maxChildren = 500;
                if (box.childNodes && box.childNodes.length > maxChildren) {
                    var extra = box.childNodes.length - maxChildren;
                    var x = 0;
                    while (x < extra) {
                        box.removeChild(box.firstChild);
                        x = x + 1;
                    }
                }
                box.scrollTop = box.scrollHeight;
            }
        } catch (e3) {}
    }


    function openInTab(url, active) {
        var api = typeof GM !== "undefined" && typeof GM.openInTab === "function";
        if (api) {
            return GM.openInTab(url, { active: !!active, insert: true, setParent: true });
        }
        var legacy = typeof GM_openInTab === "function";
        if (legacy) {
            return GM_openInTab(url, { active: !!active, insert: true, setParent: true });
        }
        window.open(url, "_blank");
        return null;
    }

    function sleep(ms) {
        return new Promise(function (resolve) {
            setTimeout(function () {
                resolve();
            }, ms);
        });
    }

    async function waitForSelector(selector, timeoutMs) {
        var start = Date.now();
        var found = null;
        while (Date.now() - start < timeoutMs) {
            found = document.querySelector(selector);
            if (found) {
                return found;
            }
            await sleep(100);
        }
        return null;
    }

    async function waitUntilHidden(selector, timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var el = document.querySelector(selector);
            if (!el) {
                return true;
            }
            var style = window.getComputedStyle(el);
            var hidden = style.display === "none" || style.visibility === "hidden" || style.opacity === "0";
            if (hidden) {
                return true;
            }
            await sleep(150);
        }
        return false;
    }

    function getStoredPos(key, fallback) {
        var raw = null;
        try {
            raw = localStorage.getItem(key);
        } catch (e) {}
        if (!raw) {
            return fallback;
        }
        return raw;
    }

    function setStoredPanelSize(w, h) {
        try {
            if (typeof w === "string") {
                localStorage.setItem(STORAGE_PANEL_WIDTH, w);
            }
            if (typeof h === "string") {
                localStorage.setItem(STORAGE_PANEL_HEIGHT, h);
            }
        } catch (e) {}
    }

    function getStoredPanelSize() {
        var w = null;
        var h = null;
        try {
            w = localStorage.getItem(STORAGE_PANEL_WIDTH);
            h = localStorage.getItem(STORAGE_PANEL_HEIGHT);
        } catch (e) {}
        if (!w) {
            w = scale(PANEL_DEFAULT_WIDTH);
        }
        if (!h) {
            h = PANEL_DEFAULT_HEIGHT;
        }
        return { width: w, height: h };
    }

    function setLogVisible(flag) {
        try {
            localStorage.setItem(STORAGE_LOG_VISIBLE, flag ? "1" : "0");
        } catch (e) {}
    }

    function getLogVisible() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_LOG_VISIBLE);
        } catch (e) {}
        if (raw === "0") {
            return false;
        }
        return true;
    }

    function clampPanelPosition(panelEl) {
        var topPx = panelEl.style.top + "";
        var rightPx = panelEl.style.right + "";
        var topVal = parseInt(topPx.replace("px", ""), 10);
        var rightVal = parseInt(rightPx.replace("px", ""), 10);
        if (isNaN(topVal)) {
            topVal = 0;
        }
        if (isNaN(rightVal)) {
            rightVal = 0;
        }
        var vh = window.innerHeight;
        var vw = window.innerWidth;
        var ph = panelEl.offsetHeight;
        var pw = panelEl.offsetWidth;
        if (!ph || ph <= 0) {
            ph = 40;
        }
        if (!pw || pw <= 0) {
            pw = 340;
        }
        var minTop = 0;
        var maxTop = vh - ph / 2;
        var minRight = -pw / 2;
        var maxRight = vw - pw / 2;
        if (maxTop < 0) {
            maxTop = 0;
        }
        if (maxRight < 0) {
            maxRight = 0;
        }
        if (topVal < minTop) {
            topVal = minTop;
        } else {
            if (topVal > maxTop) {
                topVal = maxTop;
            }
        }
        if (rightVal < minRight) {
            rightVal = minRight;
        } else {
            if (rightVal > maxRight) {
                rightVal = maxRight;
            }
        }
        panelEl.style.top = String(topVal) + "px";
        panelEl.style.right = String(rightVal) + "px";
        try {
            localStorage.setItem("activityPlanState.panel.top", panelEl.style.top);
        } catch (e1) {}
        try {
            localStorage.setItem("activityPlanState.panel.right", panelEl.style.right);
        } catch (e2) {}
    }

    function setupResizeHandle(panel, bodyContainer) {
        var handle = document.createElement("div");
        handle.style.position = "absolute";
        handle.style.width = "12px";
        handle.style.height = "12px";
        handle.style.right = "6px";
        handle.style.bottom = "6px";
        handle.style.cursor = "se-resize";
        handle.style.background = "#333";
        handle.style.borderRadius = "2px";
        handle.style.display = "block";

        var isResizing = false;
        var startX = 0;
        var startY = 0;
        var startW = 0;
        var startH = 0;

        function toInt(s) {
            var n = parseInt(String(s).replace("px", ""), 10);
            if (isNaN(n)) {
                return 0;
            }
            return n;
        }

        handle.addEventListener("mousedown", function (e) {
            if (getPanelCollapsed()) {
                return;
            }
            isResizing = true;
            startX = e.clientX;
            startY = e.clientY;
            startW = toInt(panel.style.width);
            startH = panel.offsetHeight;
            e.preventDefault();
        });

        document.addEventListener("mousemove", function (e) {
            if (!isResizing) {
                return;
            }
            var dx = e.clientX - startX;
            var dy = e.clientY - startY;
            var newW = startW + dx;
            var newH = startH + dy;

            var minW = toInt(scale(PANEL_DEFAULT_WIDTH));
            if (newW < minW) {
                newW = minW;
            }
            if (newW > PANEL_MAX_WIDTH_PX) {
                newW = PANEL_MAX_WIDTH_PX;
            }

            var minH = PANEL_HEADER_HEIGHT_PX + 88;
            if (newH < minH) {
                newH = minH;
            }

            panel.style.width = String(newW) + "px";
            panel.style.height = String(newH) + "px";

            if (bodyContainer) {
                bodyContainer.style.display = "block";
                bodyContainer.style.height = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
                bodyContainer.style.maxHeight = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
                bodyContainer.style.overflowY = "auto";
            }
        });

        document.addEventListener("mouseup", function () {
            if (!isResizing) {
                return;
            }
            isResizing = false;
            setStoredPanelSize(panel.style.width, panel.style.height);
        });

        return handle;
    }

    function updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar) {
        var collapsed = getPanelCollapsed();
        if (collapsed) {
            panel.style.width = scale(PANEL_DEFAULT_WIDTH);
            panel.style.height = scale(PANEL_HEADER_HEIGHT_PX + 20);
            panel.style.overflow = "hidden";

            if (bodyContainer) {
                bodyContainer.style.display = "none";
            }
            if (resizeHandle) {
                resizeHandle.style.display = "none";
            }
            if (collapseBtn) {
                collapseBtn.textContent = "+";
            }
        } else {
            var size = getStoredPanelSize();
            panel.style.width = size.width;
            panel.style.height = size.height;
            panel.style.overflow = "visible";

            if (bodyContainer) {
                bodyContainer.style.display = "block";
            }
            if (resizeHandle) {
                resizeHandle.style.display = "block";
            }
            if (collapseBtn) {
                collapseBtn.textContent = "—";
            }
        }
    }

    function createPopup(options) {
        options = options || {};
        var title = options.title || "Popup";
        var description = options.description || "";
        var content = options.content || "";
        var width = options.width || "400px";
        var height = options.height || "auto";
        var maxWidth = options.maxWidth || "90%";
        var maxHeight = options.maxHeight || "90%";
        var onClose = options.onClose || null;
        var popupId = "clinsparkPopup_" + Date.now();
        var existing = document.getElementById(popupId);
        if (existing) {
            existing.remove();
        }
        var popup = document.createElement("div");
        popup.id = popupId;
        popup.style.position = "fixed";
        popup.style.top = "50%";
        popup.style.left = "50%";
        popup.style.transform = "translate(-50%, -50%)";
        popup.style.zIndex = "999998";
        popup.style.background = "#111";
        popup.style.color = "#fff";
        popup.style.border = "1px solid #444";
        popup.style.borderRadius = "8";
        popup.style.padding = "10px";
        popup.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        popup.style.fontSize = "14px";
        popup.style.width = width;
        popup.style.maxWidth = maxWidth;
        popup.style.height = height;
        popup.style.maxHeight = maxHeight;
        popup.style.boxSizing = "border-box";
        popup.style.display = "flex";
        popup.style.flexDirection = "column";
        popup.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.5)";

        var headerBar = document.createElement("div");
        headerBar.style.position = "relative";
        headerBar.style.display = "grid";
        headerBar.style.gridTemplateColumns = "1fr auto";
        headerBar.style.alignItems = "center";
        headerBar.style.gap = String(PANEL_HEADER_GAP_PX) + "px";
        headerBar.style.height = String(PANEL_HEADER_HEIGHT_PX) + "px";
        headerBar.style.boxSizing = "border-box";
        headerBar.style.padding = "0 12px";
        headerBar.style.borderBottom = "1px solid #444";
        headerBar.style.cursor = "move";
        headerBar.style.userSelect = "none";

        var titleContainer = document.createElement("div");
        titleContainer.style.display = "flex";
        titleContainer.style.flexDirection = "column";
        titleContainer.style.justifyContent = "center";

        var titleEl = document.createElement("div");
        titleEl.textContent = title;
        titleEl.style.fontWeight = "600";
        titleEl.style.textAlign = "left";
        titleContainer.appendChild(titleEl);

        if (description) {
            var descEl = document.createElement("div");
            descEl.textContent = description;
            descEl.style.fontSize = "12px";
            descEl.style.color = "#aaa";
            descEl.style.textAlign = "left";
            descEl.style.marginTop = "2px";
            titleContainer.appendChild(descEl);
        }

        headerBar.appendChild(titleContainer);

        var closeBtn = document.createElement("button");
        closeBtn.textContent = "✕";
        closeBtn.style.background = "transparent";
        closeBtn.style.color = "#fff";
        closeBtn.style.border = "none";
        closeBtn.style.cursor = "pointer";
        closeBtn.style.fontSize = "18px";
        closeBtn.style.lineHeight = "1";
        closeBtn.style.padding = "4px 8px";
        closeBtn.style.borderRadius = "4px";
        closeBtn.style.width = "32px";
        closeBtn.style.height = "32px";
        closeBtn.style.display = "flex";
        closeBtn.style.alignItems = "center";
        closeBtn.style.justifyContent = "center";
        closeBtn.addEventListener("mouseenter", function () {
            closeBtn.style.background = "#333";
        });
        closeBtn.addEventListener("mouseleave", function () {
            closeBtn.style.background = "transparent";
        });
        closeBtn.addEventListener("click", function () {
            if (onClose) {
                try {
                    onClose();
                } catch (e) {
                    log("Popup onClose error: " + String(e));
                }
            }
            document.removeEventListener("keydown", keyHandler, true);
            popup.remove();
        });
        headerBar.appendChild(closeBtn);
        popup.appendChild(headerBar);

        var bodyContainer = document.createElement("div");
        bodyContainer.style.flex = "1";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.overflowX = "hidden";
        bodyContainer.style.padding = "12px";
        bodyContainer.style.boxSizing = "border-box";
        if (typeof content === "string") {
            bodyContainer.innerHTML = content;
        } else if (content && content.nodeType === 1) {
            bodyContainer.appendChild(content);
        } else if (content && typeof content === "function") {
            var result = content();
            if (typeof result === "string") {
                bodyContainer.innerHTML = result;
            } else if (result && result.nodeType === 1) {
                bodyContainer.appendChild(result);
            }
        }
        popup.appendChild(bodyContainer);

        var isDragging = false;
        var startX = 0;
        var startY = 0;
        var startLeft = 0;
        var startTop = 0;

        function clampPopupPosition() {
            var rect = popup.getBoundingClientRect();
            var vw = window.innerWidth;
            var vh = window.innerHeight;
            var left = rect.left;
            var top = rect.top;
            var right = left + rect.width;
            var bottom = top + rect.height;
            var newLeft = left;
            var newTop = top;
            if (left < 0) {
                newLeft = 0;
            } else {
                if (right > vw) {
                    newLeft = vw - rect.width;
                }
            }
            if (top < 0) {
                newTop = 0;
            } else {
                if (bottom > vh) {
                    newTop = vh - rect.height;
                }
            }
            if (newLeft !== left || newTop !== top) {
                popup.style.left = String(newLeft) + "px";
                popup.style.top = String(newTop) + "px";
                popup.style.transform = "none";
            }
        }

        headerBar.addEventListener("mousedown", function (e) {
            if (e.target === closeBtn || closeBtn.contains(e.target)) {
                return;
            }
            isDragging = true;
            headerBar.style.cursor = "grabbing";
            startX = e.clientX;
            startY = e.clientY;
            var rect = popup.getBoundingClientRect();
            startLeft = rect.left;
            startTop = rect.top;
            e.preventDefault();
        });

        document.addEventListener("mousemove", function (e) {
            if (!isDragging) {
                return;
            }
            var dx = e.clientX - startX;
            var dy = e.clientY - startY;
            var newLeft = startLeft + dx;
            var newTop = startTop + dy;
            popup.style.left = String(newLeft) + "px";
            popup.style.top = String(newTop) + "px";
            popup.style.transform = "none";
            clampPopupPosition();
        });

        document.addEventListener("mouseup", function () {
            if (!isDragging) {
                return;
            }
            isDragging = false;
            headerBar.style.cursor = "move";
            clampPopupPosition();
        });

        function keyHandler(e) {
            var code = e.key || e.code || "";
            if (code === "Escape" || code === "Esc") {
                log("Popup: Esc pressed; closing");
                closeBtn.click();
                e.preventDefault();
                e.stopPropagation();
            }
        }

        document.addEventListener("keydown", keyHandler, true);

        document.body.appendChild(popup);

        return {
            element: popup,
            close: function () {
                if (onClose) {
                    try {
                        onClose();
                    } catch (e) {
                        log("Popup onClose error: " + String(e));
                    }
                }
                document.removeEventListener("keydown", keyHandler, true);
                popup.remove();
            },
            setContent: function (newContent) {
                bodyContainer.innerHTML = "";
                if (typeof newContent === "string") {
                    bodyContainer.innerHTML = newContent;
                } else if (newContent && newContent.nodeType === 1) {
                    bodyContainer.appendChild(newContent);
                }
            }
        };
    }

    function showWrongPagePopup(featureName, requiredPage, currentPath) {
        // Create warning popup content
        var warningContainer = document.createElement("div");
        warningContainer.style.padding = "20px";
        warningContainer.style.textAlign = "center";

        var warningIcon = document.createElement("div");
        warningIcon.innerHTML = "⚠️";
        warningIcon.style.fontSize = "48px";
        warningIcon.style.marginBottom = "15px";

        var warningTitle = document.createElement("div");
        warningTitle.textContent = "Wrong Page Detected";
        warningTitle.style.fontSize = "18px";
        warningTitle.style.fontWeight = "bold";
        warningTitle.style.marginBottom = "10px";
        warningTitle.style.color = "#ff6b6b";

        var warningMessage = document.createElement("div");
        warningMessage.textContent = "You must be on the " + featureName + " page to use this feature.";
        warningMessage.style.marginBottom = "8px";
        warningMessage.style.lineHeight = "1.4";

        var currentPage = document.createElement("div");
        currentPage.textContent = "Current page: " + currentPath;
        currentPage.style.fontSize = "12px";
        currentPage.style.color = "#999";
        currentPage.style.marginBottom = "8px";

        var correctPage = document.createElement("div");
        correctPage.textContent = "Required page: " + requiredPage;
        correctPage.style.fontSize = "12px";
        correctPage.style.color = "#999";
        correctPage.style.marginBottom = "20px";

        var okButton = document.createElement("button");
        okButton.textContent = "OK";
        okButton.style.background = "#5b43c7";
        okButton.style.color = "#fff";
        okButton.style.border = "none";
        okButton.style.borderRadius = "4px";
        okButton.style.padding = "8px 24px";
        okButton.style.cursor = "pointer";
        okButton.style.fontSize = "14px";
        okButton.onmouseenter = function() { this.style.background = "#4a35a6"; };
        okButton.onmouseleave = function() { this.style.background = "#5b43c7"; };

        warningContainer.appendChild(warningIcon);
        warningContainer.appendChild(warningTitle);
        warningContainer.appendChild(warningMessage);
        warningContainer.appendChild(currentPage);
        warningContainer.appendChild(correctPage);
        warningContainer.appendChild(okButton);

        var warningPopup = createPopup({
            title: featureName + " - Page Warning",
            content: warningContainer,
            width: "400px",
            height: "auto",
            onClose: function () {
                log(featureName + ": warning popup closed");
            }
        });

        okButton.addEventListener("click", function () {
            warningPopup.close();
            log(featureName + ": warning popup acknowledged by user");
        });

        return warningPopup;
    }

    function addButtonToPanel(label, handler) {
        if (!btnRowRef) {
            PENDING_BUTTONS.push({ label: label, handler: handler });
            return;
        }
        var btn = document.createElement("button");
        btn.textContent = label;
        btn.style.background = "#ddd";
        btn.style.color = "#000";
        btn.style.border = "none";
        btn.style.borderRadius = "6px";
        btn.style.padding = "8px";
        btn.style.cursor = "pointer";
        btn.addEventListener("click", function () {
            try {
                handler();
            } catch (e) {
                log("Button handler error: " + String(e && e.message ? e.message : e));
            }
        });
        btnRowRef.appendChild(btn);
    }


    function makePanel() {
        var prior = document.getElementById(PANEL_ID);
        if (prior) {
            return prior;
        }
        var panel = document.createElement("div");
        panel.id = PANEL_ID;
        panel.style.position = "fixed";
        var savedTop = getStoredPos("activityPlanState.panel.top", "20px");
        var savedRight = getStoredPos("activityPlanState.panel.right", "20px");
        var savedSize = getStoredPanelSize();
        panel.style.top = savedTop;
        panel.style.right = savedRight;
        panel.style.zIndex = "999999";
        panel.style.background = "#111";
        panel.style.color = "#fff";
        panel.style.border = "1px solid #444";
        panel.style.borderRadius = scale(PANEL_BORDER_RADIUS_PX);
        panel.style.padding = scale(PANEL_PADDING_PX);
        panel.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        panel.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        panel.style.minWidth = scale(PANEL_DEFAULT_WIDTH);
        panel.style.width = savedSize.width;
        if (savedSize.height && savedSize.height !== PANEL_DEFAULT_HEIGHT) {
            panel.style.height = savedSize.height;
        } else {
            panel.style.height = PANEL_DEFAULT_HEIGHT;
        }
        panel.style.boxSizing = "border-box";
        var headerBar = document.createElement("div");
        headerBar.style.position = "relative";
        headerBar.style.display = "grid";
        headerBar.style.gridTemplateColumns = "auto 1fr auto";
        headerBar.style.alignItems = "center";
        headerBar.style.gap = scale(PANEL_HEADER_GAP_PX);
        headerBar.style.height = scale(PANEL_HEADER_HEIGHT_PX);
        headerBar.style.boxSizing = "border-box";
        headerBar.style.cursor = "grab";
        headerBar.style.userSelect = "none";
        var leftSpacer = document.createElement("div");
        leftSpacer.style.width = scale(HEADER_LEFT_SPACER_WIDTH_PX);
        var title = document.createElement("div");
        title.textContent = "Automator";
        title.style.fontWeight = HEADER_FONT_WEIGHT;
        title.style.textAlign = "center";
        title.style.justifySelf = "center";
        title.style.transform = "translateX(" + scale(HEADER_TITLE_OFFSET_PX) + ")";
        title.style.paddingBottom = scale(8);
        headerBar.appendChild(title);
        headerBar.appendChild(leftSpacer);
        var rightControls = document.createElement("div");
        rightControls.style.display = "inline-flex";
        rightControls.style.alignItems = "center";
        rightControls.style.gap = scale(PANEL_HEADER_GAP_PX);
        var collapseBtn = document.createElement("button");
        collapseBtn.textContent = getPanelCollapsed() ? "Expand" : "Collapse";
        collapseBtn.style.background = "transparent";
        collapseBtn.style.color = "#fff";
        collapseBtn.style.border = "none";
        collapseBtn.style.cursor = "pointer";
        collapseBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        var closeBtn = document.createElement("button");
        closeBtn.textContent = CLOSE_BTN_TEXT;
        closeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        closeBtn.style.background = "transparent";
        closeBtn.style.color = "#fff";
        closeBtn.style.border = "none";
        closeBtn.style.cursor = "pointer";

        var settingsBtn = document.createElement("button");
        settingsBtn.textContent = "⚙";
        settingsBtn.title = "Settings";
        settingsBtn.style.background = "transparent";
        settingsBtn.style.color = "#fff";
        settingsBtn.style.border = "none";
        settingsBtn.style.cursor = "pointer";
        settingsBtn.style.fontSize = scale(16);
        settingsBtn.style.padding = "4px 6px";
        settingsBtn.style.borderRadius = "4px";
        settingsBtn.style.display = "flex";
        settingsBtn.style.alignItems = "center";
        settingsBtn.style.justifyContent = "center";
        settingsBtn.addEventListener("click", function() {
            log("Settings: Button clicked");
            if (SETTINGS_POPUP_REF) {
                log("Settings: Closing existing popup");
                SETTINGS_POPUP_REF.close();
                SETTINGS_POPUP_REF = null;
            } else {
                SETTINGS_POPUP_REF = openSettingsPopup();
            }
        });

        rightControls.appendChild(settingsBtn);
        rightControls.appendChild(collapseBtn);
        rightControls.appendChild(closeBtn);

        headerBar.appendChild(rightControls);
        panel.appendChild(headerBar);
        var bodyContainer = document.createElement("div");
        bodyContainer.style.display = "block";
        bodyContainer.style.height = "calc(100% - " + scale(PANEL_HEADER_HEIGHT_PX) + ")";
        bodyContainer.style.maxHeight = "calc(100% - " + scale(PANEL_HEADER_HEIGHT_PX) + ")";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.boxSizing = "border-box";
        var btnRow = document.createElement("div");
        btnRow.style.display = "grid";
        btnRow.style.gridTemplateColumns = "1fr 1fr";
        btnRow.style.gap = scale(BUTTON_GAP_PX);
        btnRowRef = btnRow;
        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#6c757d";
        pauseBtn.style.color = "#fff";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        pauseBtn.style.padding = scale(BUTTON_PADDING_PX);
        pauseBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        pauseBtn.style.cursor = "pointer";
        pauseBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        pauseBtn.onmouseleave = function() { this.style.background = "#6c757d"; };
        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Run Barcode";
        runBarcodeBtn.style.background = "#5b43c7";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        runBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        runBarcodeBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
        var findAeBtn = document.createElement("button");
        findAeBtn.textContent = "Find Adverse Event";
        findAeBtn.style.background = "#4a90e2";
        findAeBtn.style.color = "#fff";
        findAeBtn.style.border = "none";
        findAeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findAeBtn.style.padding = scale(BUTTON_PADDING_PX);
        findAeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findAeBtn.style.cursor = "pointer";
        findAeBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        findAeBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var findFormBtn = document.createElement("button");
        findFormBtn.textContent = "Find Form";
        findFormBtn.style.background = "#4a90e2";
        findFormBtn.style.color = "#fff";
        findFormBtn.style.border = "none";
        findFormBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findFormBtn.style.padding = scale(BUTTON_PADDING_PX);
        findFormBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findFormBtn.style.cursor = "pointer";
        findFormBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        findFormBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var findStudyEventsBtn = document.createElement("button");
        findStudyEventsBtn.textContent = "Find Study Events";
        findStudyEventsBtn.style.background = "#4a90e2";
        findStudyEventsBtn.style.color = "#fff";
        findStudyEventsBtn.style.border = "none";
        findStudyEventsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findStudyEventsBtn.style.padding = scale(BUTTON_PADDING_PX);
        findStudyEventsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findStudyEventsBtn.style.cursor = "pointer";
        findStudyEventsBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        findStudyEventsBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var openEligBtn = document.createElement("button");
        openEligBtn.textContent = "Cohort Eligibility";
        openEligBtn.style.background = "#4a90e2";
        openEligBtn.style.color = "#fff";
        openEligBtn.style.border = "none";
        openEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        openEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        openEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        openEligBtn.style.cursor = "pointer";
        openEligBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        openEligBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var subjectEligBtn = document.createElement("button");
        subjectEligBtn.textContent = "Subject Eligibility";
        subjectEligBtn.style.background = "#4a90e2";
        subjectEligBtn.style.color = "#fff";
        subjectEligBtn.style.border = "none";
        subjectEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        subjectEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        subjectEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        subjectEligBtn.style.cursor = "pointer";
        subjectEligBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        subjectEligBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var saBuilderBtn = document.createElement("button");
        saBuilderBtn.textContent = "Scheduled Activities Builder";
        saBuilderBtn.style.background = "#5b43c7";
        saBuilderBtn.style.color = "#fff";
        saBuilderBtn.style.border = "none";
        saBuilderBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        saBuilderBtn.style.padding = scale(BUTTON_PADDING_PX);
        saBuilderBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        saBuilderBtn.style.cursor = "pointer";
        saBuilderBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        saBuilderBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };

        var parseMethodBtn = document.createElement("button");
        parseMethodBtn.textContent = "Item Method Forms";
        parseMethodBtn.style.background = "#4a90e2";
        parseMethodBtn.style.color = "#fff";
        parseMethodBtn.style.border = "none";
        parseMethodBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        parseMethodBtn.style.padding = scale(BUTTON_PADDING_PX);
        parseMethodBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        parseMethodBtn.style.cursor = "pointer";
        parseMethodBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        parseMethodBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var searchMethodsBtn = document.createElement("button");
        searchMethodsBtn.textContent = "Search Methods";
        searchMethodsBtn.style.background = "#5b43c7";
        searchMethodsBtn.style.color = "#fff";
        searchMethodsBtn.style.border = "none";
        searchMethodsBtn.style.borderRadius = "6px";
        searchMethodsBtn.style.padding = "8px";
        searchMethodsBtn.style.cursor = "pointer";
        searchMethodsBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        searchMethodsBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
        searchMethodsBtn.addEventListener("click", function() {
            log("[SearchMethods] Button clicked");
            openMethodsLibraryModal();
        });

        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#6c757d";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        toggleLogsBtn.style.padding = scale(BUTTON_PADDING_PX);
        toggleLogsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        toggleLogsBtn.style.cursor = "pointer";
        toggleLogsBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import I/E";
        importEligBtn.style.background = "#5b43c7";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        importEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        importEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        importEligBtn.style.cursor = "pointer";
        importEligBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        importEligBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };

        var parseStudyEventBtn = document.createElement("button");
        parseStudyEventBtn.textContent = "Parse Study Event";
        parseStudyEventBtn.style.background = "#4a90e2";
        parseStudyEventBtn.style.color = "#fff";
        parseStudyEventBtn.style.border = "none";
        parseStudyEventBtn.style.borderRadius = "6px";
        parseStudyEventBtn.style.padding = "8px";
        parseStudyEventBtn.style.cursor = "pointer";
        parseStudyEventBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        parseStudyEventBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        parseStudyEventBtn.addEventListener("click", function() {
            log("[ParseStudyEvent] Button clicked");
            APS_ParseStudyEvent();
        });

        var parseDeviationBtn = document.createElement("button");
        parseDeviationBtn.textContent = "Parse Deviation";
        parseDeviationBtn.style.background = "#5b43c7";
        parseDeviationBtn.style.color = "#fff";
        parseDeviationBtn.style.border = "none";
        parseDeviationBtn.style.borderRadius = "6px";
        parseDeviationBtn.style.padding = "8px";
        parseDeviationBtn.style.cursor = "pointer";
        parseDeviationBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        parseDeviationBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
        parseDeviationBtn.addEventListener("click", function() {
            log("[ParseDeviation] Button clicked");
            APS_ParseDeviation();
        });

        var clearLogsBtn = document.createElement("button");
        clearLogsBtn.textContent = "Clear Logs";
        clearLogsBtn.style.background = "#6c757d";
        clearLogsBtn.style.color = "#fff";
        clearLogsBtn.style.border = "none";
        clearLogsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        clearLogsBtn.style.padding = scale(BUTTON_PADDING_PX);
        clearLogsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        clearLogsBtn.style.cursor = "pointer";
        clearLogsBtn.style.fontWeight = "500";
        clearLogsBtn.style.transition = "background 0.2s";
        clearLogsBtn.onmouseenter = () => { clearLogsBtn.style.background = "#5a6268"; };
        clearLogsBtn.onmouseleave = () => { clearLogsBtn.style.background = "#6c757d"; };

        var panelButtons = [
            { el: runBarcodeBtn, label: "Run Barcode" },
            { el: saBuilderBtn, label: "Scheduled Activities Builder" },
            { el: searchMethodsBtn, label: "Search Methods" },
            { el: parseDeviationBtn, label: "Parse Deviation" },
            { el: importEligBtn, label: "Import I/E" },
            { el: findAeBtn, label: "Find Adverse Event" },
            { el: findFormBtn, label: "Find Form" },
            { el: findStudyEventsBtn, label: "Find Study Events" },
            { el: parseMethodBtn, label: "Item Method Forms" },
            { el: openEligBtn, label: "Cohort Eligibility" },
            { el: subjectEligBtn, label: "Subject Eligibility" },
            { el: parseStudyEventBtn, label: "Parse Study Event" },
            { el: pauseBtn, label: "Pause" },
            { el: clearLogsBtn, label: "Clear Logs" },
            { el: toggleLogsBtn, label: "Hide Logs" }
        ];

        for (var bi = 0; bi < panelButtons.length; bi++) {
            var btnItem = panelButtons[bi];
            if (isButtonVisible(btnItem.label)) {
                btnRow.appendChild(btnItem.el);
            }
        }

        bodyContainer.appendChild(btnRow);
        var status = document.createElement("div");
        status.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        status.style.background = "#1a1a1a";
        status.style.border = "1px solid #333";
        status.style.borderRadius = scale(STATUS_BORDER_RADIUS_PX);
        status.style.padding = scale(STATUS_PADDING_PX);
        status.style.fontSize = scale(STATUS_FONT_SIZE_PX);
        status.style.whiteSpace = "pre-wrap";
        status.textContent = "Ready";
        bodyContainer.appendChild(status);

        // UI Scale Control
        var scaleControl = document.createElement("div");
        scaleControl.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        scaleControl.style.background = "#1a1a1a";
        scaleControl.style.border = "1px solid #333";
        scaleControl.style.borderRadius = scale(STATUS_BORDER_RADIUS_PX);
        scaleControl.style.padding = scale(STATUS_PADDING_PX);
        scaleControl.style.fontSize = scale(STATUS_FONT_SIZE_PX);

        var scaleLabel = document.createElement("div");
        scaleLabel.textContent = "UI Scale: " + Math.round(UI_SCALE * 100) + "%";
        scaleLabel.style.marginBottom = "4px";
        scaleLabel.style.color = "#fff";

        var scaleSlider = document.createElement("input");
        scaleSlider.type = "range";
        scaleSlider.min = "50";
        scaleSlider.max = "100";
        scaleSlider.value = String(UI_SCALE * 100);
        scaleSlider.step = "10";
        scaleSlider.style.width = "100%";
        scaleSlider.style.cursor = "pointer";

        scaleSlider.addEventListener("input", function() {
            var newScale = parseFloat(this.value) / 100;
            updateUIScale(newScale);
            scaleLabel.textContent = "UI Scale: " + Math.round(newScale * 100) + "%";
            status.textContent = "UI Scale updated to " + Math.round(newScale * 100) + "% - Refresh to see changes";
        });

        scaleSlider.addEventListener("change", function() {
            log("UI Scale changed to " + Math.round(UI_SCALE * 100) + "%");
        });

        scaleControl.appendChild(scaleLabel);
        scaleControl.appendChild(scaleSlider);
        bodyContainer.appendChild(scaleControl);

        var logBox = document.createElement("div");
        logBox.id = LOG_ID;
        logBox.style.marginTop = scale(LOG_MARGIN_TOP_PX);
        logBox.style.height = scale(LOG_HEIGHT_PX);
        logBox.style.overflowY = "auto";
        logBox.style.background = "#141414";
        logBox.style.border = "1px solid #333";
        logBox.style.borderRadius = scale(LOG_BORDER_RADIUS_PX);
        logBox.style.padding = scale(LOG_PADDING_PX);
        logBox.style.fontSize = scale(LOG_FONT_SIZE_PX);
        logBox.style.color = "#00000";
        logBox.style.whiteSpace = "pre-wrap";
        logBox.style.wordBreak = "break-word";
        logBox.style.overflowWrap = "anywhere";
        if (!logVisible) {
            logBox.style.display = "none";
        }
        bodyContainer.appendChild(logBox);
        try {
            var rawLogs = localStorage.getItem("activityPlanState.logs");
            if (rawLogs) {
                var logsArr = JSON.parse(rawLogs);
                if (Array.isArray(logsArr)) {
                    for (var i = 0; i < logsArr.length; i++) {
                        var line = document.createElement("div");
                        line.textContent = logsArr[i];
                        logBox.appendChild(line);
                    }
                    logBox.scrollTop = logBox.scrollHeight;
                }
            }
        } catch (e) {
            console.log("[APS] Failed to restore logs: " + e);
        }
        importEligBtn.addEventListener("click", function () {
            log("ImportElig: button clicked");
            startImportEligibilityMapping();
        });
        parseMethodBtn.addEventListener("click", function () {
            openParseMethod();
        });
        pauseBtn.addEventListener("click", function () {
            var nowPaused = isPaused();
            if (nowPaused) {
                setPaused(false);
                pauseBtn.textContent = "Pause";
                status.textContent = "Resumed";
                SA_BUILDER_PAUSE = false;
                log("Resumed");
            } else {
                setPaused(true);
                pauseBtn.textContent = "Resume";
                status.textContent = "Paused";
                log("Paused");
                clearAllRunState();
                SA_BUILDER_PAUSE = true;
                SA_BUILDER_CANCELLED = true;
                clearEligibilityWorkingState();
                if (SA_BUILDER_PROGRESS_POPUP_REF) {
                    try {
                        SA_BUILDER_PROGRESS_POPUP_REF.close();
                        SA_BUILDER_PROGRESS_POPUP_REF = null;
                    } catch (e) {}
                }
                if (SA_BUILDER_POPUP_REF) {
                    try {
                        SA_BUILDER_POPUP_REF.close();
                        SA_BUILDER_POPUP_REF = null;
                    } catch (e) {}
                }
            }
        });
        clearLogsBtn.addEventListener("click", function () {
            clearLogs();
            status.textContent = "Logs cleared";
        });
        runBarcodeBtn.addEventListener("click", async function () {
            log("Run Barcode: button clicked");
            await APS_RunBarcode();
        });
        findAeBtn.addEventListener("click", function () {
            openAndLocateAdverseEvent();
        });
        findFormBtn.addEventListener("click", function () {
            openFindForm();
        });
        findStudyEventsBtn.addEventListener("click", function () {
            openFindStudyEvents();
        });

        openEligBtn.addEventListener("click", function () {
            cohortEligibilityFeature();
        });

        subjectEligBtn.addEventListener("click", function () {
            subjectEligibilityFeature();
        });
        saBuilderBtn.addEventListener("click", async function () {
            SA_BUILDER_CANCELLED = false;
            log("SA Builder: button clicked");
            if (SA_BUILDER_PAUSE) {
                log("SA Builder: Paused");
                return;
            }
            log("SA Builder: starting");
            await runSABuilder();
        });
        toggleLogsBtn.addEventListener("click", function () {
            var currentlyVisible = getLogVisible();
            var newVisible = !currentlyVisible;
            setLogVisible(newVisible);
            if (newVisible) {
                logBox.style.display = "block";
                toggleLogsBtn.textContent = "Hide Logs";
                log("Logs shown");
            } else {
                logBox.style.display = "none";
                toggleLogsBtn.textContent = "Show Logs";
                log("Logs hidden");
            }
        });
        collapseBtn.addEventListener("click", function () {
            var collapsed = getPanelCollapsed();
            if (collapsed) {
                setPanelCollapsed(false);
            } else {
                setPanelCollapsed(true);
            }
            updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar);
        });
        closeBtn.addEventListener("click", function () {
            panel.remove();
        });
        panel.appendChild(bodyContainer);
        var resizeHandle = setupResizeHandle(panel, bodyContainer);
        panel.appendChild(resizeHandle);
        var isDragging = false;
        var startX = 0;
        var startY = 0;
        var startTop = 0;
        var startRight = 0;
        function pxToInt(px) {
            var n = parseInt(String(px).replace("px", ""), 10);
            if (isNaN(n)) {
                return 0;
            }
            return n;
        }
        headerBar.addEventListener("mousedown", function (e) {
            if (e.target === collapseBtn || e.target === closeBtn || collapseBtn.contains(e.target) || closeBtn.contains(e.target)) {
                return;
            }
            isDragging = true;
            headerBar.style.cursor = "grabbing";
            startX = e.clientX;
            startY = e.clientY;
            startTop = pxToInt(panel.style.top);
            startRight = pxToInt(panel.style.right);
            e.preventDefault();
        });
        document.addEventListener("mousemove", function (e) {
            if (!isDragging) {
                return;
            }
            var dx = e.clientX - startX;
            var dy = e.clientY - startY;
            var newTop = startTop + dy;
            var newRight = startRight - dx;
            panel.style.top = String(newTop) + "px";
            panel.style.right = String(newRight) + "px";
            clampPanelPosition(panel);
        });
        document.addEventListener("mouseup", function () {
            if (!isDragging) {
                return;
            }
            isDragging = false;
            headerBar.style.cursor = "grab";
            clampPanelPosition(panel);
        });
        document.body.appendChild(panel);
        applyPanelHiddenState(panel);
        var collapsedInit = getPanelCollapsed();
        updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar);


        var t = pxToInt(panel.style.top);
        var r2 = pxToInt(panel.style.right);
        var vh = window.innerHeight || 800;
        var vw = window.innerWidth || 1280;
        var ph = panel.offsetHeight || 40;
        var pw = panel.offsetWidth || 340;
        var minTop = 0;
        var maxTop = vh - ph;
        var minRight = -pw / 2;
        var maxRight = vw - pw / 2;
        if (maxTop < 0) {
            maxTop = 0;
        }
        if (maxRight < 0) {
            maxRight = 0;
        }
        var offTop = false;
        var offRight = false;
        if (t < minTop) {
            offTop = true;
        } else {
            if (t > maxTop) {
                offTop = true;
            }
        }
        if (r2 < minRight) {
            offRight = true;
        } else {
            if (r2 > maxRight) {
                offRight = true;
            }
        }
        if (offTop || offRight) {
            panel.style.top = "20px";
            panel.style.right = "20px";
            try {
                localStorage.setItem("activityPlanState.panel.top", panel.style.top);
            } catch (e3) {}
            try {
                localStorage.setItem("activityPlanState.panel.right", panel.style.right);
            } catch (e4) {}
        }

        if (PENDING_BUTTONS.length > 0) {
            var ia = 0;
            while (ia < PENDING_BUTTONS.length) {
                var it = PENDING_BUTTONS[ia];
                addButtonToPanel(it.label, it.handler);
                ia = ia + 1;
            }
            PENDING_BUTTONS = [];
        }
        log("Panel ready");
        return panel;
    }

    function init() {
        makePanel();
        window.APS_AddButton = function (label, handler) {
            addButtonToPanel(label, handler);
        };
        bindPanelHotkeyOnce();

        parseDeviationCheckOnPageLoad();

        // Restore Subject Eligibility spinner popup if pending
        if (isSubjectEligPending()) {
            var identifier = null;
            try {
                identifier = localStorage.getItem(STORAGE_SUBJECT_ELIG_IDENTIFIER);
            } catch (e) {}
            if (identifier) {
                log("SubjectElig: Restoring spinner popup for identifier: " + identifier);
                showSubjectEligSpinnerPopup(identifier);
            }
        }

        // Cohort Eligibility auto-tab detection for subject show pages
        var cohortAutoTabFlag = null;
        try {
            cohortAutoTabFlag = localStorage.getItem(STORAGE_COHORT_ELIG_AUTO_TAB);
        } catch (e) {}

        // Subject Eligibility auto-tab detection for subject show pages
        var subjectAutoTabFlag = null;
        try {
            subjectAutoTabFlag = localStorage.getItem(STORAGE_SUBJECT_ELIG_AUTO_TAB);
        } catch (e) {}

        if ((cohortAutoTabFlag === "1" || subjectAutoTabFlag === "1") && location.pathname.indexOf("/subjects/show/") !== -1) {
            var featureName = cohortAutoTabFlag === "1" ? "CohortElig" : "SubjectElig";
            log(featureName + ": Auto-tab detected on subject show page, navigating to Eligibility tab");

            // Clear Subject Eligibility data after successful navigation
            if (subjectAutoTabFlag === "1") {
                setTimeout(function() {
                    clearSubjectEligData();
                    log("SubjectElig: Process complete, data cleared");
                }, 4000);
            }

            // Wait for eligibility tab link to be available with retry logic
            var attemptCount = 0;
            var maxAttempts = 20;
            var attemptInterval = 300;

            function tryClickEligibilityTab() {
                attemptCount++;

                var eligTabLink = document.getElementById("eligibilityTabLink");
                if (eligTabLink) {
                    log(featureName + ": Found eligibilityTabLink on attempt " + attemptCount + ", clicking...");
                    eligTabLink.click();
                    log(featureName + ": Eligibility tab clicked successfully");
                    return;
                }

                var tabLinks = document.querySelectorAll("a[href='#eligibilityTab']");
                if (tabLinks && tabLinks.length > 0) {
                    log(featureName + ": Found eligibility tab link via selector on attempt " + attemptCount + ", clicking...");
                    tabLinks[0].click();
                    log(featureName + ": Eligibility tab clicked via selector");
                    return;
                }

                if (attemptCount < maxAttempts) {
                    log(featureName + ": Eligibility tab link not found, retrying... (attempt " + attemptCount + "/" + maxAttempts + ")");
                    setTimeout(tryClickEligibilityTab, attemptInterval);
                } else {
                    log(featureName + ": Could not find eligibility tab link after " + maxAttempts + " attempts");
                }
            }

            setTimeout(tryClickEligibilityTab, 500);
        }

        // Subject Eligibility page detection
        if (isSubjectEligPending() && isSubjectsListPage()) {
            log("SubjectElig: Pending detected on subjects list page, processing...");
            setTimeout(processSubjectEligOnSubjectsList, 2000);
            return;
        }

        // Cohort Eligibility page detection
        if (isCohortEligRunning()) {
            var phase = getCohortEligPhase();
            log("CohortElig: Running detected, phase=" + phase);
            if (phase === "collectInitials" && isCohortShowPage()) {
                log("CohortElig: On cohort page, will collect assignments");
                setTimeout(processCohortEligOnCohortPage, 2000);
                return;
            } else if (phase === "crossReference" && isSubjectsListPage()) {
                log("CohortElig: On subjects list, will cross-reference");
                setTimeout(processCohortEligOnSubjectsList, 2000);
                return;
            }
        }

        if (isPaused()) {
            log("Paused; automation halted");
            return;
        }
        var onBarcodeSubjects = isBarcodeSubjectsPage();
        if (onBarcodeSubjects) {
            processBarcodeSubjectsPage();
            return;
        }
        if (location.pathname === "/secure/study/data/list") {
            var parseMethodCompleted = null;
            try { parseMethodCompleted = localStorage.getItem(STORAGE_PARSE_METHOD_COMPLETED); } catch (e) {}
            if (parseMethodCompleted === "1") {
                log("ParseMethod: skipping processFindFormOnList because Parse Method completed");
                checkAndRestoreParseMethodPopup();
            } else {
                processFindFormOnList();
                processFindStudyEventsOnList();
            }
        }


        // Check for manual navigation away from Import Eligibility process
        var pendingPopup = null;
        try {
            pendingPopup = localStorage.getItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
        } catch (e) {
            pendingPopup = null;
        }

        var runModeRaw = null;
        try {
            runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            runModeRaw = null;
        }

        if (pendingPopup === "1" && runModeRaw !== RUNMODE_ELIG_IMPORT && !isEligibilityListPage()) {
            log("ImportElig: pending popup detected but user navigated away before confirming; clearing state");
            try {
                localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
            } catch (e) {}
        }

        if (runModeRaw === RUNMODE_ELIG_IMPORT && pendingPopup !== "1" && !isEligibilityListPage()) {
            log("ImportElig: run mode set but user navigated away manually; clearing run mode");
            try {
                localStorage.removeItem(STORAGE_RUN_MODE);
            } catch (e) {}
            return;
        }

        if (runModeRaw === RUNMODE_ELIG_IMPORT) {

            if (pendingPopup === "1") {
                if (!isEligibilityListPage()) {
                    log("ImportElig: pending popup but not on list page; redirecting");
                    location.href = ELIGIBILITY_LIST_URL;
                    return;
                }

                log("ImportElig: pending popup detected; showing popup");

                try {
                    localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
                } catch (e) {
                }

                startImportEligibilityMapping();
                return;
            }

            if (!isEligibilityListPage()) {
                log("ImportElig: run mode set but not on list page; redirecting");
                location.href = ELIGIBILITY_LIST_URL;
                return;
            }

            log("ImportElig: run mode set on list page; waiting 3s before resuming");
            setTimeout(function () {
                executeEligibilityMappingAutomation();
            }, 3000);

            return;
        }

    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();