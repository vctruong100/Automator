
// ==UserScript==
// @name        ClinSpark Automator
// @namespace   vinh.activity.plan.state
// @version     1.8.1
// @description Automate various tasks in ClinSpark platform
// @match       https://cenexel.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Automator.js
// @run-at      document-idle
// @grant       GM.openInTab
// @grant       GM_openInTab
// @grant       GM.xmlHttpRequest
// @ts-check
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
            var errorPopup = createPopup({
                title: "Scheduled Activities Builder",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#f66;font-size:16px;margin-bottom:16px;">⚠️ Wrong Page</p><p>You must be on the Activity Plans Show page to use this feature.</p><p style="margin-top:12px;font-size:12px;color:#888;">Required URL: ' + SA_BUILDER_TARGET_URL + '</p></div>',
                width: "450px",
                height: "auto"
            });
            log("SA Builder: wrong page - " + location.href);
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
            description: "Check eligibility criteria for multiple subjects at once",
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
            description: "Check eligibility criteria for individual subjects",
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
        var match = false;
        if (n.key && n.key.toUpperCase() === PANEL_TOGGLE_KEY.toUpperCase()) {
            match = true;
        } else {
            if (n.code && n.code.toUpperCase() === PANEL_TOGGLE_KEY.toUpperCase()) {
                match = true;
            } else {
                if (typeof n.keyCode === "number") {
                    if (PANEL_TOGGLE_KEY.toUpperCase() === "F2") {
                        if (n.keyCode === 113) {
                            match = true;
                        }
                    }
                }
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
        }
        document.addEventListener("keydown", handler, true);
        window.addEventListener("keydown", handler, true);
        window.__APS_HOTKEY_BOUND = true;
        log("Hotkey: bound for " + String(PANEL_TOGGLE_KEY));
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

        var popup = createPopup({ title: FORM_POPUP_TITLE, description: "Auto-navigate to Form data page based on keywords and status", content: container, width: "520px", height: "auto" });

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

        var popup = createPopup({ title: STUDY_EVENT_POPUP_TITLE, description: "Auto-navigate to Study Events data page based on keywords", content: container, width: "520px", height: "auto" });

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
        var popup = createPopup({ title: "Parse Method", description: "Extract and parse method data from forms", content: container, width: "480px", height: "auto", onClose: function() { stopParseMethodAutomation(); } });
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


    var STORAGE_ELIG_FORM_EXCLUSION = "activityPlanState.elig.formExclusion";
    var STORAGE_ELIG_FORM_PRIORITY = "activityPlanState.elig.formPriority";
    var STORAGE_ELIG_FORM_PRIORITY_ONLY = "activityPlanState.elig.formPriorityOnly";
    var STORAGE_ELIG_PLAN_PRIORITY = "activityPlanState.elig.planPriority";
    var STORAGE_ELIG_IGNORE = "activityPlanState.elig.ignore";

    const STORAGE_ELIG_LAST_PLAN = "activityPlanState.elig.lastPlan";
    const STORAGE_ELIG_LAST_SA = "activityPlanState.elig.lastSA";
    const STORAGE_ELIG_LAST_ITEMREF = "activityPlanState.elig.lastItemRef";

    var DEFAULT_FORM_EXCLUSION = "check";
    var DEFAULT_FORM_PRIORITY = "mh, bm, review, process, dm, rep, subs, med, elg_pi, vitals, ecg";

    function SubjectEligibilityFunctions() {}
    // Check if run mode is still set for the given mode.
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

    // Clear persisted selected volunteer ids.
    function clearSelectedVolunteerIds() {
        try {
            localStorage.removeItem(STORAGE_SELECTED_IDS);
            log("SelectedIds cleared");
        } catch (e) {}
    }

    // Clear pending IDs from storage.
    function clearPendingIds() {
        try {
            localStorage.removeItem(STORAGE_PENDING);
            log("Pending IDs cleared");
        } catch (e) {}
    }

    // Clear run mode.
    function clearRunMode() {
        try {
            localStorage.removeItem(STORAGE_RUN_MODE);
            log("RunMode cleared");
        } catch (e) {}
    }

    // Clear continue-epoch flag.
    function clearContinueEpoch() {
        try {
            localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
            log("ContinueEpoch cleared");
        } catch (e) {}
    }
    // Clear cohort-run guard state.
    function clearCohortGuard() {
        try {
            localStorage.removeItem("activityPlanState.cohortAdd.guard");
            localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
            log("CohortGuard and editDoneMap cleared");
        } catch (e) {}
    }

    // Clear the after-refresh action string.
    function clearAfterRefresh() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_AFTER_REFRESH);
        } catch (e) {}
        if (raw) {
            try {
                localStorage.removeItem(STORAGE_AFTER_REFRESH);
                log("AfterRefresh cleared");
            } catch (e2) {}
        }
    }

    // Clear consent scan index.
    function clearConsentScanIndex() {
        try {
            localStorage.removeItem(STORAGE_CONSENT_SCAN_INDEX);
        } catch (e) {}
    }
    // Clear last volunteer id storage.
    function clearLastVolunteerId() {
        try {
            localStorage.removeItem("activityPlanState.lastVolunteerId");
            log("lastVolunteerId cleared");
        } catch (e) {}
    }
    function clearAllRunState() {
        clearRunMode();
        clearContinueEpoch();
        clearCohortGuard();
        clearPendingIds();
        clearAfterRefresh();
        clearConsentScanIndex();
        clearLastVolunteerId();
        clearSelectedVolunteerIds();
        try {
            localStorage.removeItem(STORAGE_EDIT_STUDY);
            localStorage.removeItem(STORAGE_CHECK_ELIG_LOCK);
            localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
            log("Cleared cohortAdd.editDoneMap");
        } catch (e) {}
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

        var runningBox = document.createElement("div");
        runningBox.style.display = "none";
        runningBox.style.marginTop = "10px";
        runningBox.style.textAlign = "center";
        runningBox.style.fontSize = "16px";
        runningBox.style.color = "#fff";
        runningBox.textContent = "Running";
        container.appendChild(runningBox);

        var popup = createPopup({
            title: "Import I/E",
            description: "Add I/E to Eligibility Mapping",
            content: container,
            width: "450px",
            height: "auto",
            onClose: function () {
                // X pressed → STOP everything (but don't pause)
                log("ImportElig: popup X pressed → stopping automation");
                clearAllRunState();
                clearEligibilityWorkingState();
                // Clear pending popup flag
                try {
                    localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
                } catch (e) {}
            }
        });

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

            runningBox.style.display = "block";

            var dots = 1;
            var interval = setInterval(function () {
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
                runningBox.textContent = t;
            }, 350);

            onConfirm(function (message) {
                clearInterval(interval);
                if (message) {
                    runningBox.textContent = message;
                    setTimeout(function() {
                        popup.close();
                    }, 2000);
                } else {
                    popup.close();
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


        async function attemptCheckItemMatchForSA(code, expectedPlanVal, expectedSAVal) {
            log("ImportElig: attemptCheckItemMatchForSA start");
            var saEl = document.querySelector("select#scheduledActivity");
            if (!saEl) {
                log("ImportElig: attemptCheckItemMatchForSA no SA select");
                return false;
            }
            var curSA = (saEl.value + "");
            if (String(curSA) !== String(expectedSAVal)) {
                log("ImportElig: attemptCheckItemMatchForSA SA mismatch before scan");
                return false;
            }
            var elapsed = 0;
            var step = 160;
            var max = 3200;
            while (elapsed <= max) {
                saEl = document.querySelector("select#scheduledActivity");
                if (!saEl) {
                    await sleep(step);
                    elapsed = elapsed + step;
                    continue;
                }
                curSA = (saEl.value + "");
                if (String(curSA) !== String(expectedSAVal)) {
                    log("ImportElig: attemptCheckItemMatchForSA SA changed mid-scan");
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
                                    saEl = document.querySelector("select#scheduledActivity");
                                    if (!saEl) {
                                        log("ImportElig: attemptCheckItemMatchForSA SA missing before set");
                                        return false;
                                    }
                                    curSA = (saEl.value + "");
                                    if (String(curSA) !== String(expectedSAVal)) {
                                        log("ImportElig: attemptCheckItemMatchForSA SA changed before set");
                                        return false;
                                    }
                                    itemRefSel.value = val;
                                    var evt = new Event("change", { bubbles: true });
                                    itemRefSel.dispatchEvent(evt);
                                    log("ImportElig: Check Item matched '" + String(txt) + "'");
                                    setLastMatchSelection(expectedPlanVal, expectedSAVal, val);
                                    return true;
                                }
                            }
                        }
                        i = i + 1;
                    }
                    return false;
                }
                await sleep(step);
                elapsed = elapsed + step;
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



    function startImportEligibilityMapping() {
        log("ImportElig: startImportEligibilityMapping invoked");

        // Don't set run mode here - only set it when Confirm is pressed
        // This prevents auto-redirect if user navigates away before confirming

        var path = location.pathname;

        if (path !== "/secure/crfdesign/studylibrary/eligibility/list") {
            log("ImportElig: not on eligibility list page; redirecting");

            try {
                localStorage.setItem(STORAGE_ELIG_IMPORT_PENDING_POPUP, "1");
            } catch (e) {
            }

            location.href = ELIGIBILITY_LIST_URL;
            return;
        }

        try {
            localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
        } catch (e) {
        }

        // Set flag that popup is open (but not run mode yet - that's set when Confirm is pressed)
        try {
            localStorage.setItem(STORAGE_ELIG_IMPORT_PENDING_POPUP, "1");
        } catch (e) {
        }

        buildImportEligPopup(function (doneCallback) {
            setTimeout(async function () {
                await executeEligibilityMappingAutomation();
                doneCallback("No more Eligibility Item to Check");
                location.href = ELIGIBILITY_LIST_URL + "#";
            }, 400);
        });
    }



    async function executeEligibilityMappingAutomation() {
        if (isPaused()) {
            log("ImportElig: paused at start; aborting");
            return;
        }
        if (!isEligibilityListPage()) {
            log("ImportElig: not on list page; redirecting");
            try {
                localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ELIG_IMPORT);
            } catch (e) {
            }
            location.href = ELIGIBILITY_LIST_URL;
            return;
        }
        log("ImportElig: attempting unlock before data collection");
        var unlocked = await unlockEligibilityMapping();
        if (!unlocked) {
            log("ImportElig: unlock failed; stopping");
            try {
                localStorage.removeItem(STORAGE_RUN_MODE);
            } catch (e) {
            }
            return;
        }
        log("ImportElig: unlock successful; waiting 1s for table");
        await sleep(1000);
        var collected = await collectEligibilityTableMap();
        var existingSet = collected.codeSet;
        var importedSet = getImportedItemsSet();
        var ignoreKeywords = getIgnoreKeywords();

        function shouldIgnoreCode(code) {
            if (!code || code.length === 0) {
                return false;
            }
            var codeLower = (code + "").toLowerCase();
            var k = 0;
            while (k < ignoreKeywords.length) {
                var keyword = (ignoreKeywords[k] + "").toLowerCase().trim();
                if (keyword.length > 0) {
                    if (codeLower.indexOf(keyword) >= 0) {
                        log("ImportElig: ignoring code '" + String(code) + "' (matches ignore keyword '" + String(ignoreKeywords[k]) + "')");
                        return true;
                    }
                }
                k = k + 1;
            }
            return false;
        }

        var guardIterations = 0;
        var maxIterations = 200;
        while (guardIterations < maxIterations) {
            guardIterations = guardIterations + 1;
            if (isPaused()) {
                log("ImportElig: paused; stopping loop");
                break;
            }
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                break;
            }
            var opened = await openAddEligibilityModal();
            if (!opened) {
                log("ImportElig: cannot open modal; stopping");
                break;
            }
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                break;
            }
            var eligList = await readEligibilityItemCodesFromSelect();
            var allCodes = eligList.codes;
            var codeToVal = eligList.valueMap;
            var pick = "";
            var j = 0;
            while (j < allCodes.length) {
                var c = allCodes[j];
                if (shouldIgnoreCode(c)) {
                    j = j + 1;
                    continue;
                }
                var inExisting = existingSet.has(c);
                var inImported = importedSet.has(c);
                if (!inExisting && !inImported) {
                    pick = c;
                    break;
                }
                j = j + 1;
            }
            if (!pick || pick.length === 0) {
                log("ImportElig: no new items; finishing");
                try {
                    localStorage.removeItem(STORAGE_RUN_MODE);
                } catch (e) {
                }
                clearImportedItemsSet();
                // Note: The completion message will be shown by the callback in startImportEligibilityMapping
                return;
            }
            var selOk = await selectEligibilityItemByCode(pick, codeToVal);
            if (!selOk) {
                log("ImportElig: select failed; closing modal");
                var closeBtn = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn) {
                    closeBtn.click();
                    log("ImportElig: modal closed");
                }
                await sleep(500);
                continue;
            }
            var sexOk = await setSexOptionFromEligibilitySelection(pick, codeToVal);
            if (!sexOk) {
                log("ImportElig: sexOption default");
            }
            var foundCheckItem = await trySelectCheckItemForCodeThroughAllPlansAndActivities(pick);
            if (!foundCheckItem) {
                log("ImportElig: no CheckItem match found for code '" + String(pick) + "' after scanning all plans and activities");
                addToIgnoreKeywords(pick);
                // Reload ignore keywords so the updated list is used in subsequent iterations
                ignoreKeywords = getIgnoreKeywords();
                var closeBtn2 = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn2) {
                    closeBtn2.click();
                    log("ImportElig: modal closed (no match)");
                }
                await sleep(500);
                importedSet.add(pick);
                persistImportedItemsSet(importedSet);
                clearLastMatchSelection();
                continue;
            }
            log("ImportElig: stabilizing selections before comparator");
            await stabilizeSelectionBeforeComparator(3000);
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                break;
            }
            log("ImportElig: waiting for comparator to appear after Check Item selection");
            var compReady = await waitForComparatorReady(1000);
            if (!compReady) {
                log("ImportElig: comparator never appeared; skipping");
                var closeBtn2b = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn2b) {
                    closeBtn2b.click();
                }
                await sleep(500);
                clearLastMatchSelection();
                continue;
            }
            var cmpOk = await setComparatorEQ();
            if (!cmpOk) {
                log("ImportElig: comparator fail; skipping");
                var closeBtn3 = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn3) {
                    closeBtn3.click();
                }
                await sleep(500);
                clearLastMatchSelection();
                continue;
            }
            var sfOk = await selectCodeListValueContainingSF();
            if (!sfOk) {
                log("ImportElig: SF fail; skipping");
                var closeBtn4 = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn4) {
                    closeBtn4.click();
                }
                await sleep(500);
                clearLastMatchSelection();
                continue;
            }
            var saved = await clickSaveAndWait();
            if (!saved) {
                log("ImportElig: save failed; stopping");
                break;
            }
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                break;
            }
            existingSet.add(pick);
            importedSet.add(pick);
            persistImportedItemsSet(importedSet);
            log("ImportElig: added " + String(pick));
            clearLastMatchSelection();
            await sleep(1000);
        }
        log("ImportElig: loop end; clearing run mode");
        try {
            localStorage.removeItem(STORAGE_RUN_MODE);
        } catch (e) {
        }
    }

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

        var popup = createPopup({ title: AE_POPUP_TITLE, description: "Auto-navigate to AE data page for subject", content: container, width: "380px", height: "auto" });

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
            description: "Find and pull barcode in background…",
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
        popup.style.borderRadius = "8px";
        popup.style.padding = "0";
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

        popup.close = function () {
            document.removeEventListener("keydown", keyHandler, true);
            popup.remove();
        };

        return popup;
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
        runBarcodeBtn.style.background = "#4a90e2";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        runBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        runBarcodeBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
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
        saBuilderBtn.style.background = "#4a90e2";
        saBuilderBtn.style.color = "#fff";
        saBuilderBtn.style.border = "none";
        saBuilderBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        saBuilderBtn.style.padding = scale(BUTTON_PADDING_PX);
        saBuilderBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        saBuilderBtn.style.cursor = "pointer";
        saBuilderBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        saBuilderBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

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
        importEligBtn.style.background = "#4a90e2";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        importEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        importEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        importEligBtn.style.cursor = "pointer";
        importEligBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        importEligBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(findAeBtn);
        btnRow.appendChild(findFormBtn);
        btnRow.appendChild(findStudyEventsBtn);
        btnRow.appendChild(importEligBtn);
        btnRow.appendChild(parseMethodBtn);
        btnRow.appendChild(openEligBtn);
        btnRow.appendChild(subjectEligBtn);
        btnRow.appendChild(saBuilderBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(toggleLogsBtn);
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