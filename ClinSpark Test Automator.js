
// ==UserScript==
// @name ClinSpark Test Automator
// @namespace vinh.activity.plan.state
// @version 3.2.8
// @description Run Activity Plans, Study Update (Cancel if already Active), Cohort Add, Informed Consent; draggable panel; Run ALL pipeline; Pause/Resume; Extensible buttons API;
// @match https://cenexeltest.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Test%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Test%20Automator.js
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// @ts-check
// ==/UserScript==

(function () {
    var STORAGE_PANEL_WIDTH = "activityPlanState.panel.width";
    var STORAGE_PANEL_HEIGHT = "activityPlanState.panel.height";
    // UI Scale Constants
    var UI_SCALE = 1.0; // Master scale factor (will be initialized after function definitions)
    var PANEL_DEFAULT_WIDTH = 340;
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 50;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 60;
    var PANEL_PADDING_PX = 12;
    var PANEL_BORDER_RADIUS_PX = 8;
    var PANEL_FONT_SIZE_PX = 14;
    var HEADER_FONT_WEIGHT = "600";
    var HEADER_TITLE_OFFSET_PX = 16;
    var HEADER_LEFT_SPACER_WIDTH_PX = 16;
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

    var STORAGE_KEY = "activityPlanState.run";
    var STORAGE_PENDING = "activityPlanState.pendingIds";
    var STORAGE_AFTER_REFRESH = "activityPlanState.afterRefresh";
    var STORAGE_EDIT_STUDY = "activityPlanState.editStudy";
    var STORAGE_RUN_MODE = "activityPlanState.runMode";
    var STORAGE_CONTINUE_EPOCH = "activityPlanState.continueEpoch";
    var LIST_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/activityplans/list";
    var STUDY_SHOW_URL = "https://cenexeltest.clinspark.com/secure/administration/studies/show";
    var PANEL_ID = "activityPlanStatePanel";
    var LOG_ID = "activityPlanStateLog";
    var STORAGE_SELECTED_IDS = "activityPlanState.selectedVolunteerIds";
    var STORAGE_IC_BARCODE = "activityPlanState.ic.barcode";
    var STORAGE_CONSENT_SCAN_INDEX = "activityPlanState.consent.scanIndex";
    var STORAGE_PAUSED = "activityPlanState.paused";
    var STORAGE_CHECK_ELIG_LOCK = "activityPlanState.checkEligLock";
    var btnRowRef = null;
    var PENDING_BUTTONS = [];
    var STUDY_METADATA_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/show/studymetadata";
    var STORAGE_BARCODE_SUBJECT_TEXT = "activityPlanState.barcode.subjectText";
    var STORAGE_BARCODE_SUBJECT_ID = "activityPlanState.barcode.subjectId";
    var STORAGE_BARCODE_RESULT = "activityPlanState.barcode.result";
    var STORAGE_PANEL_COLLAPSED = "activityPlanState.panel.collapsed";
    var FORM_DELAY_MS = 800;
    var BARCODE_START_TS = 0;
    var DELAY_V2_ITEM_MS = 100;
    var DELAY_V2_GROUP_RESCAN_MS = 1000;
    var RUN_FORM_V2_START_TS = 0;
    var STORAGE_FORM_VALUE_MODE = "activityPlanState.formValueMode";
    var STORAGE_IMPORT_DONE_MAP = "activityPlanState.import.doneMap";
    var STORAGE_IMPORT_IN_PROGRESS = "activityPlanState.import.inprogress";
    var STORAGE_NON_SCRN_EPOCH_INDEX = "activityPlanState.nonscrn.epochIndex";
    var STORAGE_IMPORT_SUBJECT_IDS = "activityPlanState.import.subjectIds";

    var STORAGE_NON_SCRN_SELECTED_EPOCH = "activityPlanState.nonscrn.selectedEpoch";
    var STORAGE_IMPORT_COHORT_EDIT_DONE = "activityPlanState.import.cohortEditDone";
    var STORAGE_LOG_VISIBLE = "activityPlanState.log.visible";
    var STORAGE_MANUAL_SELECT_INITIAL_REF_TIME = "activityPlanState.manualSelectInitialRefTime";
    var STORAGE_RUN_LOCK_SAMPLE_PATHS = "activityPlanState.runLockSamplePaths";
    var STORAGE_SAMPLE_PATH_AUTO_CLOSE = "activityPlanState.samplePath.autoClose";
    var STORAGE_LOCK_SAMPLE_PATHS_POPUP = "activityPlanState.lockSamplePaths.popup";
    var LOCK_SAMPLE_PATHS_POPUP_REF = null;
    var STORAGE_LOCK_ACTIVITY_PLANS_POPUP = "activityPlanState.lockActivityPlans.popup";
    var LOCK_ACTIVITY_PLANS_POPUP_REF = null;

    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const STORAGE_PANEL_HOTKEY = "activityPlanState.panel.hotkey";
    const PANEL_TOGGLE_KEY = "F2";
    const ELIGIBILITY_LIST_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";

    // Run Find Study Event

    var STORAGE_FIND_FORM_PENDING = "activityPlanState.findForm.pending";
    var STORAGE_FIND_FORM_KEYWORD = "activityPlanState.findForm.keyword";
    var STORAGE_FIND_FORM_SUBJECT = "activityPlanState.findForm.subject";
    var STORAGE_FIND_FORM_STATUS_VALUES = "activityPlanState.findForm.statusValues";

    var STORAGE_FIND_STUDY_EVENT_PENDING = "activityPlanState.findStudyEvent.pending";
    var STORAGE_FIND_STUDY_EVENT_KEYWORD = "activityPlanState.findStudyEvent.keyword";
    var STORAGE_FIND_STUDY_EVENT_SUBJECT = "activityPlanState.findStudyEvent.subject";
    var STORAGE_FIND_STUDY_EVENT_STATUS_VALUES = "activityPlanState.findStudyEvent.statusValues";

    var STUDY_EVENT_POPUP_TITLE = "Find Study Events";
    var STUDY_EVENT_POPUP_DESCRIPTION = "Auto-navigate to Study Events data page based on keywords and status";
    var STUDY_EVENT_POPUP_KEYWORD_LABEL = "Study Event Keyword";
    var STUDY_EVENT_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var STUDY_EVENT_POPUP_OK_TEXT = "Continue";
    var STUDY_EVENT_POPUP_CANCEL_TEXT = "Cancel";

    var STUDY_EVENT_NO_MATCH_TITLE = "Find Study Events";
    var STUDY_EVENT_NO_MATCH_MESSAGE = "No study event is found.";

    // Run Find Form
    var FORM_LIST_URL = "https://cenexeltest.clinspark.com/secure/study/data/list";
    var FORM_POPUP_TITLE = "Find Form";
    var FORM_POPUP_DESCRIPTION = "Auto-navigate to Form data page based on keywords and status";
    var FORM_POPUP_KEYWORD_LABEL = "Form Keyword";
    var FORM_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var FORM_POPUP_OK_TEXT = "Continue";
    var FORM_POPUP_CANCEL_TEXT = "Cancel";

    var FORM_NO_MATCH_TITLE = "Find Form";
    var FORM_NO_MATCH_MESSAGE = "No form is found.";

    var BARCODE_BG_TAB = null;
    const RUNMODE_CLEAR_MAPPING = "clearMapping";

    // Run Parse Method
    var PARSE_METHOD_POPUP_TITLE = "Parse Method";
    var PARSE_METHOD_POPUP_DESCRIPTION = "Auto-navigate to data page with forms that has methods that pulls this item.";
    var STORAGE_PARSE_METHOD_RUNNING = "activityPlanState.parseMethod.running";
    var STORAGE_PARSE_METHOD_ITEM_NAME = "activityPlanState.parseMethod.itemName";
    var STORAGE_PARSE_METHOD_RESULTS = "activityPlanState.parseMethod.results";
    var STORAGE_PARSE_METHOD_COMPLETED = "activityPlanState.parseMethod.completed";
    var RUNMODE_PARSE_METHOD = "parseMethod";
    var METHOD_LIBRARY_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/method";
    var PARSE_METHOD_CANCELED = false;
    var PARSE_METHOD_COLLECTED_METHODS = [];
    var PARSE_METHOD_COLLECTED_FORMS = [];

    // Run Data Collection
    const DATA_COLLECTION_SUBJECT_URL = "https://cenexeltest.clinspark.com/secure/datacollection/subject";
    var COLLECT_ALL_CANCELLED = false;
    var COLLECT_ALL_POPUP_REF = null;
    var RUN_ALL_POPUP_REF = null;
    var CLEAR_MAPPING_POPUP_REF = null;
    var IMPORT_ELIG_POPUP_REF = null;
    var IMPORT_COHORT_POPUP_REF = null;
    var ADD_COHORT_POPUP_REF = null;
    var ICF_BARCODE_POPUP_REF = null;
    var RUN_ALL_POPUP_TITLE = "Run All";
    var RUN_ALL_POPUP_DESCRIPTION = "Auto-navigate to Run All data page based on keywords and status";
    const STORAGE_RUN_ALL_POPUP = "activityPlanState.runAllPopup";
    const STORAGE_ICF_BARCODE_POPUP = "activityPlanState.icfBarcodePopup";
    const STORAGE_RUN_ALL_STATUS = "activityPlanState.runAllStatus";
    const STORAGE_CLEAR_MAPPING_POPUP = "activityPlanState.clearMappingPopup";
    const STORAGE_IMPORT_ELIG_POPUP = "activityPlanState.importEligPopup";
    const STORAGE_IMPORT_COHORT_POPUP = "activityPlanState.importCohortPopup";
    const STORAGE_ADD_COHORT_POPUP = "activityPlanState.addCohortPopup";

    // Run Subject Eligibility
    var ELIGIBILITY_POPUP_TITLE = "Run Subject Eligibility";
    var ELIGIBILITY_POPUP_DESCRIPTION = "Auto-navigate to Run Subject Eligibility data page based on keywords and status";
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

    var CLEAR_MAPPING_CANCELED = false;
    var CLEAR_MAPPING_PAUSE = false;


    //==========================
    // ARCHIVE/UPDATE FORMS FEATURE
    //==========================
    // This feature allows batch replacement of forms across all occurrences in the SA table.
    // Users select a source form and a target form; the source is archived everywhere
    // and the target is added in its place, preserving properties and visibility.
    //==========================
    // Archive/Update Forms Feature
    var STORAGE_ARCHIVE_UPDATE_FORMS_CANCELLED = "activityPlanState.archiveUpdateForms.cancelled";
    var ARCHIVE_UPDATE_FORMS_CANCELLED = false;
    var ARCHIVE_UPDATE_FORMS_POPUP_REF = null;
    var ARCHIVE_UPDATE_FORMS_PROGRESS_REF = null;

    function ArchiveUpdateFormsFunctions() {}

    // Normalize text specifically for visibility condition matching
    function normalizeVisibilityText(t) {
        if (typeof t !== "string") {
            return "";
        }
        var s = t.toLowerCase();
        
        // Remove time offsets like -00:05:00, +01:30:00, etc.
        s = s.replace(/[\+\-]\d{2}:\d{2}:\d{2}/g, "");
        
        // Remove trailing occurrence numbers like (1), (2), etc.
        s = s.replace(/\s*\(\d+\)\s*$/g, "");
        
        // Remove > symbol (used in collected data but not in options)
        s = s.replace(/>/g, "");
        
        // Normalize spaces around = sign
        s = s.replace(/\s*=\s*/g, "=");
        
        // Collapse multiple spaces
        s = s.replace(/\s+/g, " ");
        
        // Remove all spaces for final comparison
        s = s.replace(/\s/g, "");
        
        // Remove common punctuation
        s = s.replace(/[\-\_\(\)\[\]]/g, "");
        
        return s.trim();
    }

    // Check if on correct Activity Plan Show page
    function isOnArchiveUpdateFormsPage() {
        var pattern = /^https:\/\/cenexeltest\.clinspark\.com\/secure\/crfdesign\/activityplans\/show\/\d+$/;
        return pattern.test(location.href.split("?")[0].split("#")[0]);
    }

    // Scan SA table and collect all rows with full details
    function scanSATableForArchiveUpdate() {
        var rows = [];
        var tbody = document.getElementById("saTableBody");
        if (!tbody) {
            log("Archive/Update Forms: saTableBody not found");
            return rows;
        }
        var trs = tbody.querySelectorAll("tr");
        log("Archive/Update Forms: found " + trs.length + " rows");
        for (var i = 0; i < trs.length; i++) {
            var tr = trs[i];
            var cells = tr.querySelectorAll("td");
            if (cells.length < 4) continue;

            // Extract segment
            var segmentCell = cells[1];
            var segmentText = normalizeSAText(segmentCell.textContent);
            var segmentValue = "";
            var segmentSelect = segmentCell.querySelector("select");
            if (segmentSelect) {
                segmentValue = segmentSelect.value || "";
            }
            // Try to get value from data attributes or links
            if (!segmentValue) {
                var segLink = segmentCell.querySelector("a[href]");
                if (segLink) {
                    var segHref = segLink.getAttribute("href") || "";
                    var segMatch = segHref.match(/\/(\d+)$/);
                    if (segMatch) segmentValue = segMatch[1];
                }
            }

            // Extract study event
            var eventCell = cells[2];
            var eventText = normalizeSAText(eventCell.textContent);
            var eventValue = "";
            var eventSelect = eventCell.querySelector("select");
            if (eventSelect) {
                eventValue = eventSelect.value || "";
            }
            if (!eventValue) {
                var evLink = eventCell.querySelector("a[href]");
                if (evLink) {
                    var evHref = evLink.getAttribute("href") || "";
                    var evMatch = evHref.match(/\/(\d+)$/);
                    if (evMatch) eventValue = evMatch[1];
                }
            }

            // Extract form
            var formCell = cells[3];
            var formText = normalizeSAText(formCell.textContent);
            var formValue = "";
            var formSelect = formCell.querySelector("select");
            if (formSelect) {
                formValue = formSelect.value || "";
            }
            if (!formValue) {
                var formLink = formCell.querySelector("a[href]");
                if (formLink) {
                    var formHref = formLink.getAttribute("href") || "";
                    var formMatch = formHref.match(/\/show\/form\/(\d+)/);
                    if (formMatch) formValue = formMatch[1];
                }
            }

            // Get the row's action cell for archive/edit links
            var actionCell = cells[cells.length - 1] || cells[0];
            var archiveLink = tr.querySelector('a[href*="archivescheduledactivity"]');
            var editLink = tr.querySelector('a[href*="editscheduledactivity"]');
            var visibilityLink = tr.querySelector('a[href*="visibility"]');

            // Check if already archived (Un-Archive link present)
            var isArchived = false;
            if (archiveLink) {
                var archiveLinkText = (archiveLink.textContent || "").toLowerCase();
                if (archiveLinkText.indexOf("un-archive") !== -1) {
                    isArchived = true;
                }
            }

            var formKey = formValue || normalizeText(formText);
            var rowKey = (segmentValue || normalizeText(segmentText)) + "|" +
                (eventValue || normalizeText(eventText)) + "|" + formKey;

            rows.push({
                rowElement: tr,
                segmentText: segmentText,
                segmentValue: segmentValue,
                eventText: eventText,
                eventValue: eventValue,
                formText: formText,
                formValue: formValue,
                formKey: formKey,
                rowKey: rowKey,
                archiveLink: archiveLink,
                editLink: editLink,
                visibilityLink: visibilityLink,
                isArchived: isArchived
            });
        }
        return rows;
    }

    // Build unique forms map from rows
    function buildUniqueFormsMap(rows) {
        var formsMap = {};
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var key = row.formKey;
            if (!formsMap[key]) {
                formsMap[key] = {
                    formKey: key,
                    formText: row.formText,
                    formValue: row.formValue,
                    occurrences: []
                };
            }
            formsMap[key].occurrences.push(row);
        }
        return formsMap;
    }

    // Create the Archive/Update Forms selection GUI
    // sourceFormsMap: forms from table with occurrences (for source panel)
    // targetFormsArray: forms from Add modal dropdown [{ value, text }, ...]
    function createArchiveUpdateFormsGUI(sourceFormsMap, targetFormsArray) {
        // Build source forms array from map
        var sourceFormsArray = [];
        for (var key in sourceFormsMap) {
            if (sourceFormsMap.hasOwnProperty(key)) {
                sourceFormsArray.push(sourceFormsMap[key]);
            }
        }
        sourceFormsArray.sort(function(a, b) {
            return a.formText.localeCompare(b.formText);
        });

        // Sort target forms array
        var targetFormsSorted = targetFormsArray.slice().sort(function(a, b) {
            return a.text.localeCompare(b.text);
        });

        var selectedSource = null;
        var selectedTarget = null;

        var container = document.createElement("div");
        container.style.cssText = "display:flex;flex-direction:column;height:100%;min-height:500px;gap:16px;";

        // Header row with search bars
        var headerRow = document.createElement("div");
        headerRow.style.cssText = "display:grid;grid-template-columns:1fr 1fr;gap:16px;";

        var sourceSearchContainer = document.createElement("div");
        sourceSearchContainer.style.cssText = "display:flex;flex-direction:column;gap:8px;";
        var sourceLabel = document.createElement("div");
        sourceLabel.textContent = "Source Forms (to be replaced)";
        sourceLabel.style.cssText = "font-weight:600;font-size:14px;color:#ff6b6b;";
        var sourceSearch = document.createElement("input");
        sourceSearch.type = "text";
        sourceSearch.placeholder = "Search source forms...";
        sourceSearch.style.cssText = "padding:10px;border-radius:6px;border:1px solid #444;background:#222;color:#fff;font-size:14px;";
        sourceSearchContainer.appendChild(sourceLabel);
        sourceSearchContainer.appendChild(sourceSearch);

        var targetSearchContainer = document.createElement("div");
        targetSearchContainer.style.cssText = "display:flex;flex-direction:column;gap:8px;";
        var targetLabel = document.createElement("div");
        targetLabel.textContent = "Target Forms (replacements)";
        targetLabel.style.cssText = "font-weight:600;font-size:14px;color:#51cf66;";
        var targetSearch = document.createElement("input");
        targetSearch.type = "text";
        targetSearch.placeholder = "Search target forms...";
        targetSearch.style.cssText = "padding:10px;border-radius:6px;border:1px solid #444;background:#222;color:#fff;font-size:14px;";
        targetSearchContainer.appendChild(targetLabel);
        targetSearchContainer.appendChild(targetSearch);

        headerRow.appendChild(sourceSearchContainer);
        headerRow.appendChild(targetSearchContainer);
        container.appendChild(headerRow);

        // Panels row
        var panelsRow = document.createElement("div");
        panelsRow.style.cssText = "display:grid;grid-template-columns:1fr 1fr;gap:16px;flex:1;overflow:hidden;";

        var sourcePanel = document.createElement("div");
        sourcePanel.style.cssText = "border:2px solid #ff6b6b;border-radius:8px;background:#1a1a1a;overflow-y:auto;padding:8px;";
        sourcePanel.id = "archiveUpdateSourcePanel";

        var targetPanel = document.createElement("div");
        targetPanel.style.cssText = "border:2px solid #51cf66;border-radius:8px;background:#1a1a1a;overflow-y:auto;padding:8px;";
        targetPanel.id = "archiveUpdateTargetPanel";

        panelsRow.appendChild(sourcePanel);
        panelsRow.appendChild(targetPanel);
        container.appendChild(panelsRow);

        // Error message area
        var errorDiv = document.createElement("div");
        errorDiv.style.cssText = "color:#ff6b6b;text-align:center;font-size:14px;min-height:20px;";
        errorDiv.id = "archiveUpdateError";
        container.appendChild(errorDiv);

        // Reason inputs section
        var reasonsContainer = document.createElement("div");
        reasonsContainer.style.cssText = "display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:12px 0;";

        // Archive Reason
        var archiveReasonContainer = document.createElement("div");
        archiveReasonContainer.style.cssText = "display:flex;flex-direction:column;gap:6px;";
        var archiveReasonLabel = document.createElement("label");
        archiveReasonLabel.textContent = "Archive Reason";
        archiveReasonLabel.style.cssText = "font-size:13px;font-weight:600;color:#ccc;";
        var archiveReasonInput = document.createElement("input");
        archiveReasonInput.type = "text";
        archiveReasonInput.value = "Old version";
        archiveReasonInput.id = "archiveUpdateArchiveReason";
        archiveReasonInput.style.cssText = "padding:10px;border-radius:6px;border:1px solid #444;background:#222;color:#fff;font-size:13px;";
        archiveReasonContainer.appendChild(archiveReasonLabel);
        archiveReasonContainer.appendChild(archiveReasonInput);

        // Visibility Reason
        var visibilityReasonContainer = document.createElement("div");
        visibilityReasonContainer.style.cssText = "display:flex;flex-direction:column;gap:6px;";
        var visibilityReasonLabel = document.createElement("label");
        visibilityReasonLabel.textContent = "Visibility Reason";
        visibilityReasonLabel.style.cssText = "font-size:13px;font-weight:600;color:#ccc;";
        var visibilityReasonInput = document.createElement("input");
        visibilityReasonInput.type = "text";
        visibilityReasonInput.value = "Add visibility condition";
        visibilityReasonInput.id = "archiveUpdateVisibilityReason";
        visibilityReasonInput.style.cssText = "padding:10px;border-radius:6px;border:1px solid #444;background:#222;color:#fff;font-size:13px;";
        visibilityReasonContainer.appendChild(visibilityReasonLabel);
        visibilityReasonContainer.appendChild(visibilityReasonInput);

        reasonsContainer.appendChild(archiveReasonContainer);
        reasonsContainer.appendChild(visibilityReasonContainer);
        container.appendChild(reasonsContainer);

        // Button row
        var buttonRow = document.createElement("div");
        buttonRow.style.cssText = "display:flex;justify-content:center;gap:16px;padding-top:8px;";

        var clearAllBtn = document.createElement("button");
        clearAllBtn.textContent = "Clear All";
        clearAllBtn.style.cssText = "padding:12px 24px;border-radius:6px;border:none;background:#6c757d;color:#fff;font-size:14px;font-weight:500;cursor:pointer;transition:background 0.2s;";
        clearAllBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        clearAllBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.disabled = true;
        confirmBtn.style.cssText = "padding:12px 24px;border-radius:6px;border:none;background:#28a745;color:#fff;font-size:14px;font-weight:500;cursor:pointer;transition:background 0.2s;opacity:0.5;";

        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Cancel";
        cancelBtn.style.cssText = "padding:12px 24px;border-radius:6px;border:none;background:#dc3545;color:#fff;font-size:14px;font-weight:500;cursor:pointer;transition:background 0.2s;";
        cancelBtn.onmouseenter = function() { this.style.background = "#c82333"; };
        cancelBtn.onmouseleave = function() { this.style.background = "#dc3545"; };

        buttonRow.appendChild(clearAllBtn);
        buttonRow.appendChild(confirmBtn);
        buttonRow.appendChild(cancelBtn);
        container.appendChild(buttonRow);

        // Render forms in panels
        function renderSourcePanel(panel, searchValue) {
            panel.innerHTML = "";
            var filter = (searchValue || "").toLowerCase();
            for (var i = 0; i < sourceFormsArray.length; i++) {
                var form = sourceFormsArray[i];
                if (filter && form.formText.toLowerCase().indexOf(filter) === -1) continue;

                var item = document.createElement("div");
                item.style.cssText = "display:flex;align-items:center;gap:10px;padding:10px;margin-bottom:6px;border:1px solid #333;border-radius:6px;background:#252525;cursor:pointer;transition:all 0.2s;";
                item.dataset.formKey = form.formKey;

                var radio = document.createElement("div");
                radio.style.cssText = "width:18px;height:18px;border:2px solid #666;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;";
                radio.className = "radio-indicator";

                var label = document.createElement("div");
                label.style.cssText = "flex:1;font-size:13px;";
                label.textContent = form.formText;

                var count = document.createElement("div");
                count.style.cssText = "font-size:11px;color:#888;background:#333;padding:2px 8px;border-radius:10px;";
                count.textContent = form.occurrences.length + " occurrence" + (form.occurrences.length !== 1 ? "s" : "");

                item.appendChild(radio);
                item.appendChild(label);
                item.appendChild(count);

                if (selectedSource === form.formKey) {
                    item.style.background = "#3d2020";
                    item.style.border = "1px solid #ff6b6b";
                    radio.innerHTML = '<div style="width:10px;height:10px;background:#ff6b6b;border-radius:50%;"></div>';
                }

                item.addEventListener("mouseenter", function() {
                    if (this.dataset.formKey !== selectedSource) {
                        this.style.background = "#333";
                    }
                });
                item.addEventListener("mouseleave", function() {
                    if (this.dataset.formKey !== selectedSource) {
                        this.style.background = "#252525";
                        this.style.border = "1px solid #333";
                    }
                });

                (function(formKey, itemEl) {
                    itemEl.addEventListener("click", function() {
                        selectedSource = formKey;
                        renderSourcePanel(sourcePanel, sourceSearch.value);
                        renderTargetPanel(targetPanel, targetSearch.value);
                        validateSelection();
                    });
                })(form.formKey, item);

                panel.appendChild(item);
            }
        }

        // Render target forms panel (forms from modal dropdown)
        function renderTargetPanel(panel, searchValue) {
            panel.innerHTML = "";
            var filter = (searchValue || "").toLowerCase();
            for (var i = 0; i < targetFormsSorted.length; i++) {
                var form = targetFormsSorted[i];
                if (filter && form.text.toLowerCase().indexOf(filter) === -1) continue;

                var item = document.createElement("div");
                item.style.cssText = "display:flex;align-items:center;gap:10px;padding:10px;margin-bottom:6px;border:1px solid #333;border-radius:6px;background:#252525;cursor:pointer;transition:all 0.2s;";
                item.dataset.formKey = form.value;

                var radio = document.createElement("div");
                radio.style.cssText = "width:18px;height:18px;border:2px solid #666;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;";
                radio.className = "radio-indicator";

                var label = document.createElement("div");
                label.style.cssText = "flex:1;font-size:13px;";
                label.textContent = form.text;

                item.appendChild(radio);
                item.appendChild(label);

                if (selectedTarget === form.value) {
                    item.style.background = "#1a3d1a";
                    item.style.border = "1px solid #51cf66";
                    radio.innerHTML = '<div style="width:10px;height:10px;background:#51cf66;border-radius:50%;"></div>';
                }

                item.addEventListener("mouseenter", function() {
                    if (this.dataset.formKey !== selectedTarget) {
                        this.style.background = "#333";
                    }
                });
                item.addEventListener("mouseleave", function() {
                    if (this.dataset.formKey !== selectedTarget) {
                        this.style.background = "#252525";
                        this.style.border = "1px solid #333";
                    }
                });

                (function(formValue, itemEl) {
                    itemEl.addEventListener("click", function() {
                        selectedTarget = formValue;
                        renderSourcePanel(sourcePanel, sourceSearch.value);
                        renderTargetPanel(targetPanel, targetSearch.value);
                        validateSelection();
                    });
                })(form.value, item);

                panel.appendChild(item);
            }
        }

        function validateSelection() {
            errorDiv.textContent = "";
            var valid = false;
            if (selectedSource && selectedTarget) {
                if (selectedSource === selectedTarget) {
                    errorDiv.textContent = "Source and Target cannot be the same form.";
                } else {
                    valid = true;
                }
            }
            confirmBtn.disabled = !valid;
            confirmBtn.style.opacity = valid ? "1" : "0.5";
            if (valid) {
                confirmBtn.onmouseenter = function() { this.style.background = "#218838"; };
                confirmBtn.onmouseleave = function() { this.style.background = "#28a745"; };
            } else {
                confirmBtn.onmouseenter = null;
                confirmBtn.onmouseleave = null;
            }
        }

        // Debounce for search
        var sourceTimeout = null;
        var targetTimeout = null;

        sourceSearch.addEventListener("input", function() {
            clearTimeout(sourceTimeout);
            sourceTimeout = setTimeout(function() {
                renderSourcePanel(sourcePanel, sourceSearch.value);
            }, 200);
        });

        targetSearch.addEventListener("input", function() {
            clearTimeout(targetTimeout);
            targetTimeout = setTimeout(function() {
                renderTargetPanel(targetPanel, targetSearch.value);
            }, 200);
        });

        clearAllBtn.addEventListener("click", function() {
            selectedSource = null;
            selectedTarget = null;
            sourceSearch.value = "";
            targetSearch.value = "";
            renderSourcePanel(sourcePanel, "");
            renderTargetPanel(targetPanel, "");
            validateSelection();
        });
        cancelBtn.addEventListener("click", function() {
            if (ARCHIVE_UPDATE_FORMS_POPUP_REF) {
                ARCHIVE_UPDATE_FORMS_POPUP_REF.close();
                ARCHIVE_UPDATE_FORMS_POPUP_REF = null;
            }
        });

        // Initial render
        renderSourcePanel(sourcePanel, "");
        renderTargetPanel(targetPanel, "");

        // Return container with confirm handler
        container.getSelection = function() {
            return { 
                source: selectedSource, 
                target: selectedTarget,
                archiveReason: archiveReasonInput.value,
                visibilityReason: visibilityReasonInput.value
            };
        };
        container.confirmBtn = confirmBtn;
        container.sourceFormsMap = sourceFormsMap;
        container.targetFormsArray = targetFormsArray;

        return container;
    }

    // Create progress popup for Archive/Update Forms
    function createArchiveUpdateProgressPopup(occurrences, targetFormText) {
        var container = document.createElement("div");
        container.style.cssText = "display:flex;flex-direction:column;height:100%;min-height:400px;";

        // Status header
        var statusDiv = document.createElement("div");
        statusDiv.id = "archiveUpdateStatus";
        statusDiv.style.cssText = "text-align:center;font-size:16px;font-weight:600;padding:12px;background:#222;border-radius:6px;margin-bottom:12px;";
        statusDiv.textContent = "Preparing... 0 of " + occurrences.length;
        container.appendChild(statusDiv);

        // Progress bar
        var progressBarOuter = document.createElement("div");
        progressBarOuter.style.cssText = "width:100%;height:8px;background:#333;border-radius:4px;margin-bottom:16px;overflow:hidden;";
        var progressBarInner = document.createElement("div");
        progressBarInner.id = "archiveUpdateProgressBar";
        progressBarInner.style.cssText = "width:0%;height:100%;background:#28a745;transition:width 0.3s;";
        progressBarOuter.appendChild(progressBarInner);
        container.appendChild(progressBarOuter);

        // Items list
        var listContainer = document.createElement("div");
        listContainer.style.cssText = "flex:1;overflow-y:auto;border:1px solid #333;border-radius:6px;background:#1a1a1a;padding:8px;";
        listContainer.id = "archiveUpdateItemsList";

        for (var i = 0; i < occurrences.length; i++) {
            var occ = occurrences[i];
            var item = document.createElement("div");
            item.style.cssText = "display:flex;align-items:center;justify-content:space-between;padding:8px 12px;margin-bottom:4px;border:1px solid #333;border-radius:4px;background:#252525;font-size:12px;";
            item.id = "archiveUpdateItem_" + i;

            var itemLabel = document.createElement("div");
            itemLabel.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";
            itemLabel.textContent = occ.segmentText + " → " + occ.eventText + " → " + occ.formText;

            var itemStatus = document.createElement("div");
            itemStatus.style.cssText = "padding:2px 8px;border-radius:10px;font-size:11px;background:#444;color:#888;margin-left:8px;flex-shrink:0;";
            itemStatus.textContent = "Pending";
            itemStatus.className = "item-status";

            item.appendChild(itemLabel);
            item.appendChild(itemStatus);
            listContainer.appendChild(item);
        }
        container.appendChild(listContainer);

        // Summary area
        var summaryDiv = document.createElement("div");
        summaryDiv.id = "archiveUpdateSummary";
        summaryDiv.style.cssText = "display:none;margin-top:12px;padding:12px;background:#222;border-radius:6px;text-align:center;";
        container.appendChild(summaryDiv);

        // Cancel button
        var cancelRow = document.createElement("div");
        cancelRow.style.cssText = "display:flex;justify-content:center;padding-top:12px;";
        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Cancel";
        cancelBtn.id = "archiveUpdateCancelBtn";
        cancelBtn.style.cssText = "padding:10px 24px;border-radius:6px;border:none;background:#dc3545;color:#fff;font-size:14px;font-weight:500;cursor:pointer;";
        cancelBtn.onmouseenter = function() { this.style.background = "#c82333"; };
        cancelBtn.onmouseleave = function() { this.style.background = "#dc3545"; };
        cancelBtn.addEventListener("click", function() {
            ARCHIVE_UPDATE_FORMS_CANCELLED = true;
            cancelBtn.disabled = true;
            cancelBtn.textContent = "Cancelling...";
            cancelBtn.style.opacity = "0.5";
            log("Archive/Update Forms: cancellation requested by user");
        });
        cancelRow.appendChild(cancelBtn);
        container.appendChild(cancelRow);

        // Helper methods
        container.updateStatus = function(text) {
            var el = document.getElementById("archiveUpdateStatus");
            if (el) el.textContent = text;
        };

        container.updateProgress = function(current, total) {
            var pct = total > 0 ? Math.round((current / total) * 100) : 0;
            var bar = document.getElementById("archiveUpdateProgressBar");
            if (bar) bar.style.width = pct + "%";
            this.updateStatus("Processing " + current + " of " + total);
        };

        container.setItemStatus = function(index, status, color) {
            var item = document.getElementById("archiveUpdateItem_" + index);
            if (item) {
                var statusEl = item.querySelector(".item-status");
                if (statusEl) {
                    statusEl.textContent = status;
                    statusEl.style.background = color || "#444";
                    statusEl.style.color = "#fff";
                }
            }
        };

        container.showSummary = function(total, success, skipped, errors) {
            var summaryEl = document.getElementById("archiveUpdateSummary");
            if (summaryEl) {
                summaryEl.style.display = "block";
                summaryEl.innerHTML = '<div style="font-size:16px;font-weight:600;margin-bottom:8px;">Complete</div>' +
                    '<div>Total: ' + total + ' | Success: <span style="color:#51cf66;">' + success + '</span> | Skipped: <span style="color:#ffc107;">' + skipped + '</span> | Errors: <span style="color:#ff6b6b;">' + errors + '</span></div>';
            }
            var cancelBtn = document.getElementById("archiveUpdateCancelBtn");
            if (cancelBtn) {
                cancelBtn.textContent = "Close";
                cancelBtn.style.background = "#6c757d";
                cancelBtn.onclick = function() {
                    if (ARCHIVE_UPDATE_FORMS_PROGRESS_REF) {
                        ARCHIVE_UPDATE_FORMS_PROGRESS_REF.close();
                        ARCHIVE_UPDATE_FORMS_PROGRESS_REF = null;
                    }
                };
            }
        };

        return container;
    }

    // Collect Edit modal properties
    async function collectEditModalProperties() {
        var props = {
            mandatory: false,
            hidden: false,
            preWindow: "",
            postWindow: "",
            referenceActivity: false,
            offsetPreReference: "",
            offsetDays: "",
            offsetHours: "",
            offsetMinutes: "",
            offsetSeconds: "",
            enforceDataCollectionOrder: false,
            disableCollectionTime: false,
            formOffsetSeconds: "",
            collectionRoleRestriction: ""
        };

        await sleep(500);

        // Mandatory
        var mandatoryEl = document.getElementById("mandatory");
        if (mandatoryEl) {
            props.mandatory = mandatoryEl.checked;
        }

        // Hidden - check both checked property and attribute (for disabled checkboxes)
        var hiddenEl = document.getElementById("hidden");
        if (hiddenEl) {
            props.hidden = hiddenEl.checked || hiddenEl.hasAttribute("checked");
        }

        // Pre-Window
        var preWindowEl = document.getElementById("preWindow");
        if (preWindowEl) {
            props.preWindow = preWindowEl.value || "";
        }

        // Post-Window
        var postWindowEl = document.getElementById("postWindow");
        if (postWindowEl) {
            props.postWindow = postWindowEl.value || "";
        }

        // Reference Activity
        var refActivityEl = document.getElementById("referenceActivity");
        if (refActivityEl) {
            props.referenceActivity = refActivityEl.checked;
        }

        // Offset fields (only if referenceActivity is unchecked)
        if (!props.referenceActivity) {
            var preRefEl = document.querySelector("#offset\\.preReference, [id='offset.preReference']");
            if (preRefEl) {
                props.offsetPreReference = preRefEl.checked ? "checked" : "";
            }

            var daysEl = document.querySelector("input[name='offset.days']");
            if (daysEl) props.offsetDays = daysEl.value || "";

            var hoursEl = document.querySelector("input[name='offset.hours']");
            if (hoursEl) props.offsetHours = hoursEl.value || "";

            var minutesEl = document.querySelector("input[name='offset.minutes']");
            if (minutesEl) props.offsetMinutes = minutesEl.value || "";

            var secondsEl = document.querySelector("input[name='offset.seconds']");
            if (secondsEl) props.offsetSeconds = secondsEl.value || "";
        }

        // Enforce Data Collection Order - check the parent span's class for "checked" status
        var enforceSpan = document.querySelector("#uniform-enforceDataCollectionOrder span");
        if (enforceSpan) {
            props.enforceDataCollectionOrder = enforceSpan.className.indexOf("checked") !== -1;
            log("Archive/Update Forms: collected enforceDataCollectionOrder=" + props.enforceDataCollectionOrder + " (span class='" + enforceSpan.className + "')");
        } else {
            // Fallback: try direct checkbox
            var enforceEl = document.getElementById("enforceDataCollectionOrder");
            if (enforceEl) {
                props.enforceDataCollectionOrder = enforceEl.checked;
                log("Archive/Update Forms: collected enforceDataCollectionOrder=" + props.enforceDataCollectionOrder + " (direct checkbox)");
            }
        }

        // Disable Collection Time - check the parent span's class for "checked" status
        var disableCollectionSpan = document.querySelector("#uniform-disableCollectionTime span");
        if (disableCollectionSpan) {
            props.disableCollectionTime = disableCollectionSpan.className.indexOf("checked") !== -1;
            log("Archive/Update Forms: collected disableCollectionTime=" + props.disableCollectionTime + " (span class='" + disableCollectionSpan.className + "')");
        } else {
            // Fallback: try direct checkbox
            var disableCollectionEl = document.getElementById("disableCollectionTime");
            if (disableCollectionEl) {
                props.disableCollectionTime = disableCollectionEl.checked;
                log("Archive/Update Forms: collected disableCollectionTime=" + props.disableCollectionTime + " (direct checkbox)");
            }
        }

        // Timepoint Offset (Form Offset Seconds)
        var formOffsetEl = document.getElementById("formOffsetSeconds");
        if (formOffsetEl) {
            props.formOffsetSeconds = formOffsetEl.value || "";
            log("Archive/Update Forms: collected formOffsetSeconds='" + props.formOffsetSeconds + "'");
        }

        // Collection Role Restriction
        var collectionRoleSpan = document.querySelector("#s2id_dataCollectionApplicationUserRole .select2-chosen");
        if (collectionRoleSpan) {
            props.collectionRoleRestriction = normalizeSAText(collectionRoleSpan.textContent);
            log("Archive/Update Forms: collected collectionRoleRestriction='" + props.collectionRoleRestriction + "'");
        }

        return props;
    }

    // Collect Visibility modal properties
    async function collectVisibilityModalProperties() {
        var visibility = {
            activityPlan: "",
            scheduledActivity: "",
            item: "",
            itemValue: ""
        };

        await sleep(500);

        // Look for static text in p.form-control-static elements
        var statics = document.querySelectorAll("p.form-control-static");
        var labels = document.querySelectorAll(".control-label, label");

        // Try to map by nearby labels
        for (var i = 0; i < statics.length; i++) {
            var staticEl = statics[i];
            var text = normalizeSAText(staticEl.textContent);
            var parent = staticEl.closest(".form-group, .control-group");
            if (parent) {
                var label = parent.querySelector(".control-label, label");
                if (label) {
                    var labelText = normalizeText(label.textContent);
                    if (labelText.indexOf("activityplan") !== -1) {
                        visibility.activityPlan = text;
                    } else if (labelText.indexOf("scheduledactivity") !== -1) {
                        visibility.scheduledActivity = text;
                    } else if (labelText.indexOf("itemvalue") !== -1) {
                        visibility.itemValue = text;
                    } else if (labelText.indexOf("item") !== -1) {
                        visibility.item = text;
                    }
                }
            }
        }

        return visibility;
    }

    // Apply properties to Add modal
    async function applyPropertiesToAddModal(props) {
        log("Archive/Update Forms: applyPropertiesToAddModal called with props=" + JSON.stringify(props));
        await sleep(700);
        // Hidden - use same approach as SA Builder
        var hiddenCheckboxEl = document.querySelector("#uniform-hidden span input#hidden.checkbox");
        if (!hiddenCheckboxEl) {
            hiddenCheckboxEl = document.getElementById("hidden");
        }
        log("Archive/Update Forms: Hidden checkbox - props.hidden=" + props.hidden + ", element exists=" + !!hiddenCheckboxEl);
        if (hiddenCheckboxEl) {
            log("Archive/Update Forms: Hidden checkbox - disabled=" + hiddenCheckboxEl.disabled + ", checked=" + hiddenCheckboxEl.checked);
        }
        if (hiddenCheckboxEl && !hiddenCheckboxEl.disabled) {
            if (props.hidden && !hiddenCheckboxEl.checked) {
                hiddenCheckboxEl.click();
                log("Archive/Update Forms: Hidden checkbox checked");
                await sleep(200);
                log("Archive/Update Forms: Hidden checkbox - after click, checked=" + hiddenCheckboxEl.checked);
            } else if (!props.hidden && hiddenCheckboxEl.checked) {
                hiddenCheckboxEl.click();
                log("Archive/Update Forms: Hidden checkbox unchecked");
                await sleep(200);
                log("Archive/Update Forms: Hidden checkbox - after click, checked=" + hiddenCheckboxEl.checked);
            } else {
                log("Archive/Update Forms: Hidden checkbox - already in correct state");
            }
        }

        // Mandatory - use same approach as SA Builder
        var mandatoryEl = document.querySelector("#uniform-mandatory span input#mandatory.checkbox");
        if (!mandatoryEl) {
            mandatoryEl = document.getElementById("mandatory");
        }
        if (mandatoryEl && !mandatoryEl.disabled) {
            if (props.mandatory && !mandatoryEl.checked) {
                mandatoryEl.click();
                log("Archive/Update Forms: Mandatory checkbox checked");
                await sleep(200);
            } else if (!props.mandatory && mandatoryEl.checked) {
                mandatoryEl.click();
                log("Archive/Update Forms: Mandatory checkbox unchecked");
                await sleep(200);
            }
        }

        // Pre-Window
        var preWindowEl = document.getElementById("preWindow");
        if (preWindowEl && !preWindowEl.disabled) {
            preWindowEl.value = props.preWindow || "";
            preWindowEl.dispatchEvent(new Event("input", { bubbles: true }));
            log("Archive/Update Forms: Pre-Window set to '" + (props.preWindow || "") + "'");
        }

        // Post-Window
        var postWindowEl = document.getElementById("postWindow");
        if (postWindowEl && !postWindowEl.disabled) {
            postWindowEl.value = props.postWindow || "";
            postWindowEl.dispatchEvent(new Event("input", { bubbles: true }));
            log("Archive/Update Forms: Post-Window set to '" + (props.postWindow || "") + "'");
        }

        // Reference Activity
        var refActivityEl = document.getElementById("referenceActivity");
        if (refActivityEl && !refActivityEl.disabled) {
            if (props.referenceActivity && !refActivityEl.checked) {
                refActivityEl.click();
                log("Archive/Update Forms: Reference Activity checkbox checked");
                await sleep(200);
            } else if (!props.referenceActivity && refActivityEl.checked) {
                refActivityEl.click();
                log("Archive/Update Forms: Reference Activity checkbox unchecked");
                await sleep(200);
            }
        }

        // Offset fields (only if referenceActivity is unchecked)
        if (!props.referenceActivity) {
            await sleep(200);

            // Pre-Reference checkbox - use same approach as SA Builder
            var preRefEl = document.querySelector("#uniform-offset\\.preReference span input#offset\\.preReference.checkbox");
            if (!preRefEl) {
                preRefEl = document.getElementById("offset.preReference");
            }
            if (preRefEl && !preRefEl.disabled) {
                if (props.offsetPreReference === "checked" && !preRefEl.checked) {
                    preRefEl.click();
                    log("Archive/Update Forms: Pre-Reference checkbox checked");
                    await sleep(200);
                } else if (props.offsetPreReference !== "checked" && preRefEl.checked) {
                    preRefEl.click();
                    log("Archive/Update Forms: Pre-Reference checkbox unchecked");
                    await sleep(200);
                }
            }

            // Time offset values
            var daysEl = document.querySelector("input[name='offset.days']");
            if (daysEl && !daysEl.disabled) {
                daysEl.value = String(props.offsetDays || "0");
                daysEl.dispatchEvent(new Event("input", { bubbles: true }));
            }

            var hoursEl = document.querySelector("input[name='offset.hours']");
            if (hoursEl && !hoursEl.disabled) {
                hoursEl.value = String(props.offsetHours || "0");
                hoursEl.dispatchEvent(new Event("input", { bubbles: true }));
            }

            var minutesEl = document.querySelector("input[name='offset.minutes']");
            if (minutesEl && !minutesEl.disabled) {
                minutesEl.value = String(props.offsetMinutes || "0");
                minutesEl.dispatchEvent(new Event("input", { bubbles: true }));
            }

            var secondsEl = document.querySelector("input[name='offset.seconds']");
            if (secondsEl && !secondsEl.disabled) {
                secondsEl.value = String(props.offsetSeconds || "0");
                secondsEl.dispatchEvent(new Event("input", { bubbles: true }));
            }
            log("Archive/Update Forms: Time offset set to " + props.offsetDays + "d " + props.offsetHours + "h " + props.offsetMinutes + "m " + props.offsetSeconds + "s");
        } else {
            log("Archive/Update Forms: Skipping time offset values (Reference Activity checked)");
        }
        // Enforce Data Collection Order
        if (props.enforceDataCollectionOrder) {
            var enforceEl = document.querySelector("#uniform-enforceDataCollectionOrder span input#enforceDataCollectionOrder.checkbox");
            if (!enforceEl) {
                enforceEl = document.getElementById("enforceDataCollectionOrder");
            }
            if (enforceEl && !enforceEl.checked) {
                enforceEl.click();
                log("Archive/Update Forms: Enforce Data Collection Order checkbox checked");
                await sleep(200);
            }
        }

        // Disable Collection Time
        var disableCollectionEl = document.querySelector("#uniform-disableCollectionTime span input#disableCollectionTime.checkbox");
        if (!disableCollectionEl) {
            disableCollectionEl = document.getElementById("disableCollectionTime");
        }
        if (disableCollectionEl && !disableCollectionEl.disabled) {
            if (props.disableCollectionTime && !disableCollectionEl.checked) {
                disableCollectionEl.click();
                log("Archive/Update Forms: Disable Collection Time checkbox checked");
                await sleep(200);
            } else if (!props.disableCollectionTime && disableCollectionEl.checked) {
                disableCollectionEl.click();
                log("Archive/Update Forms: Disable Collection Time checkbox unchecked");
                await sleep(200);
            }
        }

        // Timepoint Offset (Form Offset Seconds)
        var formOffsetEl = document.getElementById("formOffsetSeconds");
        if (formOffsetEl && !formOffsetEl.disabled) {
            formOffsetEl.value = props.formOffsetSeconds || "0";
            formOffsetEl.dispatchEvent(new Event("input", { bubbles: true }));
            log("Archive/Update Forms: Timepoint Offset set to '" + (props.formOffsetSeconds || "0") + "' seconds");
        }

        // Collection Role Restriction
        if (props.collectionRoleRestriction) {
            var roleSet = await setSelect2ValueByText("dataCollectionApplicationUserRole", props.collectionRoleRestriction);
            if (roleSet) {
                log("Archive/Update Forms: Collection Role Restriction set to '" + props.collectionRoleRestriction + "'");
            } else {
                log("Archive/Update Forms: Failed to set Collection Role Restriction to '" + props.collectionRoleRestriction + "'");
            }
            await sleep(300);
        }
    }

    // Wait for Select2 options to change/load
    async function waitForSelect2OptionsChange(selectId, timeoutMs) {
        var start = Date.now();
        var maxTime = timeoutMs || 5000;
        var initialCount = 0;
        var sel = document.getElementById(selectId);
        if (sel) {
            initialCount = sel.querySelectorAll("option").length;
        }
        while (Date.now() - start < maxTime) {
            sel = document.getElementById(selectId);
            if (sel) {
                var currentCount = sel.querySelectorAll("option").length;
                if (currentCount > initialCount || currentCount > 1) {
                    await sleep(200);
                    return true;
                }
            }
            await sleep(150);
        }
        return false;
    }

    // Set Select2 value by text matching
        async function setSelect2ValueByText(selectId, targetText) {
        var sel = document.getElementById(selectId);
        if (!sel) {
            log("Archive/Update Forms: select " + selectId + " not found");
            return false;
        }

        // Use specialized normalization for visibility fields
        var isVisibilityField = selectId.indexOf("visible") === 0;
        var normalizedTarget = isVisibilityField ? normalizeVisibilityText(targetText) : normalizeText(targetText);
        var opts = sel.querySelectorAll("option");
        var matchValue = null;

        // First pass: exact match
        for (var i = 0; i < opts.length; i++) {
            var opt = opts[i];
            var optText = normalizeSAText(opt.textContent);
            var optNorm = isVisibilityField ? normalizeVisibilityText(optText) : normalizeText(optText);
            
            if (optNorm === normalizedTarget || optText === targetText) {
                matchValue = opt.value;
                log("Archive/Update Forms: exact match found for " + targetText + " -> " + optText);
                break;
            }
        }

        // Second pass: fuzzy match if exact not found
        if (!matchValue) {
            for (var j = 0; j < opts.length; j++) {
                var opt2 = opts[j];
                var optText2 = isVisibilityField ? normalizeVisibilityText(opt2.textContent) : normalizeText(opt2.textContent);
                
                if (optText2.indexOf(normalizedTarget) !== -1 || normalizedTarget.indexOf(optText2) !== -1) {
                    matchValue = opt2.value;
                    log("Archive/Update Forms: fuzzy match found for " + targetText + " -> " + opt2.textContent);
                    break;
                }
            }
        }

        if (matchValue) {
            sel.value = matchValue;
            sel.dispatchEvent(new Event("change", { bubbles: true }));
            try {
                if (window.jQuery && window.jQuery.fn.select2) {
                    window.jQuery("#" + selectId).trigger("change");
                }
            } catch (e) {}
            await sleep(300);
            return true;
        }

        log("Archive/Update Forms: could not find option for " + targetText + " in " + selectId);
        return false;
    }

    // Check if target form already exists for segment/event
    function checkTargetFormExists(rows, segmentKey, eventKey, targetFormValue) {
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var rowSegKey = row.segmentValue || normalizeText(row.segmentText);
            var rowEvKey = row.eventValue || normalizeText(row.eventText);
            if (rowSegKey === segmentKey && rowEvKey === eventKey && row.formKey === targetFormValue) {
                return true;
            }
        }
        return false;
    }

    // Find row in DOM by matching criteria
    function findRowInDOM(segmentText, eventText, formText) {
        var tbody = document.getElementById("saTableBody");
        if (!tbody) return null;
        var trs = tbody.querySelectorAll("tr");
        for (var i = 0; i < trs.length; i++) {
            var tr = trs[i];
            var cells = tr.querySelectorAll("td");
            if (cells.length < 4) continue;
            var seg = normalizeSAText(cells[1].textContent);
            var ev = normalizeSAText(cells[2].textContent);
            var form = normalizeSAText(cells[3].textContent);
            if (seg === segmentText && ev === eventText && form === formText) {
                return tr;
            }
        }
        return null;
    }

    // Click Action dropdown for a row
    async function clickRowActionDropdown(row) {
        var actionBtn = row.querySelector("button.dropdown-toggle, a.dropdown-toggle");
        if (actionBtn) {
            actionBtn.click();
            await sleep(300);
            return true;
        }
        // Try td with Actions text
        var cells = row.querySelectorAll("td");
        for (var i = 0; i < cells.length; i++) {
            var btn = cells[i].querySelector("button, a.dropdown-toggle");
            if (btn) {
                btn.click();
                await sleep(300);
                return true;
            }
        }
        return false;
    }

    // Main Archive/Update Forms execution
    async function executeArchiveUpdateForms(sourceFormKey, targetFormValue, sourceFormsMap, targetFormsArray, archiveReason, visibilityReason) {
        var sourceForm = sourceFormsMap[sourceFormKey];

        // Find target form from array by value
        var targetForm = null;
        for (var t = 0; t < targetFormsArray.length; t++) {
            if (targetFormsArray[t].value === targetFormValue) {
                targetForm = targetFormsArray[t];
                break;
            }
        }

        if (!sourceForm) {
            log("Archive/Update Forms: invalid source form");
            return;
        }

        if (!targetForm) {
            log("Archive/Update Forms: invalid target form");
            return;
        }

        var occurrences = sourceForm.occurrences.slice(); // Copy array
        var total = occurrences.length;

        if (total === 0) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">No occurrences found for selected Source form.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        log("Archive/Update Forms: processing " + total + " occurrences");

        // Create progress popup
        var progressContent = createArchiveUpdateProgressPopup(occurrences, targetForm.text);
        ARCHIVE_UPDATE_FORMS_PROGRESS_REF = createPopup({
            title: "Archive/Update Forms - Processing",
            content: progressContent,
            width: "600px",
            height: "70%",
            maxHeight: "600px",
            onClose: function() {
                ARCHIVE_UPDATE_FORMS_CANCELLED = true;
                ARCHIVE_UPDATE_FORMS_PROGRESS_REF = null;
            }
        });

        var successCount = 0;
        var skipCount = 0;
        var errorCount = 0;

        for (var i = 0; i < occurrences.length; i++) {
            if (ARCHIVE_UPDATE_FORMS_CANCELLED) {
                log("Archive/Update Forms: cancelled by user");
                break;
            }

            var occ = occurrences[i];
            progressContent.updateProgress(i + 1, total);
            progressContent.setItemStatus(i, "Processing...", "#17a2b8");

            try {
                // Re-scan to get fresh row reference
                var freshRows = scanSATableForArchiveUpdate();
                var segKey = occ.segmentValue || normalizeText(occ.segmentText);
                var evKey = occ.eventValue || normalizeText(occ.eventText);

                // Check if target already exists for this segment/event
                if (checkTargetFormExists(freshRows, segKey, evKey, targetFormValue)) {
                    log("Archive/Update Forms: target form already exists for " + occ.segmentText + " → " + occ.eventText + "; skipping");
                    progressContent.setItemStatus(i, "Skipped (exists)", "#ffc107");
                    skipCount++;
                    continue;
                }

                // Find the occurrence row
                var row = findRowInDOM(occ.segmentText, occ.eventText, occ.formText);
                if (!row) {
                    log("Archive/Update Forms: could not find row for " + occ.rowKey);
                    progressContent.setItemStatus(i, "Error (not found)", "#dc3545");
                    errorCount++;
                    continue;
                }

                // Step 1: Open Edit modal and collect properties
                var editProps = null;
                // Find edit link directly in the row (8th column)
                var editLink = row.querySelector('a[href*="/update/scheduledactivity/"]');
                if (!editLink) {
                    // Fallback: try other selectors
                    editLink = row.querySelector('a[href*="editscheduledactivity"]');
                }
                if (!editLink) {
                    // Fallback: look for Edit text
                    var links = row.querySelectorAll('a[data-toggle="modal"]');
                    for (var li = 0; li < links.length; li++) {
                        if (links[li].textContent.indexOf("Edit") !== -1) {
                            editLink = links[li];
                            break;
                        }
                    }
                }
                log("Archive/Update Forms: editLink found=" + !!editLink);
                if (editLink) {
                    editLink.click();
                    var editModal = await waitForSAModal(10000);
                    if (editModal) {
                        editProps = await collectEditModalProperties();
                        log("Archive/Update Forms: collected editProps.hidden=" + editProps.hidden + ", preWindow='" + editProps.preWindow + "', postWindow='" + editProps.postWindow + "'");
                        // Cancel the edit modal
                        var cancelBtn = editModal.querySelector("button[data-dismiss='modal'], .btn-default, .close");
                        if (cancelBtn) {
                            cancelBtn.click();
                        } else {
                            document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
                        }
                        await waitForSAModalClose(5000);
                    } else {
                        log("Archive/Update Forms: edit modal did not open for " + occ.rowKey);
                    }
                } else {
                    log("Archive/Update Forms: edit link not found for " + occ.rowKey);
                }

                if (ARCHIVE_UPDATE_FORMS_CANCELLED) break;

                // Step 2: Collect visibility if hidden
                var visibilityProps = null;
                if (editProps && editProps.hidden) {
                    log("Archive/Update Forms: Step 2 - attempting to collect visibility properties");
                    await sleep(500);
                    row = findRowInDOM(occ.segmentText, occ.eventText, occ.formText);
                    if (row) {
                        // Find visibility link directly in the row
                        var visLink = row.querySelector('a[href*="visiblecondition"]');
                        if (!visLink) {
                            visLink = row.querySelector('a[href*="visibility"]');
                        }
                        log("Archive/Update Forms: Step 2 - visibility link found=" + !!visLink);
                        if (visLink) {
                            visLink.click();
                            var visModal = await waitForSAModal(10000);
                            log("Archive/Update Forms: Step 2 - visibility modal opened=" + !!visModal);
                            if (visModal) {
                                visibilityProps = await collectVisibilityModalProperties();
                                log("Archive/Update Forms: Step 2 - collected visibilityProps=" + JSON.stringify(visibilityProps));
                                var visCancelBtn = visModal.querySelector("button[data-dismiss='modal'], .btn-default, .close");
                                if (visCancelBtn) {
                                    visCancelBtn.click();
                                } else {
                                    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
                                }
                                await waitForSAModalClose(5000);
                            } else {
                                log("Archive/Update Forms: Step 2 - visibility modal did not open");
                            }
                        } else {
                            log("Archive/Update Forms: Step 2 - visibility link not found in row");
                        }
                    } else {
                        log("Archive/Update Forms: Step 2 - could not re-find row");
                    }
                }

                if (ARCHIVE_UPDATE_FORMS_CANCELLED) break;

                // Step 3: Archive the source occurrence
                await sleep(500);
                row = findRowInDOM(occ.segmentText, occ.eventText, occ.formText);
                if (row) {
                    await clickRowActionDropdown(row);
                    var archiveLink = row.querySelector('a[href*="archivescheduledactivity"]');
                    if (archiveLink) {
                        // Check if it's Un-Archive (already archived)
                        var archiveLinkText = (archiveLink.textContent || "").toLowerCase();
                        if (archiveLinkText.indexOf("un-archive") !== -1) {
                            log("Archive/Update Forms: row already archived for " + occ.rowKey + "; skipping archive step");
                        } else {
                            archiveLink.click();
                            var archiveModal = await waitForSAModal(10000);
                            if (archiveModal) {
                                // Set reason for change
                                var reasonEl = document.getElementById("reasonForChange");
                                if (reasonEl) {
                                    reasonEl.value = archiveReason || "Old version";
                                    reasonEl.dispatchEvent(new Event("change", { bubbles: true }));
                                }
                                // Click Save
                                var saveBtn = document.getElementById("actionButton");
                                if (saveBtn) {
                                    saveBtn.click();
                                    var closed = await waitForSAModalClose(10000);
                                    if (!closed) {
                                        log("Archive/Update Forms: archive modal did not close for " + occ.rowKey);
                                        progressContent.setItemStatus(i, "Error (archive)", "#dc3545");
                                        errorCount++;
                                        continue;
                                    }
                                }
                                await sleep(1000);
                            }
                        }
                    } else {
                        log("Archive/Update Forms: archive link not found for " + occ.rowKey);
                    }
                }

                if (ARCHIVE_UPDATE_FORMS_CANCELLED) break;

                // Step 4: Add target form
                await sleep(500);
                if (!clickAddSaButton()) {
                    log("Archive/Update Forms: could not click Add button");
                    progressContent.setItemStatus(i, "Error (add)", "#dc3545");
                    errorCount++;
                    continue;
                }

                var addModal = await waitForSAModal(10000);
                if (!addModal) {
                    log("Archive/Update Forms: add modal did not open");
                    progressContent.setItemStatus(i, "Error (add modal)", "#dc3545");
                    errorCount++;
                    continue;
                }

                // Select segment by text (more reliable than value)
                await setSelect2ValueByText("segment", occ.segmentText);
                await sleep(500);

                // Select study event by text (more reliable than value)
                await setSelect2ValueByText("studyEvent", occ.eventText);
                await sleep(500);

                // Select target form by text (more reliable than value)
                await setSelect2ValueByText("form", targetForm.text);
                await sleep(500);

                // Apply copied properties
                if (editProps) {
                    await applyPropertiesToAddModal(editProps);
                }

                // Save the new scheduled activity
                var addSaveBtn = document.getElementById("actionButton");
                if (addSaveBtn) {
                    addSaveBtn.click();
                    var addClosed = await waitForSAModalClose(10000);
                    if (!addClosed) {
                        log("Archive/Update Forms: add modal did not close");
                        progressContent.setItemStatus(i, "Error (save)", "#dc3545");
                        errorCount++;
                        continue;
                    }
                }

                await sleep(1000);

                if (ARCHIVE_UPDATE_FORMS_CANCELLED) break;

                // Step 5: Set visibility if source was hidden
                log("Archive/Update Forms: Step 5 check - editProps.hidden=" + (editProps && editProps.hidden) + ", visibilityProps=" + !!visibilityProps);
                if (editProps && editProps.hidden && visibilityProps) {
                    log("Archive/Update Forms: Step 5 - attempting to set visibility on new row");
                    // Find the newly added row
                    var newRow = findRowInDOM(occ.segmentText, occ.eventText, targetForm.text);
                    if (newRow) {
                        log("Archive/Update Forms: Step 5 - new row found");
                        // Find visibility link directly in the row
                        var newVisLink = newRow.querySelector('a[href*="visiblecondition"]');
                        if (!newVisLink) {
                            newVisLink = newRow.querySelector('a[href*="visibility"]');
                        }
                        log("Archive/Update Forms: Step 5 - visibility link found=" + !!newVisLink);
                        if (newVisLink) {
                            newVisLink.click();
                            var newVisModal = await waitForSAModal(10000);
                            log("Archive/Update Forms: Step 5 - visibility modal opened=" + !!newVisModal);
                            if (newVisModal) {
                                await sleep(500);

                                // Set visibility fields one at a time with delays
                                var visSuccess = true;
                                
                                // 1. Activity Plan
                                if (visibilityProps.activityPlan) {
                                    log("Archive/Update Forms: Step 5 - setting Activity Plan: " + visibilityProps.activityPlan);
                                    var apSet = await setSelect2ValueByText("visibleActivityPlan", visibilityProps.activityPlan);
                                    if (!apSet) {
                                        log("Archive/Update Forms: Step 5 - failed to set Activity Plan");
                                        visSuccess = false;
                                    } else {
                                        await sleep(500);
                                        await waitForSelect2OptionsChange("visibleScheduledActivity", 3000);
                                        await sleep(500);
                                    }
                                }
                                
                                // 2. Scheduled Activity
                                if (visSuccess && visibilityProps.scheduledActivity) {
                                    log("Archive/Update Forms: Step 5 - setting Scheduled Activity: " + visibilityProps.scheduledActivity);
                                    var saSet = false;
                                    for (var retry = 0; retry < 3; retry++) {
                                        saSet = await setSelect2ValueByText("visibleScheduledActivity", visibilityProps.scheduledActivity);
                                        if (saSet) break;
                                        log("Archive/Update Forms: Step 5 - retry " + (retry + 1) + " for Scheduled Activity");
                                        await sleep(1000);
                                    }
                                    if (!saSet) {
                                        log("Archive/Update Forms: Step 5 - failed to set Scheduled Activity after retries");
                                        visSuccess = false;
                                    } else {
                                        await sleep(500);
                                        await waitForSelect2OptionsChange("visibleItemRef", 3000);
                                        await sleep(500);
                                    }
                                }
                                
                                // 3. Item
                                if (visSuccess && visibilityProps.item) {
                                    log("Archive/Update Forms: Step 5 - setting Item: " + visibilityProps.item);
                                    var itemSet = false;
                                    for (var retry = 0; retry < 3; retry++) {
                                        itemSet = await setSelect2ValueByText("visibleItemRef", visibilityProps.item);
                                        if (itemSet) break;
                                        log("Archive/Update Forms: Step 5 - retry " + (retry + 1) + " for Item");
                                        await sleep(1000);
                                    }
                                    if (!itemSet) {
                                        log("Archive/Update Forms: Step 5 - failed to set Item after retries");
                                        visSuccess = false;
                                    } else {
                                        await sleep(500);
                                        await waitForSelect2OptionsChange("visibleCodeListItem", 3000);
                                        await sleep(500);
                                    }
                                }
                                
                                // 4. Item Value
                                if (visSuccess && visibilityProps.itemValue) {
                                    log("Archive/Update Forms: Step 5 - setting Item Value: " + visibilityProps.itemValue);
                                    var ivSet = false;
                                    for (var retry = 0; retry < 3; retry++) {
                                        ivSet = await setSelect2ValueByText("visibleCodeListItem", visibilityProps.itemValue);
                                        if (ivSet) break;
                                        log("Archive/Update Forms: Step 5 - retry " + (retry + 1) + " for Item Value");
                                        await sleep(1000);
                                    }
                                    if (!ivSet) {
                                        log("Archive/Update Forms: Step 5 - failed to set Item Value after retries");
                                        visSuccess = false;
                                    } else {
                                        await sleep(500);
                                    }
                                }

                                // Set reason for change
                                var visReasonEl = document.getElementById("reasonForChange");
                                if (visReasonEl) {
                                    visReasonEl.value = visibilityReason || "Add visibility condition";
                                    visReasonEl.dispatchEvent(new Event("change", { bubbles: true }));
                                }

                                // Save visibility
                                var visSaveBtn = document.getElementById("actionButton");
                                if (visSaveBtn) {
                                    log("Archive/Update Forms: Step 5 - clicking Save button");
                                    visSaveBtn.click();
                                    await waitForSAModalClose(10000);
                                }
                                await sleep(500);
                                log("Archive/Update Forms: Step 5 - visibility set successfully");
                            } else {
                                log("Archive/Update Forms: Step 5 - visibility modal did not open");
                            }
                        } else {
                            log("Archive/Update Forms: Step 5 - visibility link not found in new row");
                        }
                    } else {
                        log("Archive/Update Forms: Step 5 - could not find newly added row to set visibility");
                    }
                }

                progressContent.setItemStatus(i, "Success", "#28a745");
                successCount++;
                log("Archive/Update Forms: processed " + occ.rowKey + " successfully");

            } catch (err) {
                log("Archive/Update Forms: error processing " + occ.rowKey + " - " + String(err));
                progressContent.setItemStatus(i, "Error", "#dc3545");
                errorCount++;
            }

            await sleep(500);
        }

        // Show summary
        progressContent.updateStatus("Complete");
        progressContent.showSummary(total, successCount, skipCount, errorCount);
        log("Archive/Update Forms: completed. Total=" + total + " Success=" + successCount + " Skipped=" + skipCount + " Errors=" + errorCount);
    }

    // Main entry point
    async function runArchiveUpdateForms() {
        // Check if on correct page
        if (!isOnArchiveUpdateFormsPage()) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#ff6b6b;font-size:16px;margin-bottom:12px;">⚠️ Wrong Page</p><p>Navigate to the Activity Plans Show page first.</p><p style="margin-top:12px;font-size:12px;color:#888;">Required URL: https://cenexeltest.clinspark.com/secure/crfdesign/activityplans/show/{id}</p></div>',
                width: "450px",
                height: "auto"
            });
            log("Archive/Update Forms: wrong page - " + location.href);
            return;
        }

        // Verify Add button is available
        if (isAddSaButtonDisabled()) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#ff6b6b;font-size:16px;margin-bottom:12px;">⚠️ Add Button Disabled</p><p>The Add button is currently disabled. This activity plan may no longer be in design mode.</p></div>',
                width: "400px",
                height: "auto"
            });
            log("Archive/Update Forms: Add button is disabled");
            return;
        }

        // Scan table and build forms map
        log("Archive/Update Forms: scanning table...");
        var rows = scanSATableForArchiveUpdate();
        if (rows.length === 0) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">No rows found in the Scheduled Activities table.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        var sourceFormsMap = buildUniqueFormsMap(rows);
        var sourceFormCount = Object.keys(sourceFormsMap).length;
        log("Archive/Update Forms: found " + sourceFormCount + " unique source forms from " + rows.length + " rows");

        if (sourceFormCount === 0) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">No forms found in the Scheduled Activities table.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        // Show loading popup while collecting target forms from Add modal
        var loadingPopup = createPopup({
            title: "Archive/Update Forms",
            content: '<div style="text-align:center;padding:30px;"><div style="font-size:16px;margin-bottom:12px;">Collecting available forms...</div><div id="archiveUpdateLoadingDots" style="font-size:24px;">.</div></div>',
            width: "350px",
            height: "auto"
        });
        var loadingInterval = setInterval(function() {
            var dots = document.getElementById("archiveUpdateLoadingDots");
            if (dots) {
                var d = dots.textContent;
                dots.textContent = d.length >= 3 ? "." : d + ".";
            }
        }, 400);

        // Click Add button to open modal
        log("Archive/Update Forms: clicking Add button to collect forms...");
        if (!clickAddSaButton()) {
            clearInterval(loadingInterval);
            loadingPopup.close();
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">Could not click the Add button.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        // Wait for modal to appear
        var modal = await waitForSAModal(10000);
        if (!modal) {
            clearInterval(loadingInterval);
            loadingPopup.close();
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">Modal did not appear within timeout.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        // Collect form options from the modal dropdown
        await openSelect2Dropdown("s2id_form");
        await sleep(300);
        await closeSelect2Dropdown();
        var targetFormsArray = collectSelectOptions("form");
        log("Archive/Update Forms: collected " + targetFormsArray.length + " target forms from modal");

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

        if (targetFormsArray.length === 0) {
            createPopup({
                title: "Archive/Update Forms - Error",
                content: '<div style="text-align:center;padding:20px;color:#ff6b6b;">No forms found in the Add modal dropdown.</div>',
                width: "400px",
                height: "auto"
            });
            return;
        }

        // Show selection GUI with separate source and target form lists
        var guiContent = createArchiveUpdateFormsGUI(sourceFormsMap, targetFormsArray);

        ARCHIVE_UPDATE_FORMS_POPUP_REF = createPopup({
            title: "Archive/Update Forms - Select Forms",
            content: guiContent,
            width: "900px",
            maxWidth: "95%",
            height: "70%",
            maxHeight: "700px",
            onClose: function() {
                ARCHIVE_UPDATE_FORMS_POPUP_REF = null;
            }
        });

        // Attach confirm handler
        guiContent.confirmBtn.addEventListener("click", async function() {
            var selection = guiContent.getSelection();
            if (!selection.source || !selection.target || selection.source === selection.target) {
                return;
            }

            log("Archive/Update Forms: confirmed - source=" + selection.source + " target=" + selection.target + " archiveReason='" + selection.archiveReason + "' visibilityReason='" + selection.visibilityReason + "'");

            // Close selection popup (but don't cancel)
            if (ARCHIVE_UPDATE_FORMS_POPUP_REF) {
                // Remove the onClose handler to prevent it from running
                var popupEl = ARCHIVE_UPDATE_FORMS_POPUP_REF.element;
                if (popupEl) popupEl.remove();
                ARCHIVE_UPDATE_FORMS_POPUP_REF = null;
            }

            // Execute the replacement
            await executeArchiveUpdateForms(selection.source, selection.target, guiContent.sourceFormsMap, guiContent.targetFormsArray, selection.archiveReason, selection.visibilityReason);
        });

        log("Archive/Update Forms: selection GUI displayed");
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
        if (METHODS_LIBRARY_MODAL_REF && document.body.contains(METHODS_LIBRARY_MODAL_REF)) {
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
                    
                    // Badges container
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
                    
                    // Action buttons (show on hover)
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
                        if (selectedMethod !== m) item.style.background = "#252525";
                        actions.style.display = "flex";
                    };
                    item.onmouseleave = function() {
                        if (selectedMethod !== m) item.style.background = "transparent";
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
            
            // Track as recent and save as last method
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
            if (m.tags && m.tags.length > 0) metaParts.push("Tags: " + m.tags.join(", "));
            if (m.updated) metaParts.push("Updated: " + m.updated);
            previewMeta.textContent = metaParts.join("  •  ");

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
            
            // Save session state
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
        }

        closeBtn.onclick = closeModal;
        overlay.onclick = closeModal;

        function keyHandler(e) {
            if (e.key === "Escape") {
                closeModal();
                return;
            }
            
            // Enter key when search box is focused - select first result
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
            
            // Arrow Up/Down navigation in list
            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
                var items = Array.from(listPane.querySelectorAll("[role='option']"));
                if (items.length === 0) return;
                
                var currentIndex = -1;
                for (var i = 0; i < items.length; i++) {
                    if (items[i].getAttribute("aria-selected") === "true") {
                        currentIndex = i;
                        break;
                    }
                }
                
                var nextIndex;
                if (e.key === "ArrowDown") {
                    nextIndex = currentIndex < items.length - 1 ? currentIndex + 1 : 0;
                } else {
                    nextIndex = currentIndex > 0 ? currentIndex - 1 : items.length - 1;
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
            
            // Tab cycling
            if (e.key === "Tab") {
                var focusable = modal.querySelectorAll('button, input, select, [tabindex]:not([tabindex="-1"])');
                if (focusable.length === 0) return;
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
            if (e.target === closeBtn) return;
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
            if (!isDragging) return;
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

    //=========================
    // SETTING FEATURE
    //=========================
    // This section contains functions for setting up the extension.
    // It includes functions for loading and saving settings, as well as
    // functions for updating the UI based on the current settings.
    //=========================
    var STORAGE_BUTTON_VISIBILITY = "activityPlanState.buttonVisibility";
    var SETTINGS_POPUP_REF = null;
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
            "Lock Activity Plans",
            "Lock Sample Paths",
            "Update Study Status",
            "Add Cohort Subjects",
            "Run ICF Consent",
            "Run Button (1-5)",
            "Import Cohort Subjects",
            "Run Barcode",
            "Add Existing Subject",
            "Search Methods",
            "Scheduled Activities Builder",
            "Run Form (OOR) Below Range",
            "Run Form (OOR) Above Range",
            "Run Form (IR) In Range",
            "Collect All",
            "Import I/E",
            "Clear Mapping",
            "Archive/Update Forms",
            "Item Method Forms",
            "Find Form",
            "Find Study Events",
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

    // =========================
    // UI SCALING FEATURE
    // =========================


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
    // ADD EXISTING SUBJECT FEATURE
    //==========================
    // This section contains all functions for adding an existing subject from other epochs.
    //==========================
    // Add Existing Subject Feature

    var ADD_EXISTING_SUBJECT_POPUP_TITLE = "Add Existing Subject";
    var ADD_EXISTING_SUBJECT_POPUP_DESCRIPTION = "Auto-navigate to Add Existing Subject data page based on keywords and status";
    var RUNMODE_ADD_EXISTING_SUBJECT = "addExistingSubject";
    var STORAGE_AES_POPUP = "activityPlanState.aes.popup";
    var STORAGE_AES_SELECTED_EPOCH = "activityPlanState.aes.selectedEpoch";
    var STORAGE_AES_SELECTED_EPOCH_NAME = "activityPlanState.aes.selectedEpochName";
    var STORAGE_AES_COLLECTED_DATA = "activityPlanState.aes.collectedData";
    var STORAGE_AES_SELECTED_SUBJECT = "activityPlanState.aes.selectedSubject";
    var STORAGE_AES_SELECTED_VOLUNTEER_ID = "activityPlanState.aes.selectedVolunteerId";
    var STORAGE_AES_STEP = "activityPlanState.aes.step";
    var STORAGE_AES_COHORT_ID = "activityPlanState.aes.cohortId";
    var STORAGE_AES_COHORT_EDIT_DONE = "activityPlanState.aes.cohortEditDone";
    var STORAGE_AES_ASSIGNMENT_ID = "activityPlanState.aes.assignmentId";
    var STORAGE_AES_PROGRESS = "activityPlanState.aes.progress";
    var ADD_EXISTING_SUBJECT_POPUP_REF = null;

    function clearAddExistingSubjectData() {
        try {
            localStorage.removeItem(STORAGE_AES_POPUP);
            localStorage.removeItem(STORAGE_AES_SELECTED_EPOCH);
            localStorage.removeItem(STORAGE_AES_SELECTED_EPOCH_NAME);
            localStorage.removeItem(STORAGE_AES_COLLECTED_DATA);
            localStorage.removeItem(STORAGE_AES_SELECTED_SUBJECT);
            localStorage.removeItem(STORAGE_AES_SELECTED_VOLUNTEER_ID);
            localStorage.removeItem(STORAGE_AES_STEP);
            localStorage.removeItem(STORAGE_AES_COHORT_ID);
            localStorage.removeItem(STORAGE_AES_COHORT_EDIT_DONE);
            localStorage.removeItem(STORAGE_AES_ASSIGNMENT_ID);
            localStorage.removeItem(STORAGE_AES_PROGRESS);
        } catch (e) {}
        ADD_EXISTING_SUBJECT_POPUP_REF = null;
        log("AES: data cleared");
    }

    function getAESStep() {
        try {
            return localStorage.getItem(STORAGE_AES_STEP) || "";
        } catch (e) { return ""; }
    }

    function setAESStep(step) {
        try {
            localStorage.setItem(STORAGE_AES_STEP, step);
            log("AES: step=" + step);
        } catch (e) {}
    }

    function getAESProgress() {
        try {
            var raw = localStorage.getItem(STORAGE_AES_PROGRESS);
            if (raw) return JSON.parse(raw);
        } catch (e) {}
        return { current: 0, total: 0, message: "" };
    }

    function setAESProgress(current, total, message) {
        try {
            localStorage.setItem(STORAGE_AES_PROGRESS, JSON.stringify({ current: current, total: total, message: message }));
        } catch (e) {}
    }

    function setCheckboxStateById(id, state) {
        log("setCheckboxStateById: id=" + String(id) + " targetState=" + String(!!state));
        var el = document.getElementById(id);
        if (!el) {
            log("setCheckboxStateById: element not found id=" + String(id));
            return false;
        }
        var before = !!el.checked;
        log("setCheckboxStateById: currentState=" + String(before));
        el.checked = !!state;
        var evt = new Event("change", { bubbles: true });
        el.dispatchEvent(evt);
        var wrap = el.closest("div.checker");
        var span = null;
        if (wrap) {
            span = wrap.querySelector("span");
            if (span) {
                if (state) {
                    span.classList.add("checked");
                } else {
                    span.classList.remove("checked");
                }
            }
        }
        var after = !!el.checked;
        log("setCheckboxStateById: afterState=" + String(after));
        return true;
    }

    function isScreeningEpoch(name) {
        if (!name) return false;
        var lower = name.toLowerCase();
        return lower.indexOf("screen") !== -1;
    }

    function extractVolunteerIdFromHref(href) {
        if (!href) return null;
        var m = href.match(/\/secure\/volunteers\/manage\/show\/(\d+)/);
        if (m && m[1]) return m[1];
        return null;
    }

    function parseVolunteerInfo(text) {
        if (!text) return { initials: "", age: "", gender: "", id: "" };
        var cleaned = text.replace(/\s+/g, " ").trim();
        var parts = cleaned.split(/[,\s]+/).filter(function(p) { return p.length > 0; });
        var initials = parts[0] || "";
        var age = parts[1] || "";
        var gender = parts[2] || "";
        var idMatch = cleaned.match(/\(ID:\s*(\d+)\)/i);
        var id = idMatch ? idMatch[1] : "";
        return { initials: initials, age: age, gender: gender, id: id };
    }

    function formatVolunteerDisplay(info) {
        return info.initials + " - " + info.age + " - " + info.gender + " - " + info.id;
    }

    async function waitForSelectorWithRetry(selector, timeoutMs, retries) {
        for (var i = 0; i < retries; i++) {
            var el = await waitForSelector(selector, timeoutMs);
            if (el) return el;
            log("AES: retry " + (i + 1) + "/" + retries + " for " + selector);
            await sleep(500);
        }
        return null;
    }

    async function startAddExistingSubject() {
        log("AES: starting Add Existing Subject");
        var studiesUrl = "https://cenexeltest.clinspark.com/secure/administration/studies/show";

        if (location.href.indexOf(studiesUrl) === -1) {
            log("AES: navigating to studies page");
            setRunMode(RUNMODE_ADD_EXISTING_SUBJECT);
            setAESStep("selectEpoch");
            localStorage.setItem(STORAGE_AES_POPUP, "1");
            location.href = studiesUrl;
            return;
        }

        await showEpochSelectionPopup();
    }

    async function showEpochSelectionPopup() {
        log("AES: showing epoch selection popup");

        var epochTableBody = await waitForSelectorWithRetry("#epochTableBody", 5000, 3);
        if (!epochTableBody) {
            log("AES: epochTableBody not found");
            showAESError("Could not find epoch list. Please try again.");
            return;
        }

        var rows = epochTableBody.querySelectorAll("tr");
        var epochs = [];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var anchor = row.querySelector("td a[href*='/secure/administration/studies/epoch/show/']");
            if (anchor) {
                var name = (anchor.textContent || "").trim();
                var href = anchor.getAttribute("href") || "";
                var idMatch = href.match(/\/show\/(\d+)/);
                var epochId = idMatch ? idMatch[1] : "";
                if (!isScreeningEpoch(name)) {
                    epochs.push({ id: epochId, name: name, href: href });
                }
            }
        }

        if (epochs.length === 0) {
            showAESError("No non-screening epochs found.");
            return;
        }

        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "12px";
        popupContent.style.padding = "8px";
        popupContent.style.maxHeight = "400px";
        popupContent.style.overflowY = "auto";

        var titleDiv = document.createElement("div");
        titleDiv.style.fontSize = "16px";
        titleDiv.style.fontWeight = "600";
        titleDiv.style.marginBottom = "8px";
        titleDiv.textContent = "Select Target Epoch";
        popupContent.appendChild(titleDiv);

        var descDiv = document.createElement("div");
        descDiv.style.fontSize = "13px";
        descDiv.style.color = "#aaa";
        descDiv.style.marginBottom = "12px";
        descDiv.textContent = "Choose the epoch where you want to add the existing subject. Screening epochs are excluded.";
        popupContent.appendChild(descDiv);

        for (var j = 0; j < epochs.length; j++) {
            (function(epoch) {
                var itemDiv = document.createElement("div");
                itemDiv.style.display = "flex";
                itemDiv.style.justifyContent = "space-between";
                itemDiv.style.alignItems = "center";
                itemDiv.style.padding = "10px 12px";
                itemDiv.style.background = "#1a1a1a";
                itemDiv.style.borderRadius = "6px";
                itemDiv.style.border = "1px solid #333";

                var nameSpan = document.createElement("span");
                nameSpan.textContent = epoch.name;
                nameSpan.style.fontWeight = "500";

                var confirmBtn = document.createElement("button");
                confirmBtn.textContent = "Confirm";
                confirmBtn.style.background = "#2e7d32";
                confirmBtn.style.color = "#fff";
                confirmBtn.style.border = "none";
                confirmBtn.style.borderRadius = "4px";
                confirmBtn.style.padding = "6px 12px";
                confirmBtn.style.cursor = "pointer";
                confirmBtn.style.fontWeight = "500";
                confirmBtn.addEventListener("click", async function() {
                    log("AES: epoch selected - " + epoch.name + " (id=" + epoch.id + ")");
                    localStorage.setItem(STORAGE_AES_SELECTED_EPOCH, epoch.id);
                    localStorage.setItem(STORAGE_AES_SELECTED_EPOCH_NAME, epoch.name);
                    // Remove popup without triggering onClose cleanup
                    if (ADD_EXISTING_SUBJECT_POPUP_REF && ADD_EXISTING_SUBJECT_POPUP_REF.element) {
                        ADD_EXISTING_SUBJECT_POPUP_REF.element.remove();
                        ADD_EXISTING_SUBJECT_POPUP_REF = null;
                    }
                    await collectAllSubjectsFromEpochs();
                });

                itemDiv.appendChild(nameSpan);
                itemDiv.appendChild(confirmBtn);
                popupContent.appendChild(itemDiv);
            })(epochs[j]);
        }

        ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
            title: "Add Existing Subject",
            content: popupContent,
            width: "450px",
            height: "auto",
            maxHeight: "500px",
            onClose: function() {
                log("AES: cancelled by user");
                clearAddExistingSubjectData();
                clearRunMode();
            }
        });
    }

    async function collectAllSubjectsFromEpochs() {
        log("AES: collecting subjects from all epochs");

        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "16px";
        popupContent.style.padding = "8px";
        popupContent.style.textAlign = "center";

        var statusDiv = document.createElement("div");
        statusDiv.id = "aesCollectStatus";
        statusDiv.style.fontSize = "16px";
        statusDiv.style.fontWeight = "500";
        statusDiv.textContent = "Collecting subjects from all epochs...";
        popupContent.appendChild(statusDiv);

        var progressDiv = document.createElement("div");
        progressDiv.id = "aesCollectProgress";
        progressDiv.style.fontSize = "14px";
        progressDiv.style.color = "#9df";
        progressDiv.textContent = "Initializing...";
        popupContent.appendChild(progressDiv);

        var loadingDiv = document.createElement("div");
        loadingDiv.id = "aesLoading";
        loadingDiv.style.fontSize = "14px";
        loadingDiv.style.color = "#9df";
        loadingDiv.textContent = "Running.";
        popupContent.appendChild(loadingDiv);

        ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
            title: "Add Existing Subject - Collecting Data",
            content: popupContent,
            width: "450px",
            height: "auto",
            onClose: function() {
                log("AES: cancelled during collection");
                clearAddExistingSubjectData();
                clearRunMode();
            }
        });

        var dots = 1;
        var loadingInterval = setInterval(function() {
            if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
                clearInterval(loadingInterval);
                return;
            }
            dots = (dots % 3) + 1;
            var text = "Running";
            for (var d = 0; d < dots; d++) text += ".";
            var loadEl = document.getElementById("aesLoading");
            if (loadEl) loadEl.textContent = text;
        }, 500);

        setRunMode(RUNMODE_ADD_EXISTING_SUBJECT);
        localStorage.setItem(STORAGE_AES_POPUP, "1");
        setAESStep("collecting");

        var epochTableBody = document.getElementById("epochTableBody");
        if (!epochTableBody) {
            clearInterval(loadingInterval);
            showAESError("Could not find epoch list.");
            return;
        }

        var rows = epochTableBody.querySelectorAll("tr");
        var allEpochs = [];
        for (var i = 0; i < rows.length; i++) {
            var anchor = rows[i].querySelector("td a[href*='/secure/administration/studies/epoch/show/']");
            if (anchor) {
                var name = (anchor.textContent || "").trim();
                var href = anchor.getAttribute("href") || "";
                var idMatch = href.match(/\/show\/(\d+)/);
                var epochId = idMatch ? idMatch[1] : "";
                allEpochs.push({ id: epochId, name: name, href: href });
            }
        }

        var collectedData = {};
        var selectedEpochId = localStorage.getItem(STORAGE_AES_SELECTED_EPOCH) || "";
        var selectedEpochVolunteerIds = {};

        for (var eIdx = 0; eIdx < allEpochs.length; eIdx++) {
            var epoch = allEpochs[eIdx];
            var progEl = document.getElementById("aesCollectProgress");
            if (progEl) {
                progEl.textContent = "Processing epoch " + (eIdx + 1) + "/" + allEpochs.length + ": " + epoch.name;
            }
            setAESProgress(eIdx + 1, allEpochs.length, "Processing epoch: " + epoch.name);

            var epochData = await collectSubjectsFromEpoch(epoch);
            collectedData[epoch.id] = {
                name: epoch.name,
                cohorts: epochData
            };

            if (epoch.id === selectedEpochId) {
                for (var cohortId in epochData) {
                    var assignments = epochData[cohortId].assignments || [];
                    for (var aIdx = 0; aIdx < assignments.length; aIdx++) {
                        if (assignments[aIdx].volunteerId) {
                            selectedEpochVolunteerIds[assignments[aIdx].volunteerId] = true;
                        }
                    }
                }
            }
        }

        clearInterval(loadingInterval);
        localStorage.setItem(STORAGE_AES_COLLECTED_DATA, JSON.stringify(collectedData));
        log("AES: collection complete, showing subject selection");

        await showSubjectSelectionPopup(collectedData, selectedEpochVolunteerIds);
    }

    async function collectSubjectsFromEpoch(epoch) {
        log("AES: collecting from epoch: " + epoch.name);
        var epochData = {};

        var epochUrl = location.origin + epoch.href;
        var epochDoc = await fetchPageContent(epochUrl);
        if (!epochDoc) {
            log("AES: could not fetch epoch page: " + epoch.name);
            return epochData;
        }

        var cohortListBody = epochDoc.getElementById("cohortListBody");
        if (!cohortListBody) {
            log("AES: no cohortListBody in epoch: " + epoch.name);
            return epochData;
        }

        var cohortRows = cohortListBody.querySelectorAll("tr");
        for (var cIdx = 0; cIdx < cohortRows.length; cIdx++) {
            var cohortAnchor = cohortRows[cIdx].querySelector("td a[href*='/secure/administration/studies/cohort/show/']");
            if (!cohortAnchor) continue;

            var cohortName = (cohortAnchor.textContent || "").trim();
            var cohortHref = cohortAnchor.getAttribute("href") || "";
            var cohortIdMatch = cohortHref.match(/\/show\/(\d+)/);
            var cohortId = cohortIdMatch ? cohortIdMatch[1] : "";

            var assignments = await collectAssignmentsFromCohort(cohortHref, cohortName);
            epochData[cohortId] = {
                name: cohortName,
                href: cohortHref,
                assignments: assignments
            };
        }

        return epochData;
    }

    async function collectAssignmentsFromCohort(cohortHref, cohortName) {
        log("AES: collecting assignments from cohort: " + cohortName);
        var assignments = [];

        var cohortUrl = location.origin + cohortHref;
        var cohortDoc = await fetchPageContent(cohortUrl);
        if (!cohortDoc) {
            log("AES: could not fetch cohort page: " + cohortName);
            return assignments;
        }

        var assignmentListBody = cohortDoc.getElementById("cohortAssignmentListBody");
        if (!assignmentListBody) {
            log("AES: no cohortAssignmentListBody in cohort: " + cohortName);
            return assignments;
        }

        var rows = assignmentListBody.querySelectorAll("tr.cohortAssignmentRow");
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var rowId = row.id || "";
            var assignmentIdMatch = rowId.match(/ca_(\d+)/);
            var assignmentId = assignmentIdMatch ? assignmentIdMatch[1] : "";

            var subjectSpan = row.querySelector("span.tooltips[data-original-title='Subject Number']");
            var subjectNumber = subjectSpan ? (subjectSpan.textContent || "").trim() : "";

            var volunteerAnchor = row.querySelector("a[href*='/secure/volunteers/manage/show/']");
            var volunteerId = "";
            var volunteerDisplay = "";
            if (volunteerAnchor) {
                volunteerId = extractVolunteerIdFromHref(volunteerAnchor.getAttribute("href"));
                var volunteerText = (volunteerAnchor.textContent || "").trim();
                var volunteerInfo = parseVolunteerInfo(volunteerText);
                volunteerDisplay = formatVolunteerDisplay(volunteerInfo);
            }

            var isActive = true;
            var innerTable = row.querySelector("table.clinSparkInnerTable");
            if (innerTable) {
                var tds = innerTable.querySelectorAll("td");
                for (var tdIdx = 0; tdIdx < tds.length; tdIdx++) {
                    var tdText = (tds[tdIdx].textContent || "").trim();
                    if (tdText === "Active?:") {
                        var nextTd = tds[tdIdx + 1];
                        if (nextTd) {
                            var activeValue = (nextTd.textContent || "").trim();
                            if (activeValue !== "Yes") {
                                isActive = false;
                            }
                        }
                        break;
                    }
                }
            }

            if (!isActive) {
                log("AES: skipping inactive subject: " + subjectNumber);
                continue;
            }

            if (subjectNumber && volunteerId) {
                assignments.push({
                    assignmentId: assignmentId,
                    subjectNumber: subjectNumber,
                    volunteerId: volunteerId,
                    volunteerDisplay: volunteerDisplay
                });
            }
        }

        log("AES: found " + assignments.length + " assignments in cohort: " + cohortName);
        return assignments;
    }

    async function fetchPageContent(url) {
        return new Promise(function(resolve) {
            var iframe = document.createElement("iframe");
            iframe.style.position = "fixed";
            iframe.style.top = "-9999px";
            iframe.style.left = "-9999px";
            iframe.style.width = "1px";
            iframe.style.height = "1px";
            iframe.style.visibility = "hidden";

            var timeout = setTimeout(function() {
                if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                resolve(null);
            }, 15000);

            iframe.onload = function() {
                clearTimeout(timeout);
                try {
                    var doc = iframe.contentDocument || iframe.contentWindow.document;
                    setTimeout(function() {
                        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                        resolve(doc);
                    }, 500);
                } catch (e) {
                    if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                    resolve(null);
                }
            };

            iframe.src = url;
            document.body.appendChild(iframe);
        });
    }

    async function showSubjectSelectionPopup(collectedData, excludeVolunteerIds) {
        log("AES: showing subject selection popup");

        var selectedEpochId = localStorage.getItem(STORAGE_AES_SELECTED_EPOCH) || "";
        var selectedEpochName = localStorage.getItem(STORAGE_AES_SELECTED_EPOCH_NAME) || "";

        var subjectList = [];
        var seenVolunteerIds = {};

        for (var epochId in collectedData) {
            if (epochId === selectedEpochId) continue;
            var epochInfo = collectedData[epochId];
            for (var cohortId in epochInfo.cohorts) {
                var cohortInfo = epochInfo.cohorts[cohortId];
                var assignments = cohortInfo.assignments || [];
                for (var i = 0; i < assignments.length; i++) {
                    var assignment = assignments[i];
                    if (excludeVolunteerIds[assignment.volunteerId]) continue;
                    if (seenVolunteerIds[assignment.volunteerId]) continue;
                    seenVolunteerIds[assignment.volunteerId] = true;
                    subjectList.push({
                        subjectNumber: assignment.subjectNumber,
                        volunteerId: assignment.volunteerId,
                        volunteerDisplay: assignment.volunteerDisplay,
                        epochName: epochInfo.name,
                        cohortName: cohortInfo.name
                    });
                }
            }
        }

        if (subjectList.length === 0) {
            showAESError("No eligible subjects found in other epochs.");
            return;
        }

        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "12px";
        popupContent.style.padding = "8px";
        popupContent.style.maxHeight = "450px";
        popupContent.style.overflowY = "auto";

        var titleDiv = document.createElement("div");
        titleDiv.style.fontSize = "16px";
        titleDiv.style.fontWeight = "600";
        titleDiv.style.marginBottom = "8px";
        titleDiv.textContent = "Select Subject to Add";
        popupContent.appendChild(titleDiv);

        var descDiv = document.createElement("div");
        descDiv.style.fontSize = "13px";
        descDiv.style.color = "#aaa";
        descDiv.style.marginBottom = "12px";
        descDiv.innerHTML = "Adding to epoch: <strong>" + selectedEpochName + "</strong><br>Found " + subjectList.length + " eligible subjects.";
        popupContent.appendChild(descDiv);

        for (var j = 0; j < subjectList.length; j++) {
            (function(subject) {
                var itemDiv = document.createElement("div");
                itemDiv.style.display = "flex";
                itemDiv.style.flexDirection = "column";
                itemDiv.style.padding = "10px 12px";
                itemDiv.style.background = "#1a1a1a";
                itemDiv.style.borderRadius = "6px";
                itemDiv.style.border = "1px solid #333";
                itemDiv.style.gap = "6px";

                var topRow = document.createElement("div");
                topRow.style.display = "flex";
                topRow.style.justifyContent = "space-between";
                topRow.style.alignItems = "center";

                var subjectSpan = document.createElement("span");
                subjectSpan.textContent = subject.subjectNumber;
                subjectSpan.style.fontWeight = "600";
                subjectSpan.style.fontSize = "15px";

                var selectBtn = document.createElement("button");
                selectBtn.textContent = "Select";
                selectBtn.style.background = "#1976d2";
                selectBtn.style.color = "#fff";
                selectBtn.style.border = "none";
                selectBtn.style.borderRadius = "4px";
                selectBtn.style.padding = "6px 12px";
                selectBtn.style.cursor = "pointer";
                selectBtn.style.fontWeight = "500";
                selectBtn.addEventListener("click", async function() {
                    log("AES: subject selected - " + subject.subjectNumber + " (volunteerId=" + subject.volunteerId + ")");
                    localStorage.setItem(STORAGE_AES_SELECTED_SUBJECT, subject.subjectNumber);
                    localStorage.setItem(STORAGE_AES_SELECTED_VOLUNTEER_ID, subject.volunteerId);
                    // Remove popup without triggering onClose cleanup
                    if (ADD_EXISTING_SUBJECT_POPUP_REF && ADD_EXISTING_SUBJECT_POPUP_REF.element) {
                        ADD_EXISTING_SUBJECT_POPUP_REF.element.remove();
                        ADD_EXISTING_SUBJECT_POPUP_REF = null;
                    }
                    await startCohortAdditionProcess();
                });

                topRow.appendChild(subjectSpan);
                topRow.appendChild(selectBtn);

                var locationSpan = document.createElement("span");
                locationSpan.textContent = subject.epochName + " → " + subject.cohortName;
                locationSpan.style.fontSize = "12px";
                locationSpan.style.color = "#888";

                var volunteerSpan = document.createElement("span");
                volunteerSpan.textContent = subject.volunteerDisplay;
                volunteerSpan.style.fontSize = "11px";
                volunteerSpan.style.color = "#666";

                itemDiv.appendChild(topRow);
                itemDiv.appendChild(locationSpan);
                itemDiv.appendChild(volunteerSpan);
                popupContent.appendChild(itemDiv);
            })(subjectList[j]);
        }

        ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
            title: "Add Existing Subject - Select Subject",
            content: popupContent,
            width: "500px",
            height: "auto",
            maxHeight: "550px",
            onClose: function() {
                log("AES: cancelled during subject selection");
                clearAddExistingSubjectData();
                clearRunMode();
            }
        });
    }

    async function startCohortAdditionProcess() {
        log("AES: starting cohort addition process");
        setAESStep("navigateToEpoch");

        var selectedEpochId = localStorage.getItem(STORAGE_AES_SELECTED_EPOCH) || "";
        var epochUrl = "/secure/administration/studies/epoch/show/" + selectedEpochId;

        showAESProgressPopup("Navigating to target epoch...");
        location.href = epochUrl;
    }

    function showAESProgressPopup(message) {
        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "16px";
        popupContent.style.padding = "8px";
        popupContent.style.textAlign = "center";

        var statusDiv = document.createElement("div");
        statusDiv.id = "aesProgressStatus";
        statusDiv.style.fontSize = "16px";
        statusDiv.style.fontWeight = "500";
        statusDiv.textContent = message;
        popupContent.appendChild(statusDiv);

        var loadingDiv = document.createElement("div");
        loadingDiv.id = "aesProgressLoading";
        loadingDiv.style.fontSize = "14px";
        loadingDiv.style.color = "#9df";
        loadingDiv.textContent = "Running.";
        popupContent.appendChild(loadingDiv);

        ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
            title: "Add Existing Subject",
            content: popupContent,
            width: "400px",
            height: "auto",
            onClose: function() {
                log("AES: cancelled during progress");
                clearAddExistingSubjectData();
                clearRunMode();
            }
        });

        var dots = 1;
        setInterval(function() {
            if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) return;
            dots = (dots % 3) + 1;
            var text = "Running";
            for (var d = 0; d < dots; d++) text += ".";
            var loadEl = document.getElementById("aesProgressLoading");
            if (loadEl) loadEl.textContent = text;
        }, 500);
    }

    function updateAESProgressStatus(message) {
        var statusEl = document.getElementById("aesProgressStatus");
        if (statusEl) statusEl.textContent = message;
    }

    function showAESError(message) {
        log("AES Error: " + message);
        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "16px";
        popupContent.style.padding = "8px";
        popupContent.style.textAlign = "center";

        var errorDiv = document.createElement("div");
        errorDiv.style.fontSize = "16px";
        errorDiv.style.color = "#f44336";
        errorDiv.style.fontWeight = "500";
        errorDiv.textContent = message;
        popupContent.appendChild(errorDiv);

        ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
            title: "Add Existing Subject - Error",
            content: popupContent,
            width: "400px",
            height: "auto",
            onClose: function() {
                clearAddExistingSubjectData();
                clearRunMode();
            }
        });
    }

    function showAESComplete() {
        log("AES: process complete");
        var popupContent = document.createElement("div");
        popupContent.style.display = "flex";
        popupContent.style.flexDirection = "column";
        popupContent.style.gap = "16px";
        popupContent.style.padding = "8px";
        popupContent.style.textAlign = "center";

        var successDiv = document.createElement("div");
        successDiv.style.fontSize = "18px";
        successDiv.style.color = "#4caf50";
        successDiv.style.fontWeight = "600";
        successDiv.textContent = "✓ Subject Added Successfully!";
        popupContent.appendChild(successDiv);

        var subjectNumber = localStorage.getItem(STORAGE_AES_SELECTED_SUBJECT) || "";
        var epochName = localStorage.getItem(STORAGE_AES_SELECTED_EPOCH_NAME) || "";

        var detailsDiv = document.createElement("div");
        detailsDiv.style.fontSize = "14px";
        detailsDiv.style.color = "#aaa";
        detailsDiv.innerHTML = "Subject <strong>" + subjectNumber + "</strong> has been added to epoch <strong>" + epochName + "</strong> and activated.";
        popupContent.appendChild(detailsDiv);

        // Transition existing popup instead of creating new one
        if (ADD_EXISTING_SUBJECT_POPUP_REF && ADD_EXISTING_SUBJECT_POPUP_REF.element && document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
            ADD_EXISTING_SUBJECT_POPUP_REF.setTitle("Add Existing Subject - Complete");
            ADD_EXISTING_SUBJECT_POPUP_REF.setContent(popupContent);
        } else {
            ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                title: "Add Existing Subject - Complete",
                content: popupContent,
                width: "400px",
                height: "auto",
                onClose: function() {
                    clearAddExistingSubjectData();
                    clearRunMode();
                }
            });
        }
    }

    async function processAESOnPageLoad() {
        var runMode = getRunMode();
        if (runMode !== RUNMODE_ADD_EXISTING_SUBJECT) return;

        var step = getAESStep();
        log("AES: processing on page load, step=" + step);

        if (step === "selectEpoch") {
            await showEpochSelectionPopup();
            return;
        }

        if (step === "collecting") {
            showAESProgressPopup("Data collection was interrupted. Please restart.");
            return;
        }

        if (step === "navigateToEpoch") {
            await processEpochPage();
            return;
        }

        if (step === "cohortPage") {
            await processCohortPage();
            return;
        }

        if (step === "afterCohortEdit") {
            await processAfterCohortEdit();
            return;
        }

        if (step === "afterAddSubject") {
            await processAfterAddSubject();
            return;
        }

        if (step === "activatePlan") {
            await processActivatePlan();
            return;
        }

        if (step === "activateVolunteer") {
            await processActivateVolunteer();
            return;
        }
    }

    async function processEpochPage() {
        log("AES: processing epoch page");
        showAESProgressPopup("Navigating to first cohort...");

        var cohortListBody = await waitForSelectorWithRetry("#cohortListBody", 5000, 3);
        if (!cohortListBody) {
            showAESError("Could not find cohort list in epoch.");
            return;
        }

        var firstCohortAnchor = cohortListBody.querySelector("td a[href*='/secure/administration/studies/cohort/show/']");
        if (!firstCohortAnchor) {
            showAESError("No cohort found in the selected epoch.");
            return;
        }

        var cohortHref = firstCohortAnchor.getAttribute("href") || "";
        var cohortIdMatch = cohortHref.match(/\/show\/(\d+)/);
        var cohortId = cohortIdMatch ? cohortIdMatch[1] : "";
        localStorage.setItem(STORAGE_AES_COHORT_ID, cohortId);
        setAESStep("cohortPage");

        location.href = cohortHref;
    }

    async function processCohortPage() {
        log("AES: processing cohort page");
        var cohortEditDone = localStorage.getItem(STORAGE_AES_COHORT_EDIT_DONE);

        if (cohortEditDone !== "1") {
            showAESProgressPopup("Configuring cohort settings...");
            await editCohortSettings();
        } else {
            await addSubjectToCohort();
        }
    }

    async function editCohortSettings() {
        log("AES: editing cohort settings");
        updateAESProgressStatus("Opening cohort edit modal...");

        var actionsBtn = await waitForSelector("button.dropdown-toggle[data-toggle='dropdown']", 5000);
        if (!actionsBtn) {
            var allBtns = document.querySelectorAll("button.dropdown-toggle");
            for (var i = 0; i < allBtns.length; i++) {
                if ((allBtns[i].textContent || "").toLowerCase().indexOf("actions") !== -1) {
                    actionsBtn = allBtns[i];
                    break;
                }
            }
        }

        if (!actionsBtn) {
            showAESError("Could not find Actions button.");
            return;
        }

        actionsBtn.click();
        await sleep(500);

        var editLink = document.querySelector("a[href*='/secure/administration/manage/studies/cohort/update/'][data-toggle='modal']");
        if (!editLink) {
            showAESError("Could not find Edit link in Actions menu.");
            return;
        }

        editLink.click();

        var modal = await waitForSelectorWithRetry("#ajaxModal", 10000, 3);
        if (!modal) {
            showAESError("Edit modal did not open.");
            return;
        }

        await sleep(1500);
        updateAESProgressStatus("Configuring checkboxes...");

        setCheckboxStateById("subjectInitiation", false);
        setCheckboxStateById("sourceVolunteerDatabase", false);
        setCheckboxStateById("sourceAppointments", false);
        setCheckboxStateById("sourceAppointmentsCohort", false);
        setCheckboxStateById("sourceScreeningCohorts", true);
        setCheckboxStateById("sourceLeadInCohorts", true);
        setCheckboxStateById("sourceRandomizationCohorts", true);
        setCheckboxStateById("allowSubjectsActiveInCohorts", true);
        setCheckboxStateById("allowSubjectsActiveInStudies", true);
        setCheckboxStateById("requireVolunteerRecruitment", false);
        setCheckboxStateById("allowRecruitmentEligible", false);
        setCheckboxStateById("allowRecruitmentIdentified", false);
        setCheckboxStateById("allowRecruitmentIneligible", false);
        setCheckboxStateById("allowRecruitmentRemoved", false);
        setCheckboxStateById("allowEligibilityEligible", true);
        setCheckboxStateById("allowEligibilityPending", true);
        setCheckboxStateById("allowEligibilityIneligible", true);
        setCheckboxStateById("allowEligibilityUnspecified", true);
        setCheckboxStateById("allowStatusActive", true);
        setCheckboxStateById("allowStatusComplete", false);
        setCheckboxStateById("allowStatusTerminated", false);
        setCheckboxStateById("allowStatusWithdrawn", false);
        setCheckboxStateById("requireInformedConsent", false);
        setCheckboxStateById("requireOverVolunteeringCheck", false);

        var reason = document.querySelector("textarea#reasonForChange");
        if (reason) {
            reason.value = "test";
            var evt = new Event("input", { bubbles: true });
            reason.dispatchEvent(evt);
        }

        await sleep(500);
        updateAESProgressStatus("Saving cohort settings...");

        var saveBtn = document.getElementById("actionButton");
        if (!saveBtn) {
            showAESError("Could not find Save button in modal.");
            return;
        }

        localStorage.setItem(STORAGE_AES_COHORT_EDIT_DONE, "1");
        setAESStep("afterCohortEdit");
        saveBtn.click();
    }

    async function processAfterCohortEdit() {
        log("AES: processing after cohort edit");
        await sleep(2000);
        await addSubjectToCohort();
    }

    async function addSubjectToCohort() {
        log("AES: adding subject to cohort");
        showAESProgressPopup("Opening Add Subject modal...");

        var assignmentsActionBtn = document.getElementById("assignmentsActionMenu");
        if (!assignmentsActionBtn) {
            showAESError("Could not find assignments Action button.");
            return;
        }

        assignmentsActionBtn.click();
        await sleep(500);

        var addLink = document.getElementById("addCohortAssignmentButton");
        if (!addLink) {
            addLink = document.querySelector("a[id='addCohortAssignmentButton']");
        }
        if (!addLink) {
            showAESError("Could not find Add button.");
            return;
        }

        addLink.click();

        var modal = await waitForSelectorWithRetry("#ajaxModal", 10000, 3);
        if (!modal) {
            showAESError("Add Subject modal did not open.");
            return;
        }

        await sleep(1500);
        updateAESProgressStatus("Selecting activity plan...");

        var activityPlanSelect = document.getElementById("activityPlan");
        if (activityPlanSelect) {
            var options = activityPlanSelect.querySelectorAll("option");
            for (var i = 0; i < options.length; i++) {
                var val = (options[i].value || "").trim();
                if (val.length > 0) {
                    activityPlanSelect.value = val;
                    var evt = new Event("change", { bubbles: true });
                    activityPlanSelect.dispatchEvent(evt);
                    log("AES: selected activity plan index " + i);
                    break;
                }
            }
        }

        await sleep(1000);

        var initialRefInput = document.getElementById("initialSegmentReference");
        if (initialRefInput && initialRefInput.offsetParent !== null) {
            var now = new Date();
            var formatted = formatDateForClinSpark(now);
            initialRefInput.value = formatted;
            var evt2 = new Event("change", { bubbles: true });
            initialRefInput.dispatchEvent(evt2);
            log("AES: set initial reference time to " + formatted);
        }

        await sleep(500);
        updateAESProgressStatus("Searching for subject...");

        var subjectNumber = localStorage.getItem(STORAGE_AES_SELECTED_SUBJECT) || "";
        log("AES: searching for subject number: " + subjectNumber);

        var s2container = modal.querySelector('#s2id_volunteer');
        if (!s2container) {
            s2container = modal.querySelector('.select2-container.form-control.select2');
        }
        if (!s2container) {
            log("AES: Select2 container not found");
            showAESError("Could not find volunteer select2 container.");
            return;
        }

        var s2choice = s2container.querySelector('a.select2-choice');
        if (s2choice) {
            s2choice.click();
            log("AES: clicked select2 choice to open dropdown");
            await sleep(150);
        }

        var focusser = s2container.querySelector('input.select2-focusser');
        if (focusser) {
            focusser.focus();
            var kd = new KeyboardEvent("keydown", { bubbles: true, key: "ArrowDown", keyCode: 40 });
            focusser.dispatchEvent(kd);
            focusser.click();
            await sleep(150);
        }

        var s2drop = document.querySelector('#select2-drop.select2-drop-active');
        if (!s2drop) {
            s2drop = s2container.querySelector('.select2-drop');
        }
        if (!s2drop) {
            s2choice = s2container.querySelector('a.select2-choice');
            if (s2choice) {
                s2choice.click();
            }
            s2drop = await waitForSelector('#select2-drop.select2-drop-active', 2000);
            if (!s2drop) {
                s2drop = s2container.querySelector('.select2-drop');
            }
        }
        if (!s2drop) {
            log("AES: Select2 drop not found");
            showAESError("Could not open volunteer dropdown.");
            return;
        }

        var s2input = s2drop.querySelector('input.select2-input');
        if (!s2input) {
            s2input = await waitForSelector('#select2-drop.select2-drop-active input.select2-input', 2000);
        }
        if (!s2input) {
            s2input = s2container.querySelector('input.select2-input');
        }
        if (!s2input) {
            log("AES: Select2 input not found");
            showAESError("Could not find volunteer search input.");
            return;
        }

        log("AES: found search input: " + (s2input.id || "no-id"));

        s2input.value = subjectNumber;
        var inpEvt = new Event("input", { bubbles: true });
        s2input.dispatchEvent(inpEvt);
        var keyEvt = new KeyboardEvent("keyup", { bubbles: true, key: subjectNumber, keyCode: subjectNumber.charCodeAt(0) });
        s2input.dispatchEvent(keyEvt);
        log("AES: typed subject number: " + subjectNumber);

        var selectionConfirmed = false;
        var confirmWait = 0;
        var confirmMax = 12000;
        while (confirmWait < confirmMax) {
            var enterDown = new KeyboardEvent("keydown", { bubbles: true, key: "Enter", keyCode: 13 });
            s2input.dispatchEvent(enterDown);
            var enterUp = new KeyboardEvent("keyup", { bubbles: true, key: "Enter", keyCode: 13 });
            s2input.dispatchEvent(enterUp);
            await sleep(400);

            var containerClass = s2container.getAttribute("class") + "";
            var hasAllow = containerClass.indexOf("select2-allowclear") !== -1;
            var notOpen = containerClass.indexOf("select2-dropdown-open") === -1;
            var chosenEl = null;
            chosenEl = s2container.querySelector('.select2-chosen');
            if (!chosenEl) {
                chosenEl = s2container.querySelector('[id^="select2-chosen-"]');
            }
            var chosenText = "";
            if (chosenEl) {
                chosenText = (chosenEl.textContent + "").trim();
            }
            var notSearch = chosenText.trim().toLowerCase() !== "search";

            if (hasAllow && notOpen && notSearch) {
                selectionConfirmed = true;
                log("AES: selection confirmed; chosenText=" + chosenText);
                break;
            }
            confirmWait = confirmWait + 400;
        }

        if (!selectionConfirmed) {
            log("AES: selection not confirmed after " + confirmMax + "ms");
            showAESError("Could not confirm volunteer selection for subject " + subjectNumber);
            return;
        }

        log("AES: volunteer selected successfully");
        setAESStep("afterAddSubject");

        var saveBtn = document.getElementById("actionButton");
        if (saveBtn) {
            saveBtn.click();
        }
    }

    function formatDateForClinSpark(date) {
        var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        var day = String(date.getDate()).padStart(2, "0");
        var month = months[date.getMonth()];
        var year = date.getFullYear();
        var hours = String(date.getHours()).padStart(2, "0");
        var mins = String(date.getMinutes()).padStart(2, "0");
        var secs = String(date.getSeconds()).padStart(2, "0");
        return day + month + year + " " + hours + ":" + mins + ":" + secs + " PST";
    }

    async function processAfterAddSubject() {
        log("AES: processing after add subject");
        await sleep(2000);

        var alertDanger = document.querySelector("div.alert.alert-danger");
        if (alertDanger) {
            showAESError("Error encountered. Stopping program.");
            return;
        }

        updateAESProgressStatus("Locating new assignment...");

        var volunteerId = localStorage.getItem(STORAGE_AES_SELECTED_VOLUNTEER_ID) || "";
        var subjectNumber = localStorage.getItem(STORAGE_AES_SELECTED_SUBJECT) || "";

        var assignmentListBody = await waitForSelectorWithRetry("#cohortAssignmentListBody", 5000, 3);
        if (!assignmentListBody) {
            showAESError("Could not find assignment list after save.");
            return;
        }

        var rows = assignmentListBody.querySelectorAll("tr.cohortAssignmentRow");
        var targetRow = null;
        var assignmentId = "";

        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var volunteerAnchor = row.querySelector("a[href*='/secure/volunteers/manage/show/" + volunteerId + "']");
            if (volunteerAnchor) {
                // Also verify this is the correct subject number
                var subjectAnchor = row.querySelector("a[href*='/secure/study/subjects/show/']");
                if (subjectAnchor) {
                    var subjectText = (subjectAnchor.textContent || "").trim();
                    // Check if subject number matches (could be "111-1005" or "1003 / 111-1005")
                    if (subjectText.indexOf(subjectNumber) !== -1) {
                        targetRow = row;
                        var rowId = row.id || "";
                        var match = rowId.match(/ca_(\d+)/);
                        if (match) assignmentId = match[1];
                        log("AES: found matching assignment - volunteerId=" + volunteerId + ", subjectNumber=" + subjectNumber + ", assignmentId=" + assignmentId);
                        break;
                    }
                }
            }
        }

        if (!targetRow || !assignmentId) {
            showAESError("Could not find the newly added assignment.");
            return;
        }

        localStorage.setItem(STORAGE_AES_ASSIGNMENT_ID, assignmentId);
        log("AES: found assignment id=" + assignmentId);

        await processActivatePlan();
    }

    async function processActivatePlan() {
        log("AES: processing activate plan");
        var assignmentId = localStorage.getItem(STORAGE_AES_ASSIGNMENT_ID) || "";

        updateAESProgressStatus("Activating plan...");

        var activatePlanLink = document.getElementById("cohortAssignmentActivatePlanLink_" + assignmentId);

        if (!activatePlanLink) {
            activatePlanLink = document.querySelector("a[onclick*='activatePlan(" + assignmentId + ")']");
        }

        if (!activatePlanLink) {
            var row = document.getElementById("ca_" + assignmentId);
            if (row) {
                var actionBtn = row.querySelector("button.dropdown-toggle");
                if (actionBtn) {
                    actionBtn.click();
                    await sleep(500);
                    activatePlanLink = document.querySelector("a[onclick*='activatePlan(" + assignmentId + ")']");
                }
            }
        }

        if (activatePlanLink) {
            var isDisabled = activatePlanLink.closest("li.disabled") !== null;
            if (isDisabled) {
                log("AES: Activate Plan is disabled, treating as already activated");
                setAESStep("activateVolunteer");
                await processActivateVolunteer();
                return;
            }

            setAESStep("activateVolunteer");
            activatePlanLink.click();
            log("AES: clicked Activate Plan");

            // Wait for and click OK button in confirmation modal
            await sleep(500);
            var okBtn = await waitForSelector('button[data-bb-handler="confirm"].btn.btn-primary', 3000);
            if (okBtn) {
                okBtn.click();
                log("AES: clicked OK on Activate Plan confirmation");
                await sleep(1000);
            } else {
                log("AES: OK button not found after Activate Plan");
            }

            await processActivateVolunteer();
        } else {
            log("AES: Activate Plan link not found, may already be activated");
            setAESStep("activateVolunteer");
            await processActivateVolunteer();
        }
    }

    async function processActivateVolunteer() {
        log("AES: processing activate volunteer");
        var assignmentId = localStorage.getItem(STORAGE_AES_ASSIGNMENT_ID) || "";

        updateAESProgressStatus("Activating volunteer...");

        var maxWait = 30000;
        var waited = 0;
        var activateVolunteerLink = null;

        while (waited < maxWait) {
            activateVolunteerLink = document.getElementById("cohortAssignmentActivateSubjectLink_" + assignmentId);
            if (!activateVolunteerLink) {
                activateVolunteerLink = document.querySelector("a[onclick*='activateVolunteer(" + assignmentId + ")']");
            }

            if (activateVolunteerLink) {
                var isDisabled = activateVolunteerLink.closest("li.disabled") !== null;
                if (!isDisabled) {
                    break;
                }
            }

            await sleep(1000);
            waited += 1000;

            var row = document.getElementById("ca_" + assignmentId);
            if (row) {
                var actionBtn = row.querySelector("button.dropdown-toggle");
                if (actionBtn) {
                    actionBtn.click();
                    await sleep(300);
                }
            }
        }

        if (activateVolunteerLink) {
            var isDisabled = activateVolunteerLink.closest("li.disabled") !== null;
            if (isDisabled) {
                log("AES: Both Activate Plan and Activate Volunteer are disabled - treating as complete");
                showAESComplete();
                return;
            }

            activateVolunteerLink.click();
            log("AES: clicked Activate Volunteer");

            // Wait for and click OK button in confirmation modal
            await sleep(500);
            var okBtn = await waitForSelector('button[data-bb-handler="confirm"].btn.btn-primary', 3000);
            if (okBtn) {
                okBtn.click();
                log("AES: clicked OK on Activate Volunteer confirmation");
                await sleep(1000);
            } else {
                log("AES: OK button not found after Activate Volunteer");
            }

            showAESComplete();
        } else {
            log("AES: Activate Volunteer link not found after waiting");
            showAESComplete();
        }
    }

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
    var SA_BUILDER_PROGRESS_POPUP_REF = null;
    var SA_BUILDER_CANCELLED = false;
    var SA_BUILDER_PAUSE = false;
    var SA_BUILDER_TARGET_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/activityplans/show/";

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
        btnRow.style.gap = scale(BUTTON_GAP_PX);
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

    function submitForm(url, formData) {
        return new Promise(function(resolve, reject) {
            if (typeof GM !== "undefined" && typeof GM.xmlHttpRequest === "function") {
                GM.xmlHttpRequest({
                    method: "POST",
                    url: url,
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded"
                    },
                    data: formData,
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

    //==========================
    // FIND FORM FEATURE
    //==========================
    // This section contains all functions related to finding forms.
    // This feature automates pull any subject identifier found on page,
    // request user for form keyword, and then search for form based on the keyword.
    //==========================

    function setPanelHidden(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_HIDDEN, flag ? "1" : "0");
            log("Panel hidden state set to " + String(flag));
        } catch (e) {
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
            if (n3.length > 0) {
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
                if (hasSlash) {
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
                if (nc.length > 0) {
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

        var popup = createPopup({ title: FORM_POPUP_TITLE, content: container, width: "520px", height: "auto" });

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
        var popup = createPopup({ title: PARSE_METHOD_POPUP_TITLE, content: container, width: "480px", height: "auto", onClose: function() { stopParseMethodAutomation(); } });
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
                if (table) { var tbody = table.querySelector("tbody"); if (tbody) { var rows = tbody.querySelectorAll("tr"); for (var i = 0; i < rows.length; i++) { var tds = rows[i].querySelectorAll("td"); if (tds.length >= 1) { var link = tds[0].querySelector("a"); if (link) { var name = (link.textContent || "").trim(); var href = link.getAttribute("href") || ""; if (name && href) { var url = href.indexOf("http") === 0 ? href : "https://cenexeltest.clinspark.com" + href; methods.push({ name: name, url: url }); } } } } } }
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
                for (var i = 0; i < groups.length; i++) { var items = groups[i].querySelectorAll("li.list-group-item a"); for (var j = 0; j < items.length; j++) { var href = items[j].getAttribute("href") || ""; if (href.indexOf("/show/itemgroup/") >= 0) { var fullUrl = href.indexOf("http") === 0 ? href : "https://cenexeltest.clinspark.com" + href; refs.push({ url: fullUrl }); } } }
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
    // COLLECT ALL FEATURES
    //==========================
    //==========================

    // Return true if we are on the Data Collection > Subject page.
    function isDataCollectionSubjectPage() {
        var p = location.pathname + "";
        var ok = false;
        if (p === "/secure/datacollection/subject") {
            ok = true;
        } else {
            ok = false;
        }
        if (ok) {
            log("CollectAll: detected Data Collection Subject page");
        } else {
            log("CollectAll: NOT on Data Collection Subject page; current path=" + p);
        }
        return ok;
    }

    // Find and return all collectable form rows in the table body.
    function getFormDataRows() {
        var body = document.querySelector("tbody#formDataTableBody");
        if (!body) {
            log("CollectAll: tbody#formDataTableBody not found");
            return [];
        }
        var rows = body.querySelectorAll("tr[id^='formDataRow_']");
        var out = [];
        var i = 0;
        while (i < rows.length) {
            out.push(rows[i]);
            i = i + 1;
        }
        log("CollectAll: found " + String(out.length) + " formDataRow_* rows");
        return out;
    }

    // Parse numeric form id from a <tr id="formDataRow_#####">
    function getFormDataRowId(tr) {
        if (!tr) {
            return "";
        }
        var idAttr = tr.id + "";
        var m = idAttr.match(/formDataRow_(\d+)/);
        if (m && m[1]) {
            return m[1];
        }
        return "";
    }

    // Extract form name from the first <td> of a form row
    function getFormNameFromRow(tr) {
        if (!tr) {
            return "";
        }
        var firstTd = tr.querySelector("td:first-child");
        if (!firstTd) {
            return "";
        }
        var formSpan = firstTd.querySelector("span.tooltips[data-original-title='Form']");
        if (formSpan) {
            return (formSpan.textContent + "").trim();
        }
        return "";
    }

    // Extract collection status from the <tr> element's class
    function getFormStatusFromRow(tr) {
        if (!tr) {
            return "";
        }
        var className = tr.className + "";
        // Look for dataCollectionStatus_* pattern
        var statusMatch = className.match(/dataCollectionStatus_([^\s]+)/);
        if (statusMatch && statusMatch[1]) {
            return statusMatch[1];
        }
        // Fallback: look for other status indicators
        if (className.indexOf("formDataUnscheduled") !== -1) {
            return "Unscheduled";
        }
        if (className.indexOf("formDataIncomplete") !== -1) {
            return "Incomplete";
        }
        if (className.indexOf("formDataUnsaved") !== -1) {
            return "Completed";
        }
        return "Unknown";
    }

    // Locate the "Collect" button inside a form row.
    function findCollectButtonInRow(tr) {
        if (!tr) {
            return null;
        }
        var btns = tr.querySelectorAll("button.btn.green.btn-sm");
        var i = 0;
        while (i < btns.length) {
            var b = btns[i];
            var t = (b.textContent + "").trim();
            var hasText = t.toLowerCase().indexOf("collect") !== -1;
            if (hasText) {
                return b;
            }
            i = i + 1;
        }
        var alt = tr.querySelector("button[onclick^='showModal(']");
        if (alt) {
            return alt;
        }
        return null;
    }

    // Return true if the barcode verify modal is present AND visible.
    function isBarcodeVerifyModalVisible() {
        var div = document.getElementById("requireSubjectBarcodeVerifyDiv");
        if (!div) {
            log("CollectAll: requireSubjectBarcodeVerifyDiv not present");
            return false;
        }
        var styleAttr = div.getAttribute("style") || "";
        var hidden = styleAttr.indexOf("display: none") !== -1;
        if (hidden) {
            log("CollectAll: barcode verify modal found but style=display: none (treated as not visible)");
            return false;
        }
        log("CollectAll: barcode verify modal is visible (style not 'display: none')");
        return true;
    }

    // Return true if the form order modal is present AND visible.
    function isFormOrderModalVisible() {
        var div = document.getElementById("requireFormOrderDiv");
        if (!div) {
            log("CollectAll: requireFormOrderDiv not present");
            return false;
        }
        var styleAttr = div.getAttribute("style") || "";
        var hidden = styleAttr.indexOf("display: none") !== -1;
        if (hidden) {
            log("CollectAll: form order modal found but style=display: none (treated as not visible)");
            return false;
        }
        log("CollectAll: form order modal is visible (style not 'display: none')");
        return true;
    }

    // Handle the form order modal: fill textarea and click Save.
    async function handleFormOrderModal() {
        log("CollectAll: handling form order modal");

        var textarea = await waitForSelector("textarea#ordercomment", 6000);
        if (!textarea) {
            log("CollectAll: ordercomment textarea not found in form order modal");
            return false;
        }

        log("CollectAll: found ordercomment textarea; setting value to 'Test'");
        textarea.value = "Test";
        var evt = new Event("input", { bubbles: true });
        textarea.dispatchEvent(evt);

        var saveBtn = await waitForSelector("button#saveOutOfOrderButton", 3000);
        if (!saveBtn) {
            log("CollectAll: saveOutOfOrderButton not found");
            return false;
        }

        log("CollectAll: clicking Save button in form order modal");
        saveBtn.click();

        // Wait for modal to close
        var closed = await waitForAnyModalToClose(6000);
        if (closed) {
            log("CollectAll: form order modal closed successfully");
        } else {
            log("CollectAll: form order modal close wait timed out");
        }

        return true;
    }

    // Await modal close: waits until no visible Bootstrap modal remains or timeout.
    async function waitForAnyModalToClose(timeoutMs) {
        var max = typeof timeoutMs === "number" ? timeoutMs : 12000;
        var start = Date.now();
        while (Date.now() - start < max) {
            var openModal = document.querySelector(".modal.in, .modal.show");
            if (!openModal) {
                log("CollectAll: no .modal.in/.modal.show detected; assuming closed");
                return true;
            }
            await sleep(200);
        }
        log("CollectAll: waitForAnyModalToClose reached timeout");
        return false;
    }

    // Wait for the form table to refresh (simple delay + presence check).
    async function waitForFormTableRefresh(timeoutMs) {
        var max = typeof timeoutMs === "number" ? timeoutMs : 8000;
        var start = Date.now();
        while (Date.now() - start < max) {
            var rows = getFormDataRows();
            if (rows.length >= 0) {
                return true;
            }
            await sleep(300);
        }
        log("CollectAll: waitForFormTableRefresh timeout");
        return false;
    }


    // Fetch barcode completely in the background without opening a new tab.
    // Uses the subject ID directly if available, or uses a hidden iframe to search.
    async function fetchBarcodeInBackground(subjectText, subjectId) {
        log("Background Barcode: starting search subjectText='" + String(subjectText) + "' subjectId='" + String(subjectId) + "'");

        if (subjectId && subjectId.length > 0) {
            var result = "S" + String(subjectId);
            log("Background Barcode: using direct ID, result=" + result);
            return result;
        }

        if (!subjectText || subjectText.length === 0) {
            log("Background Barcode: no subjectText or subjectId provided");
            return null;
        }

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
                    var maxRetries = 40;
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


    // Main "Collect All" feature that combines Run Barcode (if needed) + Run Form (IR) for each first row repeatedly.
    async function runCollectAll() {
        log("CollectAll: start requested");

        // Clear any previous data and reset cancellation flag
        clearCollectAllData();
        COLLECT_ALL_CANCELLED = false;

        var onPage = isDataCollectionSubjectPage();
        if (!onPage) {
            log("CollectAll: redirecting to Data Collection Subject page and aborting current run");
            location.href = DATA_COLLECTION_SUBJECT_URL;
            return;
        }

        // Create popup container with status and list
        var popupContainer = document.createElement("div");
        popupContainer.style.display = "flex";
        popupContainer.style.flexDirection = "column";
        popupContainer.style.gap = "12px";

        var statusDiv = document.createElement("div");
        statusDiv.style.textAlign = "center";
        statusDiv.style.fontSize = "16px";
        statusDiv.style.color = "#fff";
        statusDiv.style.padding = "8px";
        statusDiv.textContent = "Running.";

        var listContainer = document.createElement("div");
        listContainer.style.maxHeight = "400px";
        listContainer.style.overflowY = "auto";
        listContainer.style.border = "1px solid #444";
        listContainer.style.borderRadius = "4px";
        listContainer.style.padding = "8px";
        listContainer.style.background = "#1a1a1a";
        listContainer.style.fontSize = "13px";

        var listTitle = document.createElement("div");
        listTitle.textContent = "Collected Forms:";
        listTitle.style.fontWeight = "600";
        listTitle.style.marginBottom = "8px";
        listTitle.style.color = "#fff";
        listContainer.appendChild(listTitle);

        var formsList = document.createElement("div");
        formsList.id = "collectAllFormsList";
        formsList.style.display = "flex";
        formsList.style.flexDirection = "column";
        formsList.style.gap = "4px";
        listContainer.appendChild(formsList);

        popupContainer.appendChild(statusDiv);
        popupContainer.appendChild(listContainer);

        var pop = createPopup({
            title: "Collect All",
            content: popupContainer,
            width: "450px",
            height: "500px",
            onClose: function() {
                COLLECT_ALL_CANCELLED = true;
                clearCollectAllData();
                log("CollectAll: cancelled by user (close button)");
            }
        });

        COLLECT_ALL_POPUP_REF = pop;

        // Function to update the status text with animation
        var animDots = 1;
        var animTimer = setInterval(function () {
            if (COLLECT_ALL_CANCELLED) {
                try {
                    clearInterval(animTimer);
                } catch (e) {}
                return;
            }
            animDots = animDots + 1;
            if (animDots > 3) {
                animDots = 1;
            }
            var t = "Running";
            var i = 0;
            while (i < animDots) {
                t = t + ".";
                i = i + 1;
            }
            statusDiv.textContent = t;
        }, 500);

        // Function to add a form to the list
        function addFormToList(formName, formStatus, formId) {
            var listItem = document.createElement("div");
            listItem.style.display = "flex";
            listItem.style.justifyContent = "space-between";
            listItem.style.alignItems = "center";
            listItem.style.padding = "6px 8px";
            listItem.style.background = "#222";
            listItem.style.borderRadius = "4px";
            listItem.style.border = "1px solid #333";

            var nameSpan = document.createElement("span");
            nameSpan.textContent = formName || "Unknown Form";
            nameSpan.style.color = "#fff";
            nameSpan.style.flex = "1";

            var statusSpan = document.createElement("span");
            statusSpan.textContent = formStatus || "Unknown";
            statusSpan.style.color = "#9df";
            statusSpan.style.marginLeft = "12px";
            statusSpan.style.fontSize = "12px";
            statusSpan.style.padding = "2px 8px";
            statusSpan.style.background = "#333";
            statusSpan.style.borderRadius = "3px";

            listItem.appendChild(nameSpan);
            listItem.appendChild(statusSpan);
            formsList.appendChild(listItem);

            // Auto-scroll to bottom
            listContainer.scrollTop = listContainer.scrollHeight;
        }

        var processed = {};
        var safety = 0;
        var safetyMax = 1000;

        while (safety < safetyMax) {
            if (COLLECT_ALL_CANCELLED) {
                log("CollectAll: cancelled; stopping run");
                try {
                    clearInterval(animTimer);
                } catch (e1) {}
                statusDiv.textContent = "Cancelled";
                statusDiv.style.color = "#f99";
                return;
            }

            if (isPaused()) {
                log("CollectAll: paused detected; stopping run");
                try {
                    clearInterval(animTimer);
                } catch (e1) {}
                statusDiv.textContent = "Paused";
                statusDiv.style.color = "#ff9";
                clearCollectAllData();
                return;
            }

            var rows = getFormDataRows();
            if (!rows || rows.length === 0) {
                log("CollectAll: no form rows found; finishing");
                break;
            }

            var pickedRow = null;
            var pickedFormId = "";
            var i = 0;
            while (i < rows.length) {
                var tr = rows[i];
                var fid = getFormDataRowId(tr);
                if (fid && fid.length > 0) {
                    var already = processed[fid] === true;
                    if (!already) {
                        pickedRow = tr;
                        pickedFormId = fid;
                        break;
                    }
                }
                i = i + 1;
            }

            if (!pickedRow) {
                log("CollectAll: no new rows (all remaining are repeats); finishing");
                break;
            }

            // Extract form name and status before processing
            var formName = getFormNameFromRow(pickedRow);
            var formStatus = getFormStatusFromRow(pickedRow);

            log("CollectAll: picked first unprocessed formId=" + String(pickedFormId) + " name=" + String(formName) + " status=" + String(formStatus));
            processed[pickedFormId] = true;

            var collectBtn = findCollectButtonInRow(pickedRow);
            if (!collectBtn) {
                log("CollectAll: Collect button not found in formId=" + String(pickedFormId) + "; moving to next");
                await sleep(600);
                safety = safety + 1;
                continue;
            }

            log("CollectAll: clicking Collect for formId=" + String(pickedFormId));
            collectBtn.click();

            // Wait a moment for modals to appear
            await sleep(300);

            // Check for form order modal first (may appear before or after barcode modal)
            var formOrderDiv = await waitForSelector("#requireFormOrderDiv", 6000);
            var formOrderVisible = false;
            if (formOrderDiv) {
                formOrderVisible = isFormOrderModalVisible();
            }

            if (formOrderVisible) {
                log("CollectAll: form order modal detected; handling it");
                await handleFormOrderModal();
                // After handling form order modal, wait and check if barcode modal appears
                await sleep(500);
            }

            // Check for barcode modal (may appear before or after form order modal)
            var barcodeDiv = await waitForSelector("#requireSubjectBarcodeVerifyDiv", 6000);
            var barcodeVisible = false;
            if (barcodeDiv) {
                barcodeVisible = isBarcodeVerifyModalVisible();
            } else {
                log("CollectAll: barcode verify container not found after Collect click (may not be required)");
            }

            if (barcodeVisible) {
                log("CollectAll: barcode required; executing Run Barcode feature");
                await APS_RunBarcode();
                var okClosed = await waitForAnyModalToClose(3000);
                if (okClosed) {
                    log("CollectAll: barcode modal flow appears closed");
                } else {
                    log("CollectAll: barcode modal close wait timed out; continuing cautiously");
                }
                // After barcode modal closes, check again for form order modal (in case it appears after)
                await sleep(500);
                formOrderDiv = document.getElementById("requireFormOrderDiv");
                if (formOrderDiv) {
                    formOrderVisible = isFormOrderModalVisible();
                    if (formOrderVisible) {
                        log("CollectAll: form order modal detected after barcode; handling it");
                        await handleFormOrderModal();
                    }
                }
            } else {
                log("CollectAll: barcode not required");
                // Even if barcode is not required, check again for form order modal (in case it appears later)
                await sleep(500);
                formOrderDiv = document.getElementById("requireFormOrderDiv");
                if (formOrderDiv) {
                    formOrderVisible = isFormOrderModalVisible();
                    if (formOrderVisible) {
                        log("CollectAll: form order modal detected (no barcode); handling it");
                        await handleFormOrderModal();
                    }
                }
            }

            log("CollectAll: launching Run Form (IR) for formId=" + String(pickedFormId));
            RUN_FORM_V2_START_TS = Date.now();
            setFormValueMode("ir");
            await runFormAutomationV2();

            var closeSpan = await waitForSelector("span[data-dismiss='modal']", 8000);
            if (closeSpan) {
                log("CollectAll: clicking Close span after form processing");
                closeSpan.click();
            } else {
                log("CollectAll: Close span not found; attempting to ensure modal is closed");
            }

            var closed = await waitForAnyModalToClose(3000);
            if (!closed) {
                log("CollectAll: modal did not close within timeout; proceeding to rescan");
            } else {
                log("CollectAll: modal closed; waiting for table refresh");
            }

            // Add form to list after collection
            addFormToList(formName, formStatus, pickedFormId);

            await waitForFormTableRefresh(1000);
            await sleep(100);
            safety = safety + 1;
        }

        try {
            clearInterval(animTimer);
        } catch (e1) {}

        statusDiv.textContent = "Completed";
        statusDiv.style.color = "#9f9";
        log("CollectAll: completed");

        try {
            clearBarcodeResult();
        } catch (e2) {}
    }

    //==========================
    // CLEAR SUBJECT ELIGIBILITY FEATURE
    //==========================
    // This section contains all functions related to clearing subject eligibility.
    // This feature automates clearing all existing eligibility mapping in the table.
    //==========================

    function ClearEligibilityFunctions() {}

    function startClearMapping() {
        if (CLEAR_MAPPING_CANCELED) {
            log("ClearMapping: startClearMapping cancelled");
            return;
        }
        log("ClearMapping: startClearMapping invoked");

        try {
            localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_CLEAR_MAPPING);
        } catch (e) {
        }

        var path = location.pathname;
        if (path !== "/secure/crfdesign/studylibrary/eligibility/list") {
            log("ClearMapping: not on eligibility list page; redirecting");
            location.href = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
            return;
        }

        executeClearMappingAutomation();
    }

    async function executeClearMappingAutomation() {
        if (CLEAR_MAPPING_CANCELED) {
            log("ClearMapping: executeClearMappingAutomation cancelled");
            return;
        }
        log("ClearMapping: executor started");

        var path = location.pathname;
        if (path !== "/secure/crfdesign/studylibrary/eligibility/list") {
            log("ClearMapping: wrong page; redirecting");
            try {
                localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_CLEAR_MAPPING);
            } catch (e) {
            }
            location.href = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
            return;
        }

        var tbody = await waitForSelector("tbody#eligibilityRefTableBody", 15000);
        if (!tbody) {
            log("ClearMapping: table body missing; stopping");
            clearRunMode();
            return;
        }

        var rows = tbody.querySelectorAll("tr");
        log("ClearMapping: rows found=" + String(rows.length));

        if (rows.length === 0) {
            log("ClearMapping: no rows remain; clearing run mode");
            clearRunMode();
            return;
        }

        await deleteFirstEligibilityRow();
    }



    async function deleteFirstEligibilityRow() {
        if (CLEAR_MAPPING_CANCELED) {
            log("ClearMapping: deleteFirstEligibilityRow cancelled");
            return;
        }
        log("ClearMapping: deleteFirstEligibilityRow started");

        var tbody = await waitForSelector("tbody#eligibilityRefTableBody", 15000);
        if (!tbody) {
            log("ClearMapping: tbody missing in deleteFirstEligibilityRow");
            clearRunMode();
            return;
        }

        var rows = tbody.querySelectorAll("tr");
        if (rows.length === 0) {
            log("ClearMapping: no rows present during deleteFirstEligibilityRow");
            clearRunMode();
            return;
        }

        var tr = rows[0];
        var tds = tr.querySelectorAll("td");
        if (!tds) {
            log("ClearMapping: row has no tds; stopping");
            clearRunMode();
            return;
        }
        if (tds.length < 9) {
            log("ClearMapping: row has insufficient columns; stopping");
            clearRunMode();
            return;
        }

        var actionTd = tds[8];
        var btn = actionTd.querySelector("button.dropdown-toggle");
        if (!btn) {
            log("ClearMapping: dropdown toggle not found");
            clearRunMode();
            return;
        }

        btn.click();
        log("ClearMapping: action dropdown opened");
        await sleep(400);

        var items = document.querySelectorAll("ul.dropdown-menu li a");
        var deleteLink = null;

        var i = 0;
        while (i < items.length) {
            if (CLEAR_MAPPING_CANCELED) {
                log("ClearMapping: deleteFirstEligibilityRow cancelled");
                return;
            }
            var a = items[i];
            var txt = (a.textContent + "").trim().toLowerCase();
            if (txt.indexOf("delete") >= 0) {
                deleteLink = a;
                break;
            }
            i = i + 1;
        }

        if (!deleteLink) {
            log("ClearMapping: delete link not found");
            clearRunMode();
            return;
        }

        if (CLEAR_MAPPING_CANCELED) {
            log("ClearMapping: deleteFirstEligibilityRow cancelled");
            return;
        }
        log("ClearMapping: clicking delete link");
        deleteLink.click();
        await sleep(600);

        var okBtn = null;
        var waited = 0;
        var step = 150;
        var maxWait = 4000;

        while (waited < maxWait) {
            if (CLEAR_MAPPING_CANCELED) {
                log("ClearMapping: deleteFirstEligibilityRow cancelled");
                return;
            }
            okBtn = document.querySelector("button[data-bb-handler='confirm'].btn.btn-primary");
            if (okBtn) {
                break;
            }
            await sleep(step);
            waited = waited + step;
        }

        if (!okBtn) {
            log("ClearMapping: OK button not found in modal");
            clearRunMode();
            return;
        }

        log("ClearMapping: clicking OK button in modal");
        okBtn.click();
        await sleep(1500);

        log("ClearMapping: delete confirmed; forcing page reload");
        localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_CLEAR_MAPPING);
        location.reload();
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

            // Show redirect popup
            var redirectMessage = document.createElement("div");
            redirectMessage.style.textAlign = "center";
            redirectMessage.style.fontSize = "16px";
            redirectMessage.style.color = "#fff";
            redirectMessage.style.padding = "20px";
            redirectMessage.textContent = "Click Import Eligibility Mapping again to run.";

            var redirectPopup = createPopup({
                title: "Import Eligibility Mapping",
                content: redirectMessage,
                width: "400px",
                height: "auto"
            });

            location.href = ELIGIBILITY_LIST_URL;
            return;
        }

        try {
            localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
        } catch (e) {
        }

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
                try {
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {}
                break;
            }
            var opened = await openAddEligibilityModal();
            if (!opened) {
                log("ImportElig: cannot open modal; stopping");
                break;
            }
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                try {
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {}
                break;
            }
            // Wait a bit for modal to fully load
            await sleep(500);
            var eligList = await readEligibilityItemCodesFromSelect();
            var allCodes = eligList.codes;
            var codeToVal = eligList.valueMap;
            log("ImportElig: read " + String(allCodes.length) + " codes from select");
            var pick = "";
            var j = 0;
            var ignoredCount = 0;
            var existingCount = 0;
            var importedCount = 0;
            while (j < allCodes.length) {
                var c = allCodes[j];
                if (shouldIgnoreCode(c)) {
                    ignoredCount = ignoredCount + 1;
                    j = j + 1;
                    continue;
                }
                var inExisting = existingSet.has(c);
                var inImported = importedSet.has(c);
                if (inExisting) {
                    existingCount = existingCount + 1;
                }
                if (inImported) {
                    importedCount = importedCount + 1;
                }
                if (!inExisting && !inImported) {
                    pick = c;
                    log("ImportElig: found new code to process: " + String(c));
                    break;
                }
                j = j + 1;
            }
            log("ImportElig: code filtering - ignored=" + String(ignoredCount) + ", existing=" + String(existingCount) + ", imported=" + String(importedCount) + ", selected=" + String(pick || "none"));
            if (!pick || pick.length === 0) {
                log("ImportElig: no new items; finishing");
                // Close modal before finishing
                var closeBtn = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeBtn) {
                    closeBtn.click();
                    await sleep(500);
                }
                try {
                    localStorage.removeItem(STORAGE_RUN_MODE);
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {
                }
                clearImportedItemsSet();
                // Note: The completion message will be shown by the callback in startImportEligibilityMapping
                return;
            }
            log("ImportElig: attempting to select code: " + String(pick));
            var selOk = await selectEligibilityItemByCode(pick, codeToVal);
            if (!selOk) {
                log("ImportElig: select failed; closing modal");
                addToImportEligFailedList(pick);
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
                addToImportEligFailedList(pick);
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
                try {
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {}
                break;
            }
            log("ImportElig: waiting for comparator to appear after Check Item selection");
            var compReady = await waitForComparatorReady(1000);
            if (!compReady) {
                log("ImportElig: comparator never appeared; skipping");
                addToImportEligFailedList(pick);
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
                addToImportEligFailedList(pick);
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
                addToImportEligFailedList(pick);
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
                addToImportEligFailedList(pick);
                break;
            }
            if (!isRunModeSet(RUNMODE_ELIG_IMPORT)) {
                log("ImportElig: run mode cleared (X pressed); stopping loop");
                try {
                    localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                } catch (e) {}
                break;
            }
            existingSet.add(pick);
            importedSet.add(pick);
            persistImportedItemsSet(importedSet);
            addToImportEligCompletedList(pick);
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

    //==========================
    // RUN FORM FEATURES (OOR A, OOR B, IR)
    //==========================
    // This section contains all functions related to form automation features:
    // - Run Form (OOR) A: Out of Range values below minimum
    // - Run Form (OOR) B: Out of Range values above maximum
    // - Run Form (IR): In Range values within valid range
    // These functions handle form field filling, item group processing, range parsing,
    // and value selection based on the specified mode (oorA, oorB, or ir).
    //==========================
    function FormAutomationFunctions() {}

    async function handleRepeatingItemGroupAddNew(groupDiv) {
        var addBtn = groupDiv.querySelector("a.btn.btn-default.blue.pull-right");
        if (!addBtn) {
            return false;
        }

        log("Run Form V2: repeating item group detected, clicking Add New");

        addBtn.click();
        await sleep(500);

        var okBtn = await waitForSelector("button[data-bb-handler='confirm'].btn.btn-primary", 8000);
        if (okBtn) {
            okBtn.click();
            log("Run Form V2: Add New confirm clicked");
            await sleep(1200);
        } else {
            log("Run Form V2: Add New confirm button not found");
        }

        var backLink = await waitForSelector("span[onclick^='collectModalFormAction']", 8000);
        if (backLink) {
            backLink.click();
            log("Run Form V2: returning to Form Details");
            await sleep(1500);
        } else {
            log("Run Form V2: Form Details return span not found");
        }

        return true;
    }

    function setFormValueMode(s) {
        try {
            localStorage.setItem(STORAGE_FORM_VALUE_MODE, String(s));
        } catch (e) {}
    }

    function getFormValueMode() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_FORM_VALUE_MODE);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    function getItemTextCellFromRow(tr) {
        if (!tr) {
            return null;
        }
        var td = tr.querySelector("td.itemText");
        return td;
    }

    function getRangeTextFromHelp(itemTextTd) {
        var t = "";
        if (itemTextTd) {
            var help = itemTextTd.querySelector("div.itemHelpText");
            if (help) {
                t = getText(help);
            }
        }
        return t;
    }

    function getRangeTextFromItemText(itemTextTd) {
        var t = "";
        if (itemTextTd) {
            t = getText(itemTextTd);
        }
        return t;
    }

    function parseRangeSpecFromText(t) {
        if (typeof t !== "string") {
            return null;
        }
        var s = t;
        s = s.replace(/\u00A0/g, " ");
        s = s.trim();

        var m2 = s.match(/(\d+(?:\.\d+)?)\s*(?:-|–|—|to|TO)\s*(\d+(?:\.\d+)?)/i);
        if (m2 && m2[1] && m2[2]) {
            var a = parseFloat(m2[1]);
            var b = parseFloat(m2[2]);
            if (!isNaN(a) && !isNaN(b)) {
                return { kind: "between", min: a, max: b };
            }
        }

        var mle = s.match(/(?:≤|<=)\s*(\d+(?:\.\d+)?)/);
        if (mle && mle[1]) {
            var vle = parseFloat(mle[1]);
            if (!isNaN(vle)) {
                return { kind: "le", t: vle };
            }
        }

        var mlt = s.match(/<\s*(\d+(?:\.\d+)?)/);
        if (mlt && mlt[1]) {
            var vlt = parseFloat(mlt[1]);
            if (!isNaN(vlt)) {
                return { kind: "lt", t: vlt };
            }
        }

        var mge = s.match(/(?:≥|>=)\s*(\d+(?:\.\d+)?)/);
        if (mge && mge[1]) {
            var vge = parseFloat(mge[1]);
            if (!isNaN(vge)) {
                return { kind: "ge", t: vge };
            }
        }

        var mgt = s.match(/>\s*(\d+(?:\.\d+)?)/);
        if (mgt && mgt[1]) {
            var vgt = parseFloat(mgt[1]);
            if (!isNaN(vgt)) {
                return { kind: "gt", t: vgt };
            }
        }

        return null;
    }

    function findRangeSpecForRow(tr) {
        var td = getItemTextCellFromRow(tr);
        var helpText = getRangeTextFromHelp(td);
        var spec = parseRangeSpecFromText(helpText);
        if (spec) {
            return spec;
        }
        var itemText = getRangeTextFromItemText(td);
        var spec2 = parseRangeSpecFromText(itemText);
        if (spec2) {
            return spec2;
        }
        return null;
    }

    function randomIntInInclusiveRange(a, b) {
        var min = Math.ceil(a);
        var max = Math.floor(b);
        var n = Math.floor(Math.random() * (max - min + 1)) + min;
        return n;
    }

    function pickIntegerForSpec(spec, mode) {
        if (!spec || !mode) {
            return null;
        }
        if (spec.kind === "between") {
            var lo = Math.floor(spec.min);
            var hi = Math.floor(spec.max);
            if (mode === "ir") {
                if (hi < lo) {
                    return null;
                }
                return randomIntInInclusiveRange(lo, hi);
            }
            if (mode === "oorA") {
                var belowA = lo - 1;
                if (belowA >= 0) {
                    return belowA;
                } else {
                    return 0;
                }
            }
            if (mode === "oorB") {
                return hi + 1;
            }
            if (mode === "oor") {
                var below = lo - 1;
                var above = hi + 1;
                var coin = Math.random() < 0.5;
                if (coin) {
                    if (below >= 0) {
                        return below;
                    } else {
                        return above;
                    }
                } else {
                    return above;
                }
            }
            return null;
        }
        if (spec.kind === "lt") {
            var t1 = Math.floor(spec.t);
            if (mode === "ir") {
                var maxVal = t1 - 1;
                if (maxVal >= 0) {
                    return randomIntInInclusiveRange(0, maxVal);
                } else {
                    return 0;
                }
            }
            if (mode === "oorA") {
                return t1 + 1;
            }
            if (mode === "oorB") {
                return t1 + 1;
            }
            if (mode === "oor") {
                return t1;
            }
            return null;
        }
        if (spec.kind === "le") {
            var t2 = Math.floor(spec.t);
            if (mode === "ir") {
                if (t2 >= 0) {
                    return randomIntInInclusiveRange(0, t2);
                } else {
                    return 0;
                }
            }
            if (mode === "oorA") {
                return t2 + 1;
            }
            if (mode === "oorB") {
                return t2 + 1;
            }
            if (mode === "oor") {
                return t2 + 1;
            }
            return null;
        }
        if (spec.kind === "gt") {
            var t3 = Math.floor(spec.t);
            if (mode === "ir") {
                var minVal = t3 + 1;
                var maxVal = t3 + 100;
                return randomIntInInclusiveRange(minVal, maxVal);
            }
            if (mode === "oorA") {
                var v3a = t3 - 1;
                if (v3a >= 0) {
                    return v3a;
                } else {
                    return 0;
                }
            }
            if (mode === "oorB") {
                var v3b = t3 - 1;
                if (v3b >= 0) {
                    return v3b;
                } else {
                    return 0;
                }
            }
            if (mode === "oor") {
                return t3;
            }
            return null;
        }
        if (spec.kind === "ge") {
            var t4 = Math.floor(spec.t);
            if (mode === "ir") {
                var minVal = t4;
                var maxVal = t4 + 100;
                return randomIntInInclusiveRange(minVal, maxVal);
            }
            if (mode === "oorA") {
                var v4a = t4 - 1;
                if (v4a >= 0) {
                    return v4a;
                } else {
                    return 0;
                }
            }
            if (mode === "oorB") {
                var v4b = t4 - 1;
                if (v4b >= 0) {
                    return v4b;
                } else {
                    return 0;
                }
            }
            if (mode === "oor") {
                var v2 = t4 - 1;
                if (v2 >= 0) {
                    return v2;
                } else {
                    return 0;
                }
            }
            return null;
        }
        return null;
    }

    function pickDecimalForSpec(spec, mode, places) {
        if (!spec || !mode) {
            return null;
        }
        var p = 0;
        if (typeof places === "number") {
            p = places;
        }
        var iv = pickIntegerForSpec(spec, mode);
        if (iv === null || iv === undefined) {
            if (spec.kind === "between") {
                var lo = spec.min;
                var hi = spec.max;
                if (mode === "ir") {
                    var base = Math.random() * (hi - lo) + lo;
                    var out = p <= 0 ? Math.round(base) : parseFloat(base.toFixed(p));
                    return out;
                }
                if (mode === "oorA") {
                    var belowA = spec.min - 1;
                    var outA = p <= 0 ? Math.round(belowA) : parseFloat(belowA.toFixed(p));
                    return outA;
                }
                if (mode === "oorB") {
                    var aboveB = spec.max + 1;
                    var outB = p <= 0 ? Math.round(aboveB) : parseFloat(aboveB.toFixed(p));
                    return outB;
                }
                if (mode === "oor") {
                    var below = spec.min - 1;
                    var above = spec.max + 1;
                    var coin = Math.random() < 0.5;
                    var pick = coin ? below : above;
                    var out2 = p <= 0 ? Math.round(pick) : parseFloat(pick.toFixed(p));
                    return out2;
                }
                return null;
            }
            return null;
        }
        var out3 = p <= 0 ? Math.round(iv) : parseFloat(iv.toFixed(p));
        return out3;
    }

    function getText(el) {
        if (!el) {
            return "";
        }
        var s = (el.textContent + "").replace(/\u00A0/g, " ");
        s = s.trim();
        return s;
    }

    // Check whether a visible item group (itemGroupData_*) contains at least one
    function groupHasActionableItems(groupDiv) {
        if (!groupDiv) {
            return false;
        }
        var gid = parseItemGroupId(groupDiv);
        if (!gid || gid.length === 0) {
            return false;
        }
        var table = groupDiv.querySelector("table#collectTable_" + String(gid));
        if (!table) {
            return false;
        }
        var rows = table.querySelectorAll("tr[id^=\"itemDataCollectRow_\"]");
        var i = 0;
        while (i < rows.length) {
            var tr = rows[i];
            var styleAttr = tr.getAttribute("style") || "";
            var hiddenInline = styleAttr.indexOf("display: none") !== -1;
            var visibleInline = styleAttr.indexOf("display: table-row") !== -1 || styleAttr.trim().length === 0;
            if (visibleInline && !hiddenInline) {
                var td = tr.querySelector("td.itemControl");
                var actionable = hasActionableControl(td);
                if (actionable) {
                    return true;
                }
            }
            i = i + 1;
        }
        return false;
    }
    // Find the next visible itemGroupData container that has actionable items and has not yet been processed in this V2 run.
    function findNextVisibleUnprocessedGroupV2(formEl, processedGroups) {
        if (!formEl) {
            return null;
        }
        var groups = formEl.querySelectorAll("div[id^=\"itemGroupData_\"]");
        var i = 0;
        while (i < groups.length) {
            var g = groups[i];
            var gid = parseItemGroupId(g);
            var styleAttr = g.getAttribute("style") || "";
            var visibleInherit = styleAttr.indexOf("display: inherit") !== -1;
            var visibleEmpty = styleAttr.trim().length === 0;
            var visibleGroup = visibleInherit || visibleEmpty;
            var already = processedGroups[gid] === true;
            if (visibleGroup && !already) {
                var hasAny = groupHasActionableItems(g) || !!g.querySelector("button[onclick^=\"return unlockDeviceItems(\"]");
                if (hasAny) {
                    return g;
                }
            }
            i = i + 1;
        }
        return null;
    }

    // Fill actionable item rows one-by-one inside a collectTable container
    async function fillItemControlsOneByOneV2(containerEl) {
        if (!containerEl) {
            return;
        }
        var filledIds = {};
        var safety = 0;
        var safetyMax = 500;
        while (safety < safetyMax) {
            if (isPaused()) {
                log("Run Form V2: paused during one-by-one fill");
                return;
            }
            var rows = containerEl.querySelectorAll("tr[id^=\"itemDataCollectRow_\"]");
            var progressed = false;
            var i = 0;
            while (i < rows.length) {
                var tr = rows[i];
                var idVal = tr.id + "";
                var isDevice = idVal.indexOf("deviceConnectRow_") !== -1;
                var hasInvoke = tr.classList.contains("invokeDeviceTr");
                if (isDevice) {
                    i = i + 1;
                    continue;
                }
                if (hasInvoke) {
                    i = i + 1;
                    continue;
                }
                var styleAttr = tr.getAttribute("style") || "";
                var hiddenInline = styleAttr.indexOf("display: none") !== -1;
                var visibleInline = styleAttr.indexOf("display: table-row") !== -1 || styleAttr.trim().length === 0;
                if (!visibleInline || hiddenInline) {
                    i = i + 1;
                    continue;
                }
                var td = tr.querySelector("td.itemControl");
                var rowId = getItemRowId(tr);
                var alreadyFilled = filledIds[rowId] === true;
                var actionable = !alreadyFilled && hasActionableControl(td);
                if (actionable) {
                    log("Run Form V2: filling single rowId=" + String(rowId));
                    await fillSingleItemControl(td);
                    if (rowId && rowId.length > 0) {
                        filledIds[rowId] = true;
                    }
                    progressed = true;
                    await sleep(DELAY_V2_ITEM_MS);
                    break;
                }
                i = i + 1;
            }
            if (!progressed) {
                log("Run Form V2: no actionable items remain in container");
                break;
            }
            safety = safety + 1;
        }
    }

    // Process a single item group directly (no Collect click)

    async function processItemGroupDirectV2(groupDiv) {
        var gid = parseItemGroupId(groupDiv);
        if (!gid || gid.length === 0) {
            log("Run Form V2: direct path missing gid");
            return;
        }

        var didAdd = await handleRepeatingItemGroupAddNew(groupDiv);
        if (didAdd) {
            log("Run Form V2: Add New flow completed for gid=" + String(gid));
            return;
        }

        var unlockBtn = groupDiv.querySelector("button[onclick^=\"return unlockDeviceItems(\"]");
        if (unlockBtn) {
            unlockBtn.click();
            await sleep(FORM_DELAY_MS);
            var reasonBox = await waitForSelector("textarea#unlockDeviceItemsReason", 12000);
            if (reasonBox) {
                reasonBox.value = "Test";
                var evt = new Event("input", { bubbles: true });
                reasonBox.dispatchEvent(evt);
                log("Run Form V2: unlock reason set gid=" + String(gid));
            }
            var okBtn = await waitForSelector("button.btn.btn.green[data-bb-handler=\"success\"]", 12000);
            if (okBtn) {
                okBtn.click();
                await sleep(FORM_DELAY_MS);
            }
        }

        var table = groupDiv.querySelector("table#collectTable_" + String(gid));
        var waited = 0;
        var maxWait = 12000;

        while (!table && waited < maxWait) {
            await sleep(300);
            waited = waited + 300;
            table = groupDiv.querySelector("table#collectTable_" + String(gid));
        }

        if (!table) {
            log("Run Form V2: collectTable not found gid=" + String(gid));
            return;
        }

        log("Run Form V2: filling table gid=" + String(gid));
        await fillItemControlsOneByOneV2(table);
    }


    // Run Form V2 orchestration:
    async function runFormAutomationV2() {
        if (isPaused()) {
            log("Run Form V2: paused; skipping start");
            return;
        }
        var processedGroups = {};
        var safety = 0;
        var safetyMax = 500;
        while (safety < safetyMax) {
            var modal = document.querySelector("#ajaxModal, .modal");
            var formEl = null;
            if (modal) {
                formEl = modal.querySelector("form#modalInput");
            }
            if (!formEl) {
                formEl = document.querySelector("form#modalInput");
            }
            if (!formEl) {
                log("Run Form V2: form#modalInput not found");
                break;
            }
            var nextGroup = findNextVisibleUnprocessedGroupV2(formEl, processedGroups);
            if (!nextGroup) {
                log("Run Form V2: no next visible group with actionable items");
                break;
            }
            var gid = parseItemGroupId(nextGroup);
            log("Run Form V2: processing itemGroup gid=" + String(gid));
            await processItemGroupDirectV2(nextGroup);
            processedGroups[gid] = true;
            await sleep(DELAY_V2_GROUP_RESCAN_MS);
            safety = safety + 1;
        }
        var finalBtn = document.querySelector("button.btn.green[onclick^=\"return saveAndPostModalForm(false, '/secure/datacollection/formdata/validateform/\"]");
        if (finalBtn) {
            finalBtn.click();
            log("Run Form V2: clicked final Save and Return");

            await sleep(1000);

            var deviationTextarea = document.querySelector("textarea#userDeviationReason");
            if (deviationTextarea) {
                log("Run Form V2: deviation modal detected (found userDeviationReason textarea)");
                deviationTextarea.value = "Test";
                var evt = new Event("input", { bubbles: true });
                deviationTextarea.dispatchEvent(evt);
                log("Run Form V2: deviation reason set to 'Test'");

                await sleep(200);

                var saveBtn = document.querySelector("button.btn.btn.green[data-bb-handler=\"success\"]");
                if (saveBtn) {
                    saveBtn.click();
                    log("Run Form V2: clicked deviation modal Save button");
                } else {
                    log("Run Form V2: deviation modal Save button not found");
                }
            }
        } else {
            log("Run Form V2: final Save and Return button not found");
        }
        if (RUN_FORM_V2_START_TS > 0) {
            var secs = (Date.now() - RUN_FORM_V2_START_TS) / 1000;
            var s = secs.toFixed(2);
            log("Run Form V2: elapsed " + String(s) + " s");
            RUN_FORM_V2_START_TS = 0;
        }
    }
    // Extract the numeric item group id from element id.
    function parseItemGroupId(groupDiv) {
        var idAttr = groupDiv.id + "";
        var m = idAttr.match(/itemGroupData_(\d+)/);
        if (m && m[1]) {
            return m[1];
        }
        return "";
    }

    // Return a random integer between min and max inclusive.
    function randomIntInclusive(min, max) {
        var a = Math.ceil(min);
        var b = Math.floor(max);
        return Math.floor(Math.random() * (b - a + 1)) + a;
    }

    // Fill or interact with a single item control (text, number, checkbox, radio, time).

    async function fillSingleItemControl(controlTd) {
        if (!controlTd) {
            return;
        }
        var noval = controlTd.querySelector("span.novalue");
        if (noval) {
            log("Run Form: itemControl novalue; skipping");
            return;
        }
        var timeBtn = controlTd.querySelector("div.timeSetButtons button[id$=\"_currentTime\"].btn.btn-xs.dark");
        if (timeBtn) {
            log("Run Form: clicking Current Time");
            timeBtn.click();
            return;
        }
        var radios = controlTd.querySelectorAll("div.radio-list input[type=\"radio\"]");
        if (radios && radios.length > 0) {
            var tr = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
            var rowId = getItemRowId(tr);

            var repeatRadio = null;
            var i = 0;
            while (i < radios.length) {
                var radio = radios[i];
                var label = radio.closest("label");
                if (label) {
                    var labelText = (label.textContent + "").toLowerCase();
                    if (labelText.indexOf("repeat") !== -1) {
                        repeatRadio = radio;
                        log("Run Form: found radio with 'repeat' in label rowId=" + String(rowId));
                        break;
                    }
                }
                i = i + 1;
            }

            var ok = await radioSelectBestEffort(controlTd, rowId, repeatRadio);
            log("Run Form: radio-select result rowId=" + String(rowId) + " ok=" + String(ok));
            return;
        }
        var ta = controlTd.querySelector("textarea.collectInput.text");
        if (ta) {
            if (ta.hasAttribute("readonly")) {
                ta.removeAttribute("readonly");
            }
            if (ta.disabled === true) {
                ta.disabled = false;
            }
            ta.value = "Test";
            var evt = new Event("input", { bubbles: true });
            ta.dispatchEvent(evt);
            log("Run Form: textarea filled");
            return;
        }
        var textBox = controlTd.querySelector("input.collectInput.text");
        if (textBox) {
            if (textBox.hasAttribute("readonly")) {
                textBox.removeAttribute("readonly");
            }
            if (textBox.disabled === true) {
                textBox.disabled = false;
            }
            textBox.value = "Test";
            var evtX = new Event("input", { bubbles: true });
            textBox.dispatchEvent(evtX);
            log("Run Form: text input set value=Test");
            return;
        }
        var intBox = controlTd.querySelector("input.collectInput.integer");
        if (intBox) {
            if (intBox.hasAttribute("readonly")) {
                intBox.removeAttribute("readonly");
            }
            if (intBox.disabled === true) {
                intBox.disabled = false;
            }
            var modeInt = getFormValueMode();
            var n = null;
            if (modeInt === "oor" || modeInt === "ir" || modeInt === "oorA" || modeInt === "oorB") {
                var trI = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
                var specI = findRangeSpecForRow(trI);
                if (specI) {
                    var pickI = pickIntegerForSpec(specI, modeInt);
                    if (pickI !== null && pickI !== undefined) {
                        n = pickI;
                    }
                }
            }
            if (n === null || n === undefined) {
                n = randomIntInclusive(1, 100);
            }
            intBox.value = String(n);
            var evt2 = new Event("input", { bubbles: true });
            intBox.dispatchEvent(evt2);
            log("Run Form: integer input set value=" + String(n));
            return;
        }
        var decBox = controlTd.querySelector("input.collectInput.decimal");
        if (decBox) {
            if (decBox.hasAttribute("readonly")) {
                decBox.removeAttribute("readonly");
            }
            if (decBox.disabled === true) {
                decBox.disabled = false;
            }
            var p = getDecimalPlacesFromMeta(controlTd);
            var modeDec = getFormValueMode();
            var d = null;
            if (modeDec === "oor" || modeDec === "ir" || modeDec === "oorA" || modeDec === "oorB") {
                var trD = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
                var specD = findRangeSpecForRow(trD);
                if (specD) {
                    var pickD = pickDecimalForSpec(specD, modeDec, p);
                    if (pickD !== null && pickD !== undefined) {
                        d = pickD;
                    }
                }
            }
            if (d === null || d === undefined) {
                var base = Math.random() * (100 - 1) + 1;
                if (p <= 0) {
                    d = String(Math.round(base));
                } else {
                    d = String(parseFloat(base.toFixed(p)));
                }
            } else {
                if (p <= 0) {
                    d = String(Math.round(d));
                } else {
                    d = String(parseFloat(Number(d).toFixed(p)));
                }
            }
            decBox.value = d;
            var evt3 = new Event("input", { bubbles: true });
            decBox.dispatchEvent(evt3);
            log("Run Form: decimal input set value=" + String(d));
            return;
        }
        var chk = controlTd.querySelector("input[type=\"checkbox\"]");
        if (chk) {
            var tr2 = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
            var rowId2 = getItemRowId(tr2);
            var ok2 = await checkboxSelectBestEffort(controlTd, rowId2);
            log("Run Form: checkbox-select result rowId=" + String(rowId2) + " ok=" + String(ok2));
            return;
        }

        var s2container = null;

        s2container = controlTd.querySelector("div.select2-container.collectInput");

        if (!s2container) {
            var spanWrap = controlTd.querySelector("span");
            if (spanWrap) {
                s2container = spanWrap.querySelector("div.select2-container.collectInput");
            }
        }

        if (!s2container) {
            s2container = controlTd.querySelector("div.select2-container");
        }

        if (s2container) {

            var choice = s2container.querySelector("a.select2-choice");
            choice.click();
            log("Run Form: select2 open (arrow)");

            await sleep(200);

            var drop = document.querySelector(".select2-drop-active");

            if (!drop) {
                log("Run Form: no .select2-drop-active; retrying");
                choice.click();
                await sleep(300);
                drop = document.querySelector(".select2-drop-active");
            }

            if (!drop) {
                log("Run Form: still no active drop — ABORT");
                return;
            }

            var results = drop.querySelectorAll("li.select2-result-selectable div.select2-result-label");

            if (results.length === 0) {
                log("Run Form: no results in .select2-drop-active");
                return;
            }

            var valid = [];
            var i = 0;
            while (i < results.length) {
                var t = (results[i].textContent + "").trim();
                if (t.length > 0 && t !== "-") {
                    valid.push(results[i]);
                }
                i++;
            }

            if (valid.length === 0) {
                log("Run Form: only '-' placeholder found");
                return;
            }

            var pick = valid[Math.floor(Math.random() * valid.length)];
            pick.closest("li").click();

            log("Run Form: Select2 picked: " + pick.textContent.trim());
            await sleep(150);

            return;
        }




        log("Run Form: itemControl has no known actionable inputs; skipping");
    }

    // Extract the numeric item row id from a row element's id.
    function getItemRowId(tr) {
        if (!tr) {
            return "";
        }
        var idAttr = tr.id + "";
        var m = idAttr.match(/itemDataCollectRow_(\d+)/);
        if (m && m[1]) {
            return m[1];
        }
        return "";
    }

    // Parse decimal-places information from field metadata.
    function getDecimalPlacesFromMeta(controlTd) {
        var decs = 1;
        var tr = null;
        if (controlTd) {
            tr = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
        }
        if (!tr) {
            return decs;
        }
        var metaTd = tr.querySelector("td.itemMeta");
        if (!metaTd) {
            return decs;
        }
        var items = metaTd.querySelectorAll("ul.metaList li");
        var i = 0;
        var fmtText = "";
        while (i < items.length) {
            var t = (items[i].textContent + "").trim();
            var hasFmt = t.toLowerCase().indexOf("format:") !== -1;
            if (hasFmt) {
                fmtText = t;
                break;
            }
            i = i + 1;
        }
        if (!fmtText || fmtText.length === 0) {
            return decs;
        }
        var parts = fmtText.split(":");
        var fmt = "";
        if (parts.length >= 2) {
            fmt = (parts[1] + "").trim();
        }
        if (!fmt || fmt.length === 0) {
            return decs;
        }
        var m = fmt.match(/\.(.+)$/);
        if (!m || !m[1]) {
            return 0;
        }
        var tail = m[1];
        var count = 0;
        var j = 0;
        while (j < tail.length) {
            var ch = tail[j];
            var ok = ch === "#" || ch === "0";
            if (ok) {
                count = count + 1;
            }
            j = j + 1;
        }
        return count;
    }

    // Return true if the control cell contains an empty/actionable input.

    function hasActionableControl(controlTd) {
        if (!controlTd) {
            return false;
        }

        var noval = controlTd.querySelector("span.novalue");
        if (noval) {
            return false;
        }

        var timeBtn = controlTd.querySelector("div.timeSetButtons button[id$=\"_currentTime\"].btn.btn-xs.dark");
        if (timeBtn) {
            return true;
        }

        var radios = controlTd.querySelectorAll("div.radio-list input[type=\"radio\"]");
        if (radios && radios.length > 0) {
            var anyChecked = controlTd.querySelector("div.radio-list input[type=\"radio\"]:checked");
            if (!anyChecked) {
                return true;
            }
            return true;
        }

        var ta = controlTd.querySelector("textarea.collectInput.text");
        if (ta) {
            return true;
        }

        var tb = controlTd.querySelector("input.collectInput.text");
        if (tb) {
            return true;
        }

        var intBox = controlTd.querySelector("input.collectInput.integer");
        if (intBox) {
            return true;
        }

        var decBox = controlTd.querySelector("input.collectInput.decimal");
        if (decBox) {
            return true;
        }

        var chk = controlTd.querySelector("input[type=\"checkbox\"]");
        if (chk) {
            return true;
        }

        return false;
    }

    // Try multiple strategies to select a radio option reliably.
    async function radioSelectBestEffort(controlTd, rowId, preferredRadio) {
        var radios = controlTd.querySelectorAll("div.radio-list input[type=\"radio\"]");
        if (!radios || radios.length === 0) {
            return false;
        }
        var r = null;
        if (preferredRadio) {
            r = preferredRadio;
            log("Run Form: using preferred radio (repeat) rowId=" + String(rowId));
        } else {
            var randomIndex = Math.floor(Math.random() * radios.length);
            r = radios[0];
            log("Run Form: selecting radio index=" + String(randomIndex) + " rowId=" + String(rowId));
        }
        if (r.disabled === true) {
            r.disabled = false;
            log("Run Form: radio enabled rowId=" + String(rowId));
        }
        if (r.hasAttribute("readonly")) {
            r.removeAttribute("readonly");
            log("Run Form: radio readonly removed rowId=" + String(rowId));
        }
        var before = r.checked;
        log("Run Form: radio attempt rowId=" + String(rowId) + " checkedBefore=" + String(before));
        r.focus();
        r.click();
        var evt1 = new Event("change", { bubbles: true });
        r.dispatchEvent(evt1);
        var evt2 = new Event("input", { bubbles: true });
        r.dispatchEvent(evt2);
        await sleep(80);
        var after1 = r.checked || !!controlTd.querySelector("div.radio-list input[type=\"radio\"]:checked");
        log("Run Form: radio after click rowId=" + String(rowId) + " checked=" + String(after1));
        if (after1) {
            return true;
        }
        r.checked = true;
        var evt3 = new Event("change", { bubbles: true });
        r.dispatchEvent(evt3);
        await sleep(60);
        var after2 = r.checked || !!controlTd.querySelector("div.radio-list input[type=\"radio\"]:checked");
        log("Run Form: radio after force-check rowId=" + String(rowId) + " checked=" + String(after2));
        if (after2) {
            return true;
        }
        var label = r.closest("label");
        if (label) {
            label.click();
            await sleep(80);
        }
        var after3 = !!controlTd.querySelector("div.radio-list input[type=\"radio\"]:checked");
        log("Run Form: radio after label click rowId=" + String(rowId) + " checked=" + String(after3));
        if (after3) {
            return true;
        }
        var spanWrap = r.closest("span");
        if (spanWrap) {
            spanWrap.click();
            await sleep(60);
        }
        var after4 = !!controlTd.querySelector("div.radio-list input[type=\"radio\"]:checked");
        log("Run Form: radio final state rowId=" + String(rowId) + " checked=" + String(after4));
        return after4;
    }

    // Fill actionable item rows one by one until no progress occurs.
    async function fillItemControlsOneByOne(containerEl) {
        if (!containerEl) {
            return;
        }
        var filledIds = {};
        var safety = 0;
        var safetyMax = 500;
        while (safety < safetyMax) {
            if (isPaused()) {
                log("Run Form: paused during one-by-one fill");
                return;
            }
            var rows = containerEl.querySelectorAll("tr[id^=\"itemDataCollectRow_\"]");
            var progressed = false;
            var i = 0;
            while (i < rows.length) {
                var tr = rows[i];
                var idVal = tr.id + "";
                var isDevice = idVal.indexOf("deviceConnectRow_") !== -1;
                var hasInvoke = tr.classList.contains("invokeDeviceTr");
                if (isDevice) {
                    i = i + 1;
                    continue;
                }
                if (hasInvoke) {
                    i = i + 1;
                    continue;
                }
                var styleAttr = tr.getAttribute("style") || "";
                var hiddenInline = styleAttr.indexOf("display: none") !== -1;
                var visibleInline = styleAttr.indexOf("display: table-row") !== -1 || styleAttr.trim().length === 0;
                if (!visibleInline || hiddenInline) {
                    i = i + 1;
                    continue;
                }
                var td = tr.querySelector("td.itemControl");
                var rowId = getItemRowId(tr);

                var actionable = hasActionableControl(td);
                if (actionable) {
                    log("Run Form V2: FORCING overwrite rowId=" + String(rowId));
                    await fillSingleItemControl(td);

                    progressed = true;
                    await sleep(DELAY_V2_ITEM_MS);
                    break;
                }
                i = i + 1;
            }
            if (!progressed) {
                log("Run Form: no actionable items remain in container");
                break;
            }
            safety = safety + 1;
        }
    }

    // Ensure an audit reason is present when checking or changing data.
    function ensureAuditDefault(controlTd) {
        if (!controlTd) {
            return;
        }
        var hiddenAudit = controlTd.querySelector("input.auditReasonForChange.collectInput[id$=\"_auditRecord\"]");
        if (!hiddenAudit) {
            var tr = controlTd.closest("tr[id^=\"itemDataCollectRow_\"]");
            if (tr) {
                hiddenAudit = tr.querySelector("input.auditReasonForChange.collectInput[id$=\"_auditRecord\"]");
            }
        }
        if (!hiddenAudit) {
            return;
        }
        var val = (hiddenAudit.value + "").trim();
        if (val.length === 0) {
            hiddenAudit.value = "Data clarification";
            var evt1 = new Event("input", { bubbles: true });
            hiddenAudit.dispatchEvent(evt1);
            var evt2 = new Event("change", { bubbles: true });
            hiddenAudit.dispatchEvent(evt2);
            log("Run Form: audit reason defaulted to Data clarification");
        }
    }

    // Attempt to check a checkbox control with multiple strategies.
    async function checkboxSelectBestEffort(controlTd, rowId) {
        var cb = controlTd.querySelector("input[type=\"checkbox\"]");
        if (!cb) {
            return false;
        }
        if (cb.disabled === true) {
            cb.disabled = false;
            log("Run Form: checkbox enabled rowId=" + String(rowId));
        }
        if (cb.hasAttribute("readonly")) {
            cb.removeAttribute("readonly");
            log("Run Form: checkbox readonly removed rowId=" + String(rowId));
        }
        var before = cb.checked;
        log("Run Form: checkbox attempt rowId=" + String(rowId) + " checkedBefore=" + String(before));
        cb.focus();
        cb.click();
        var evt1 = new Event("change", { bubbles: true });
        cb.dispatchEvent(evt1);
        var evt2 = new Event("input", { bubbles: true });
        cb.dispatchEvent(evt2);
        await sleep(80);
        var after1 = cb.checked;
        log("Run Form: checkbox after click rowId=" + String(rowId) + " checked=" + String(after1));
        if (after1) {
            ensureAuditDefault(controlTd);
            return true;
        }
        cb.checked = true;
        var evt3 = new Event("change", { bubbles: true });
        cb.dispatchEvent(evt3);
        await sleep(60);
        var after2 = cb.checked;
        log("Run Form: checkbox after force-check rowId=" + String(rowId) + " checked=" + String(after2));
        if (after2) {
            ensureAuditDefault(controlTd);
            return true;
        }
        var label = cb.closest("label");
        if (label) {
            label.click();
            await sleep(80);
        }
        var spanWrap = cb.closest("span");
        if (spanWrap) {
            spanWrap.click();
            await sleep(60);
        }
        var after3 = cb.checked;
        log("Run Form: checkbox final state rowId=" + String(rowId) + " checked=" + String(after3));
        if (after3) {
            ensureAuditDefault(controlTd);
        }
        return after3;
    }

    //==========================
    // RUN SAMPLE PATHS FEATURE
    //==========================
    // This section contains all functions and constants related to the Run Sample Paths feature.
    // This feature automates the process of locking sample paths by navigating through
    // the sample paths list, detail pages, and update pages to set the locked status.
    //==========================
    function RunSamplePathsFunctions() {}

    async function fetchAndLockSamplePath(samplePathUrl, samplePathName) {
        try {
            log("Locking Sample Path: " + samplePathName);

            var detailHtml = await fetchPage(samplePathUrl);
            var detailDoc = parseHtml(detailHtml);

            var successAlert = detailDoc.querySelector('div.alert.alert-success.alert-dismissable');
            if (successAlert) {
                var alertText = (successAlert.textContent + "").trim();
                if (alertText.indexOf("The sample path has been updated") !== -1) {
                    log("Sample Path already locked: " + samplePathName);
                    return { success: true, message: "Already locked" };
                }
            }

            var editLink = detailDoc.querySelector('a[href*="/secure/samples/configure/paths/update/"]');
            if (!editLink) {
                log("Edit link not found for: " + samplePathName);
                return { success: false, message: "Edit link not found" };
            }

            var updatePath = editLink.getAttribute("href");
            var updateUrl = location.origin + updatePath;

            var updateHtml = await fetchPage(updateUrl);
            var updateDoc = parseHtml(updateHtml);

            var lockCheckbox = updateDoc.querySelector('input#locked');
            if (!lockCheckbox) {
                log("Lock checkbox not found for: " + samplePathName);
                return { success: false, message: "Lock checkbox not found" };
            }

            var isAlreadyLocked = lockCheckbox.hasAttribute("checked") || lockCheckbox.checked;
            if (isAlreadyLocked) {
                log("Sample Path already locked (checkbox): " + samplePathName);
                return { success: true, message: "Already locked" };
            }

            var formElement = updateDoc.querySelector('form');
            if (!formElement) {
                log("Form not found for: " + samplePathName);
                return { success: false, message: "Form not found" };
            }

            var formData = "";
            var inputs = formElement.querySelectorAll('input, textarea, select');
            var i = 0;
            while (i < inputs.length) {
                var input = inputs[i];
                var name = input.getAttribute("name");
                var type = input.getAttribute("type");
                var value = "";

                if (name) {
                    if (name === "locked") {
                        value = "on";
                    } else if (name === "reasonForChange") {
                        value = "Automated lock via script";
                    } else if (type === "checkbox" || type === "radio") {
                        if (input.checked || input.hasAttribute("checked")) {
                            value = input.value || "on";
                        } else {
                            i = i + 1;
                            continue;
                        }
                    } else if (input.tagName.toLowerCase() === "select") {
                        var selectedOption = input.querySelector("option[selected]");
                        if (selectedOption) {
                            value = selectedOption.value || "";
                        } else {
                            value = input.value || "";
                        }
                    } else {
                        value = input.value || "";
                    }

                    if (formData.length > 0) {
                        formData += "&";
                    }
                    formData += encodeURIComponent(name) + "=" + encodeURIComponent(value);
                }
                i = i + 1;
            }

            log("Submitting lock for: " + samplePathName);
            var submitUrl = formElement.getAttribute("action");
            if (!submitUrl || submitUrl === "" || submitUrl === "#") {
                submitUrl = updateUrl;
            } else if (submitUrl.indexOf("http") !== 0) {
                submitUrl = location.origin + submitUrl;
            }

            var resultHtml = await submitForm(submitUrl, formData);
            var resultDoc = parseHtml(resultHtml);

            var resultAlert = resultDoc.querySelector('div.alert.alert-success.alert-dismissable');
            if (resultAlert) {
                var resultText = (resultAlert.textContent + "").trim();
                if (resultText.indexOf("The sample path has been updated") !== -1) {
                    log("Successfully locked: " + samplePathName);
                    return { success: true, message: "Locked successfully" };
                }
            }

            log("Lock submitted but success not confirmed for: " + samplePathName);
            return { success: true, message: "Submitted (unconfirmed)" };

        } catch (error) {
            log("Error locking " + samplePathName + ": " + String(error));
            return { success: false, message: String(error) };
        }
    }

    async function processLockSamplePathsPage() {
        log("processLockSamplePathsPage: start");

        var flag = null;

        try {
            flag = localStorage.getItem(STORAGE_RUN_LOCK_SAMPLE_PATHS);
        } catch (e) { }

        if (!flag) {
            log("processLockSamplePathsPage: flag not set, exiting");
            return;
        }

        log("processLockSamplePathsPage: flag detected");

        var tbody = await waitForSelector("tbody#deviceTbody", 10000);
        if (!tbody) {
            log("processLockSamplePathsPage: tbody#deviceTbody not found");
            return;
        }

        var rows = tbody.querySelectorAll("tr");
        log("processLockSamplePathsPage: rows found=" + String(rows.length));

        var unlockedPaths = [];
        var i = 0;
        while (i < rows.length) {
            var row = rows[i];

            var tds = row.querySelectorAll("td");
            var tdCount = tds.length;

            if (tdCount >= 4) {
                var lockCell = tds[3];
                var lockText = (lockCell.textContent + "").trim();

                if (lockText.toLowerCase() === "no") {
                    var link = tds[0].querySelector("a");
                    var nameCell = tds[0];
                    var pathName = (nameCell.textContent + "").trim();

                    if (link) {
                        var href = link.getAttribute("href") + "";
                        if (href.length > 0) {
                            var fullUrl = location.origin + href;
                            unlockedPaths.push({ url: fullUrl, name: pathName });
                        }
                    }
                }
            }

            i = i + 1;
        }

        log("processLockSamplePathsPage: found " + String(unlockedPaths.length) + " unlocked paths");

        if (unlockedPaths.length === 0) {
            log("processLockSamplePathsPage: no unlocked paths to process");
            try {
                localStorage.removeItem(STORAGE_RUN_LOCK_SAMPLE_PATHS);
            } catch (e) { }

            var mode = getRunMode();
            if (mode === "all") {
                await sleep(1000);
                log("processLockSamplePathsPage: continuing to Study Metadata for eligibility lock");
                updateRunAllPopupStatus("Locking Eligibility");
                try {
                    localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
                } catch (e) {}
                location.href = STUDY_METADATA_URL + "?autoeliglock=1";
            }
            return;
        }

        var popupContainer = document.createElement("div");
        popupContainer.style.display = "flex";
        popupContainer.style.flexDirection = "column";
        popupContainer.style.gap = "16px";
        popupContainer.style.padding = "8px";

        var statusDiv = document.createElement("div");
        statusDiv.id = "lockSamplePathsStatus";
        statusDiv.style.textAlign = "center";
        statusDiv.style.fontSize = "18px";
        statusDiv.style.color = "#fff";
        statusDiv.style.fontWeight = "500";
        statusDiv.textContent = "Locking Sample Paths";

        var progressDiv = document.createElement("div");
        progressDiv.id = "lockSamplePathsProgress";
        progressDiv.style.textAlign = "center";
        progressDiv.style.fontSize = "16px";
        progressDiv.style.color = "#9df";
        progressDiv.textContent = "Processing 0/" + String(unlockedPaths.length);

        var loadingAnimation = document.createElement("div");
        loadingAnimation.id = "lockSamplePathsLoading";
        loadingAnimation.style.textAlign = "center";
        loadingAnimation.style.fontSize = "14px";
        loadingAnimation.style.color = "#9df";
        loadingAnimation.textContent = "Running.";

        var countsDiv = document.createElement("div");
        countsDiv.id = "lockSamplePathsCounts";
        countsDiv.style.textAlign = "center";
        countsDiv.style.fontSize = "14px";
        countsDiv.style.color = "#ccc";
        countsDiv.innerHTML = "<span style='color:#9f9'>Success: 0</span> | <span style='color:#f99'>Failed: 0</span>";

        popupContainer.appendChild(statusDiv);
        popupContainer.appendChild(progressDiv);
        popupContainer.appendChild(loadingAnimation);
        popupContainer.appendChild(countsDiv);

        LOCK_SAMPLE_PATHS_POPUP_REF = createPopup({
            title: "Lock Sample Paths",
            content: popupContainer,
            width: "400px",
            height: "auto",
            onClose: function() {
                log("Lock Sample Paths: cancelled by user");
                try {
                    localStorage.removeItem(STORAGE_RUN_LOCK_SAMPLE_PATHS);
                    localStorage.removeItem(STORAGE_LOCK_SAMPLE_PATHS_POPUP);
                } catch (e) {}
                LOCK_SAMPLE_PATHS_POPUP_REF = null;
            }
        });

        try {
            localStorage.setItem(STORAGE_LOCK_SAMPLE_PATHS_POPUP, "1");
        } catch (e) {}

        var dots = 1;
        var loadingInterval = setInterval(function() {
            if (!LOCK_SAMPLE_PATHS_POPUP_REF || !document.body.contains(LOCK_SAMPLE_PATHS_POPUP_REF.element)) {
                clearInterval(loadingInterval);
                return;
            }
            dots = dots + 1;
            if (dots > 3) {
                dots = 1;
            }
            var text = "Running";
            var d = 0;
            while (d < dots) {
                text = text + ".";
                d = d + 1;
            }
            if (loadingAnimation) {
                loadingAnimation.textContent = text;
            }
        }, 500);

        var successCount = 0;
        var failCount = 0;
        var j = 0;
        while (j < unlockedPaths.length) {
            var path = unlockedPaths[j];
            log("Processing (" + String(j + 1) + "/" + String(unlockedPaths.length) + "): " + path.name);

            if (progressDiv) {
                progressDiv.textContent = "Processing " + String(j + 1) + "/" + String(unlockedPaths.length) + ": " + path.name;
            }

            var result = await fetchAndLockSamplePath(path.url, path.name);
            if (result.success) {
                successCount = successCount + 1;
            } else {
                failCount = failCount + 1;
            }

            if (countsDiv) {
                countsDiv.innerHTML = "<span style='color:#9f9'>Success: " + String(successCount) + "</span> | <span style='color:#f99'>Failed: " + String(failCount) + "</span>";
            }

            await sleep(500);
            j = j + 1;
        }

        clearInterval(loadingInterval);

        log("processLockSamplePathsPage: completed. Success=" + String(successCount) + " Failed=" + String(failCount));

        if (statusDiv) {
            statusDiv.textContent = "Completed";
            statusDiv.style.color = "#9f9";
        }
        if (loadingAnimation) {
            loadingAnimation.textContent = "All sample paths processed";
        }
        if (progressDiv) {
            progressDiv.textContent = "Processed " + String(unlockedPaths.length) + "/" + String(unlockedPaths.length);
        }

        try {
            localStorage.removeItem(STORAGE_RUN_LOCK_SAMPLE_PATHS);
            localStorage.removeItem(STORAGE_LOCK_SAMPLE_PATHS_POPUP);
            log("processLockSamplePathsPage: flag cleared");
        } catch (e) { }

        var mode = getRunMode();
        if (mode === "all") {
            await sleep(2000);
            if (LOCK_SAMPLE_PATHS_POPUP_REF && LOCK_SAMPLE_PATHS_POPUP_REF.close) {
                LOCK_SAMPLE_PATHS_POPUP_REF.close();
            }
            LOCK_SAMPLE_PATHS_POPUP_REF = null;
            log("processLockSamplePathsPage: continuing to Study Metadata for eligibility lock");
            updateRunAllPopupStatus("Locking Eligibility");
            try {
                localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
            } catch (e) {}
            location.href = STUDY_METADATA_URL + "?autoeliglock=1";
        }
    }

    async function processLockSamplePathDetailPage() {
        log("processLockSamplePathDetailPage: not needed with background approach");
    }

    async function processLockSamplePathUpdatePage() {
        log("processLockSamplePathUpdatePage: not needed with background approach");
    }

    //==========================
    // RUN IMPORT COHORT SUBJECT FEATURE
    //==========================
    // This section contains all functions related to importing cohort subjects from
    // non-screening epochs. This feature automates the process of editing cohort settings,
    // importing subjects from other cohorts, and activating volunteers. Functions handle
    // epoch selection, cohort editing, import modal interactions, and volunteer activation.
    //==========================

    function ImportCohortSubjectsFunctions() {}
    // Clear used letters.
    function clearUsedLetters() {
        try {
            localStorage.removeItem("activityPlanState.cohortAdd.usedLetters");
        } catch (e) {}
    }

    // Pick a random unused letter from a-z (excluding 'u' and 'x' which have no results).
    function pickRandomUnusedLetter(allLetters, usedLetters) {
        var excludedLetters = ["u", "x"];
        var unused = [];
        var i = 0;
        while (i < allLetters.length) {
            var letter = allLetters[i];
            // Skip excluded letters
            var isExcluded = false;
            var k = 0;
            while (k < excludedLetters.length) {
                if (excludedLetters[k] === letter) {
                    isExcluded = true;
                    break;
                }
                k = k + 1;
            }
            if (isExcluded) {
                i = i + 1;
                continue;
            }
            var found = false;
            var j = 0;
            while (j < usedLetters.length) {
                if (usedLetters[j] === letter) {
                    found = true;
                    break;
                }
                j = j + 1;
            }
            if (!found) {
                unused.push(letter);
            }
            i = i + 1;
        }
        if (unused.length === 0) {
            return null;
        }
        var randomIdx = Math.floor(Math.random() * unused.length);
        return unused[randomIdx];
    }

    // Parse volunteer id from a select2 chosen label text like "(ID: 123)".
    function extractVolunteerIdFromChosenText(txt) {
        if (typeof txt !== "string") {
            return "";
        }
        var m = txt.match(/\(ID:\s*(\d+)\)/i);
        if (m && m[1]) {
            return m[1];
        }
        return "";
    }
    // Store used letters for cohort add process.
    function setUsedLetters(letters) {
        try {
            localStorage.setItem("activityPlanState.cohortAdd.usedLetters", JSON.stringify(letters));
        } catch (e) {}
    }

    // Get used letters for cohort add process.
    function getUsedLetters() {
        var raw = null;
        try {
            raw = localStorage.getItem("activityPlanState.cohortAdd.usedLetters");
        } catch (e) {}
        if (!raw) {
            return [];
        }
        var arr = [];
        try {
            arr = JSON.parse(raw);
        } catch (e2) {}
        if (Array.isArray(arr)) {
            return arr;
        }
        return [];
    }

    async function processEpochShowPageForImport() {
        if (isPaused()) {
            log("Paused; exiting processEpochShowPageForImport");
            return;
        }
        var auto = getQueryParam("autoepochimport");
        var go = auto === "1" || getRunMode() === "epochImport";
        if (!go) {
            return;
        }
        var anchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        if (anchors.length === 0) {
            var idx = getNonScrnEpochIndex();
            idx = idx + 1;
            setNonScrnEpochIndex(idx);
            location.href = STUDY_SHOW_URL;
            log("Epoch has no cohorts; advancing to next epoch");
            return;
        }
        var target = anchors[0];
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            var idx2 = getNonScrnEpochIndex();
            idx2 = idx2 + 1;
            setNonScrnEpochIndex(idx2);
            location.href = STUDY_SHOW_URL;
            log("Cohort href missing; advancing to next epoch");
            return;
        }
        location.href = location.origin + href + "?autocohortimport=1";
        log("Routing to first cohort show for import");
    }

    function setImportDone(cohortId) {
        log("setImportDone: cohortId=" + String(cohortId));
        var map = getImportDoneMap();
        map[String(cohortId)] = 1;
        try {
            localStorage.setItem(STORAGE_IMPORT_DONE_MAP, JSON.stringify(map));
        } catch (e) {}
    }

    function getImportDoneMap() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_IMPORT_DONE_MAP);
        } catch (e) {}
        var map = {};
        if (raw) {
            try {
                map = JSON.parse(raw);
            } catch (e2) {}
        }
        if (typeof map !== "object" || map === null) {
            map = {};
        }
        return map;
    }

    function getNonScrnEpochIndex() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_NON_SCRN_EPOCH_INDEX);
        } catch (e) {}
        if (!raw) {
            return 0;
        }
        var n = parseInt(raw, 10);
        if (isNaN(n)) {
            return 0;
        }
        return n;
    }
    function setImportSubjectIds(ids) {
        try {
            localStorage.setItem(STORAGE_IMPORT_SUBJECT_IDS, JSON.stringify(ids));
        } catch (e) {}
    }
    function clearImportSubjectIds() {
        try {
            localStorage.removeItem(STORAGE_IMPORT_SUBJECT_IDS);
        } catch (e) {}
    }
    function getImportSubjectIds() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_IMPORT_SUBJECT_IDS);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        try {
            var parsed = JSON.parse(raw);
            if (Array.isArray(parsed)) {
                return parsed;
            }
        } catch (e) {}
        return [];
    }

    function extractVolunteerIdFromFreeText(txt) {
        log("extractVolunteerIdFromFreeText: input='" + String(txt) + "'");
        if (typeof txt !== "string") {
            return "";
        }
        var s = txt + "";
        var m1 = s.match(/\(ID:\s*(\d+)\)/i);
        if (m1 && m1[1]) {
            log("extractVolunteerIdFromFreeText: matched paren ID=" + String(m1[1]));
            return m1[1];
        }
        var m2 = s.match(/ID\s*:\s*(\d+)/i);
        if (m2 && m2[1]) {
            log("extractVolunteerIdFromFreeText: matched bare ID=" + String(m2[1]));
            return m2[1];
        }
        log("extractVolunteerIdFromFreeText: no match");
        return "";
    }

    function isImportDone(cohortId) {
        var map = getImportDoneMap();
        var flag = !!map[String(cohortId)];
        log("isImportDone: cohortId=" + String(cohortId) + " result=" + String(flag));
        return flag;
    }

    // Set cohort-run guard state to coordinate multi-step flows.
    function setCohortGuard(state) {
        try {
            localStorage.setItem("activityPlanState.cohortAdd.guard", String(state));
            log("CohortGuard=" + String(state));
        } catch (e) {}
    }

    // Retrieve cohort-run guard state.
    function getCohortGuard() {
        var raw = null;
        try {
            raw = localStorage.getItem("activityPlanState.cohortAdd.guard");
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }


    function markImportDoneIfSuccessOnLoad(cohortId) {
        log("markImportDoneIfSuccessOnLoad: start cohortId=" + String(cohortId));
        var success = hasSuccessAlert();
        log("markImportDoneIfSuccessOnLoad: successAlert=" + String(success));
        if (success && cohortId && cohortId.length > 0) {
            setImportDone(cohortId);
            log("markImportDoneIfSuccessOnLoad: setImportDone for cohortId=" + String(cohortId));
            return true;
        }
        return false;
    }
    // Wait for the cohort/list table element to appear.
    async function waitForListTable(timeoutMs) {
        var el = await waitForSelector('table#listTable', timeoutMs);
        return el;
    }

    // Find cohort table row referencing a given volunteer id.
    function findCohortRowByVolunteerId(volId) {
        var table = document.querySelector('table#listTable');
        if (!table) {
            return null;
        }
        var rows = table.querySelectorAll('tr');
        var i = 0;
        while (i < rows.length) {
            var row = rows[i];
            var aList = row.querySelectorAll('a[href^="/secure/volunteers/manage/show/"]');
            var j = 0;
            var found = false;
            while (j < aList.length) {
                var a = aList[j];
                var href = a.getAttribute("href") + "";
                var m = href.match(/\/show\/(\d+)\//);
                var id = "";
                if (m && m[1]) {
                    id = m[1];
                } else {
                    m = href.match(/\/show\/(\d+)$/);
                    if (m && m[1]) {
                        id = m[1];
                    }
                }
                var same = id === String(volId);
                if (same) {
                    found = true;
                    break;
                }
                j = j + 1;
            }
            if (found) {
                return row;
            }
            i = i + 1;
        }
        return null;
    }
    // Return the row action button element for cohort actions.
    function getRowActionButton(row) {
        if (!row) {
            return null;
        }
        var btn = row.querySelector('button[id^="cohortAssignmentAction_"]');
        return btn;
    }


    // Find 'Activate Plan' link in a row's dropdown menu.
    function getMenuLinkActivatePlan(row) {
        if (!row) {
            return null;
        }
        var menu = row.querySelector('ul.dropdown-menu');
        if (!menu) {
            return null;
        }
        var links = menu.querySelectorAll('a');
        var i = 0;
        while (i < links.length) {
            var a = links[i];
            var txt = (a.textContent + "").trim().toLowerCase();
            var disabled = a.parentElement && a.parentElement.classList.contains('disabled');
            var isPlan = txt.indexOf('activate plan') !== -1;
            if (isPlan && !disabled) {
                return a;
            }
            var onclick = a.getAttribute('onclick') + "";
            var hasOn = onclick.indexOf('activatePlan(') !== -1;
            if (hasOn) {
                return a;
            }
            i = i + 1;
        }
        return null;
    }


    // Find 'Activate Volunteer' link in a row's dropdown menu.
    function getMenuLinkActivateVolunteer(row) {
        if (!row) {
            return null;
        }
        var menu = row.querySelector('ul.dropdown-menu');
        if (!menu) {
            return null;
        }
        var idLink = menu.querySelector('li a[id^="cohortAssignmentActivateSubjectLink_"]');
        if (idLink) {
            return idLink;
        }
        var links = menu.querySelectorAll('a');
        var i = 0;
        while (i < links.length) {
            var a = links[i];
            var txt = (a.textContent + "").trim().toLowerCase();
            var disabled = a.parentElement && a.parentElement.classList.contains('disabled');
            var isVol = txt.indexOf('activate volunteer') !== -1;
            if (isVol && !disabled) {
                return a;
            }
            var onclick = a.getAttribute('onclick') + "";
            var hasOn = onclick.indexOf('activateVolunteer(') !== -1;
            if (hasOn) {
                return a;
            }
            i = i + 1;
        }
        return null;
    }

    // Wait for Activate Volunteer link to become available for an id.
    async function waitForActivateVolunteerById(volId, timeoutMs) {
        var start = Date.now();
        var attempts = 0;
        while (Date.now() - start < timeoutMs) {
            if (isPaused()) {
                log("Activation: paused; stopping search for volunteerId=" + String(volId));
                return null;
            }
            var row = findCohortRowByVolunteerId(volId);
            if (row) {
                var actionBtn = getRowActionButton(row);
                if (actionBtn) {
                    actionBtn.click();
                    await sleep(250);
                    var linkById = row.querySelector('ul.dropdown-menu li a[id^="cohortAssignmentActivateSubjectLink_"]');
                    var link = null;
                    if (linkById) {
                        link = linkById;
                    }
                    if (!link) {
                        link = getMenuLinkActivateVolunteer(row);
                    }
                    if (link) {
                        log("Activate Volunteer link found after attempts=" + String(attempts));
                        return { row: row, actionBtn: actionBtn, link: link };
                    }
                }
            }
            attempts = attempts + 1;
            var waited = Date.now() - start;
            if (waited >= 1000 && waited % 1000 < 600) {
                log("Activation: still searching for volunteerId=" + String(volId) + " waited=" + String(waited) + "ms");
            }
            await sleep(600);
        }
        log("Activation: row not found for volunteerId=" + String(volId) + " after " + String(timeoutMs) + "ms; skipping");
        return null;
    }

    function parseCohortIdFromUpdateHref(href) {
        log("parseCohortIdFromUpdateHref: href=" + String(href));
        var id = "";
        if (typeof href === "string") {
            var m = href.match(/cohort\/update\/(\d+)/);
            if (m && m[1]) {
                id = m[1];
            } else {
                var tail = href.split("/").pop();
                var n = parseInt(String(tail), 10);
                if (!isNaN(n)) {
                    id = String(n);
                }
            }
        }
        log("parseCohortIdFromUpdateHref: parsed id=" + String(id));
        return id;
    }

    function setCheckboxStateById(id, state) {
        log("setCheckboxStateById: id=" + String(id) + " targetState=" + String(!!state));
        var el = document.getElementById(id);
        if (!el) {
            log("setCheckboxStateById: element not found id=" + String(id));
            return false;
        }
        var before = !!el.checked;
        log("setCheckboxStateById: currentState=" + String(before));
        el.checked = !!state;
        var evt = new Event("change", { bubbles: true });
        el.dispatchEvent(evt);
        var wrap = el.closest("div.checker");
        var span = null;
        if (wrap) {
            span = wrap.querySelector("span");
            if (span) {
                if (state) {
                    span.classList.add("checked");
                } else {
                    span.classList.remove("checked");
                }
            }
        }
        var after = !!el.checked;
        var wrapFlag = wrap ? "1" : "0";
        var spanFlag = span ? "1" : "0";
        log("setCheckboxStateById: afterState=" + String(after) + " wrapper=" + wrapFlag + " span=" + spanFlag);
        return true;
    }

    // Orchestrate cohort add + activation flow and continue to consent if in ALL mode.
    async function processCohortShowPage() {
        log("processCohortShowPage start");
        var auto = getQueryParam("autocohort");
        var mode = getRunMode();
        var go = (mode === "epoch" || mode === "all") || (auto === "1" && mode && mode.length > 0);
        if (!go) {
            log("Cohort page not in run mode");
            return;
        }
        if (isPaused()) {
            log("Paused; skipping cohort automation");
            clearRunMode();
            clearContinueEpoch();
            clearCohortGuard();
            return;
        }
        var g = getCohortGuard();
        log("CohortGuard at start=" + String(g));
        var successOnLoad = hasSuccessAlert();
        if (successOnLoad && g !== "postsave") {
            log("Success alert on load; proceeding to activation flow");
            var listReadyX = await waitForListTable(12000);
            if (!listReadyX) {
                await sleep(800);
            }
            var targetVolX = getLastVolunteerId();
            if (!targetVolX) {
                targetVolX = getLastSelectedVolunteerId();
            }
            log("Using id for activation=" + String(targetVolX));
            var targetRowX = null;
            var waitedRowX = 0;
            var maxWaitRowX = 30000;
            while (waitedRowX < maxWaitRowX) {
                if (isPaused()) {
                    log("Activation: paused; stopping row search for volunteerId=" + String(targetVolX));
                    clearRunMode();
                    clearContinueEpoch();
                    clearCohortGuard();
                    return;
                }
                targetRowX = findCohortRowByVolunteerId(targetVolX);
                if (targetRowX) {
                    break;
                }
                await sleep(300);
                waitedRowX = waitedRowX + 300;
                if (waitedRowX >= 1000 && waitedRowX % 1000 < 300) {
                    log("Activation: still searching for volunteerId=" + String(targetVolX) + " waited=" + String(waitedRowX) + "ms");
                }
            }
            if (!targetRowX) {
                log("Target row not found for id=" + String(targetVolX));
                clearRunMode();
                clearContinueEpoch();
                clearCohortGuard();
                return;
            }
            var actionBtnX = getRowActionButton(targetRowX);
            if (!actionBtnX) {
                log("Action button not found in row");
                clearRunMode();
                clearContinueEpoch();
                clearCohortGuard();
                return;
            }
            actionBtnX.click();
            await sleep(300);
            var planLinkX = getMenuLinkActivatePlan(targetRowX);
            if (!planLinkX) {
                log("Activate Plan link not found");
                clearRunMode();
                clearContinueEpoch();
                clearCohortGuard();
                return;
            }
            planLinkX.click();
            var ok1X = await clickBootboxOk(5000);
            if (!ok1X) {
                await sleep(500);
            }
            var foundX = await waitForActivateVolunteerById(targetVolX, 45000);
            if (!foundX) {
                clearRunMode();
                clearContinueEpoch();
                clearCohortGuard();
                return;
            }
            foundX.link.click();
            var ok2X = await clickBootboxOk(5000);
            if (!ok2X) {
                await sleep(500);
            }
            if (mode === "all") {
                setRunMode("consent");
                await sleep(3000);
                location.href = STUDY_SHOW_URL + "?autoconsent=1";
                updateRunAllPopupStatus("Running ICF Barcode");
                log("Continuing ALL to consent after pause");
                return;
            }
            clearContinueEpoch();
            clearRunMode();
            setCohortGuard("done");
            clearSelectedVolunteerIds();
            // Clear editDoneMap when program completely finishes
            try {
                localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
                log("Cleared editDoneMap on program completion");
            } catch (e) {}
            log("Activation flow completed");
            return;
        }
        log("Edit Prep: Opening Actions dropdown on cohort show");
        var openedEdit = await clickActionsDropdownIfNeeded();
        log("Edit Prep: Actions dropdown opened=" + String(!!openedEdit));
        if (!openedEdit) {
            log("Edit Prep: Actions dropdown not found");
            return;
        }
        log("Edit Prep: Locating Edit link");
        var editLink = document.querySelector('a[href^="/secure/administration/manage/studies/cohort/update/"][data-toggle="modal"]');
        var editHref = "";
        var cohortId = "";
        if (editLink) {
            editHref = editLink.getAttribute("href") + "";
            log("Edit Prep: Edit link found href=" + String(editHref));
            cohortId = parseCohortIdFromUpdateHref(editHref);
        } else {
            log("Edit Prep: Cohort Edit link not found");
        }
        var storageKeyAdd = "activityPlanState.cohortAdd.editDoneMap";
        var addMapRaw = null;
        try {
            addMapRaw = localStorage.getItem(storageKeyAdd);
        } catch (eMapRead) {}
        var addMap = {};
        if (addMapRaw) {
            try {
                addMap = JSON.parse(addMapRaw);
            } catch (eMapParse) {}
        }
        if (typeof addMap !== "object" || addMap === null) {
            addMap = {};
        }
        var alreadyEditedAdd = false;
        if (cohortId && cohortId.length > 0) {
            alreadyEditedAdd = !!addMap[String(cohortId)];
        }
        log("Edit Prep Guard: alreadyEditedAdd=" + String(alreadyEditedAdd) + " cohortId=" + String(cohortId));
        if (editLink && !alreadyEditedAdd) {
            log("Edit Prep: Clicking Edit link");
            editLink.click();
            var editModal = await waitForSelector("#ajaxModal, .modal", 6000);
            log("Edit Prep: Edit modal present=" + String(!!editModal));
            if (!editModal) {
                log("Edit Prep: Edit modal did not open");
            } else {
                log("Edit Prep: Checking form-group structure");
                var groups = editModal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
                var groupCount = groups ? groups.length : 0;
                log("Edit Prep: form-group count=" + String(groupCount));
                if (!groups || groups.length === 0) {
                    var waited = 0;
                    var maxWait = 6000;
                    var step = 300;
                    while (waited < maxWait) {
                        if (isPaused()) {
                            log("Paused; exiting inside form-group wait");
                            return;
                        }
                        groups = editModal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
                        groupCount = groups ? groups.length : 0;
                        if (groupCount > 0) {
                            break;
                        }
                        await sleep(step);
                        waited = waited + step;
                    }
                    log("Edit Prep: post-wait form-group count=" + String(groupCount));
                }
                if (!groups || groups.length === 0) {
                    var idsList = [
                        "subjectInitiation",
                        "sourceVolunteerDatabase",
                        "sourceAppointments",
                        "sourceAppointmentsCohort",
                        "sourceScreeningCohorts",
                        "sourceLeadInCohorts",
                        "sourceRandomizationCohorts",
                        "allowSubjectsActiveInCohorts",
                        "allowSubjectsActiveInStudies",
                        "requireVolunteerRecruitment",
                        "allowRecruitmentEligible",
                        "allowRecruitmentIdentified",
                        "allowRecruitmentIneligible",
                        "allowRecruitmentRemoved",
                        "allowEligibilityEligible",
                        "allowEligibilityPending",
                        "allowEligibilityIneligible",
                        "allowEligibilityUnspecified",
                        "allowStatusActive",
                        "allowStatusComplete",
                        "allowStatusTerminated",
                        "allowStatusWithdrawn",
                        "requireInformedConsent",
                        "requireOverVolunteeringCheck"
                    ];
                    var foundAny = false;
                    var checkWait = 0;
                    var checkMax = 6000;
                    while (checkWait < checkMax) {
                        if (isPaused()) {
                            log("Paused; exiting inside checkbox presence wait");
                            return;
                        }
                        var k = 0;
                        while (k < idsList.length) {
                            var el = document.getElementById(idsList[k]);
                            if (el) {
                                foundAny = true;
                                break;
                            }
                            k = k + 1;
                        }
                        if (foundAny) {
                            break;
                        }
                        await sleep(300);
                        checkWait = checkWait + 300;
                    }
                    log("Edit Prep: checkbox presence=" + String(foundAny));
                    if (!foundAny) {
                        log("Edit Prep: Edit modal inputs not present");
                    }
                }
                log("Edit Prep: Applying Subject/Volunteer Source checkboxes");
                var ok6 = true;
                ok6 = ok6 && setCheckboxStateById("subjectInitiation", true);
                ok6 = ok6 && setCheckboxStateById("sourceVolunteerDatabase", true);
                ok6 = ok6 && setCheckboxStateById("sourceAppointments", true);
                ok6 = ok6 && setCheckboxStateById("sourceAppointmentsCohort", true);
                ok6 = ok6 && setCheckboxStateById("sourceScreeningCohorts", true);
                ok6 = ok6 && setCheckboxStateById("sourceLeadInCohorts", true);
                ok6 = ok6 && setCheckboxStateById("sourceRandomizationCohorts", true);
                var ok7 = true;
                ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInCohorts", true);
                ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInStudies", true);
                var ok8 = true;
                ok8 = ok8 && setCheckboxStateById("requireVolunteerRecruitment", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentEligible", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentIdentified", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentIneligible", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentRemoved", false);
                var ok9 = true;
                ok9 = ok9 && setCheckboxStateById("allowEligibilityEligible", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityPending", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityIneligible", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityUnspecified", true);
                var ok10 = true;
                ok10 = ok10 && setCheckboxStateById("allowStatusActive", true);
                ok10 = ok10 && setCheckboxStateById("allowStatusComplete", false);
                ok10 = ok10 && setCheckboxStateById("allowStatusTerminated", false);
                ok10 = ok10 && setCheckboxStateById("allowStatusWithdrawn", false);
                var ok11 = true;
                ok11 = ok11 && setCheckboxStateById("requireInformedConsent", false);
                ok11 = ok11 && setCheckboxStateById("requireOverVolunteeringCheck", false);
                var reason = editModal.querySelector('textarea#reasonForChange');
                if (reason) {
                    reason.value = "Test";
                    var evtR = new Event("input", { bubbles: true });
                    reason.dispatchEvent(evtR);
                }
                log("Edit Prep: Saving Edit modal");
                var saveBtnEdit = await waitForSelector("#actionButton", 5000);
                if (saveBtnEdit) {
                    if (cohortId && cohortId.length > 0) {
                        addMap[String(cohortId)] = 1;
                        try {
                            localStorage.setItem(storageKeyAdd, JSON.stringify(addMap));
                        } catch (eMapStore) {}
                    }
                    saveBtnEdit.click();
                    log("Edit Prep: Edit modal Save clicked");
                    await sleep(2000);
                    log("Edit Prep: Edit modal post-save wait done");
                } else {
                    log("Edit Prep: Edit modal Save button not found");
                }
            }
        } else {
            log("Edit Prep Guard: skipping Edit modal for cohortId=" + String(cohortId));
        }
        setCohortGuard("inprogress");
        clearLastVolunteerId();
        clearSelectedVolunteerIds();
        clearUsedLetters();
        var letters = "abcdefghijklmnopqrstuvwxyz".split("");
        var usedLetters = getUsedLetters();
        var savedOk = false;
        var currentVolunteerId = "";
        var attempts = 0;
        var maxAttempts = 26;
        while (attempts < maxAttempts && usedLetters.length < letters.length) {
            var opened = await clickActionsDropdownIfNeeded();
            if (!opened) {
                return;
            }
            var addLink = await waitForSelector('a#addCohortAssignmentButton[data-toggle="modal"]', 3000);
            if (!addLink) {
                addLink = document.querySelector('a[href^="/secure/study/cohortassign/manage/save/"][data-toggle="modal"]');
            }
            if (!addLink) {
                log("Add Cohort Assignment link not found");
                return;
            }
            addLink.click();
            log("Add Cohort Assignment clicked");
            var modal = await waitForSelector("#ajaxModal, .modal", 5000);
            if (!modal) {
                log("Modal not found");
                return;
            }
            var planSel = await waitForSelector('select#activityPlan', 5000);
            if (!planSel) {
                log("ActivityPlan select not found");
                return;
            }
            var opts = planSel.querySelectorAll("option");
            var chosen = null;
            var j = 0;
            while (j < opts.length) {
                var op = opts[j];
                var val = op.value + "";
                if (val && val.length > 0) {
                    chosen = op;
                    break;
                }
                j = j + 1;
            }
            if (chosen) {
                planSel.value = chosen.value;
                var evt1 = new Event("change", { bubbles: true });
                planSel.dispatchEvent(evt1);
                log("ActivityPlan chosen value=" + String(chosen.value));
            }
            var searchSel = await waitForSelector('select#cohortAssignmentSearch', 5000);
            if (!searchSel) {
                log("cohortAssignmentSearch not found");
                return;
            }
            var needAll = searchSel.value !== "AllVolunteers";
            if (needAll) {
                searchSel.value = "AllVolunteers";
                var evt2 = new Event("change", { bubbles: true });
                searchSel.dispatchEvent(evt2);
                log("Search set to AllVolunteers");
            }
            var s2container = modal.querySelector('#s2id_volunteer');
            if (!s2container) {
                s2container = modal.querySelector('.select2-container.form-control.select2');
            }
            if (!s2container) {
                log("Select2 container not found");
                return;
            }
            var s2choice = s2container.querySelector('a.select2-choice');
            if (s2choice) {
                s2choice.click();
                await sleep(150);
            }
            var focusser = s2container.querySelector('input.select2-focusser');
            if (focusser) {
                focusser.focus();
                var kd = new KeyboardEvent("keydown", { bubbles: true, key: "ArrowDown", keyCode: 40 });
                focusser.dispatchEvent(kd);
                focusser.click();
                await sleep(150);
            }
            var s2drop = document.querySelector('#select2-drop.select2-drop-active');
            if (!s2drop) {
                s2drop = s2container.querySelector('.select2-drop');
            }
            if (!s2drop) {
                s2choice = s2container.querySelector('a.select2-choice');
                if (s2choice) {
                    s2choice.click();
                }
                s2drop = await waitForSelector('#select2-drop.select2-drop-active', 2000);
                if (!s2drop) {
                    s2drop = s2container.querySelector('.select2-drop');
                }
            }
            if (!s2drop) {
                log("Select2 drop not found");
                return;
            }
            var s2input = s2drop.querySelector('input.select2-input');
            if (!s2input) {
                s2input = await waitForSelector('#select2-drop.select2-drop-active input.select2-input', 2000);
            }
            if (!s2input) {
                s2input = s2container.querySelector('input.select2-input');
            }
            if (!s2input) {
                log("Select2 input not found");
                return;
            }
            var letter = pickRandomUnusedLetter(letters, usedLetters);
            if (!letter) {
                log("No unused letters remaining; exiting");
                break;
            }
            usedLetters.push(letter);
            setUsedLetters(usedLetters);
            s2input.value = letter;
            var inpEvt = new Event("input", { bubbles: true });
            s2input.dispatchEvent(inpEvt);
            var keyEvt = new KeyboardEvent("keyup", { bubbles: true, key: letter, keyCode: letter.toUpperCase().charCodeAt(0) });
            s2input.dispatchEvent(keyEvt);
            log("Typed random letter=" + String(letter) + " (used: " + JSON.stringify(usedLetters) + ")");
            var selectionConfirmed = false;
            var confirmWait = 0;
            var confirmMax = 12000;
            while (confirmWait < confirmMax) {
                var enterDown = new KeyboardEvent("keydown", { bubbles: true, key: "Enter", keyCode: 13 });
                s2input.dispatchEvent(enterDown);
                var enterUp = new KeyboardEvent("keyup", { bubbles: true, key: "Enter", keyCode: 13 });
                s2input.dispatchEvent(enterUp);
                await sleep(400);
                var containerClass = s2container.getAttribute("class") + "";
                var hasAllow = containerClass.indexOf("select2-allowclear") !== -1;
                var notOpen = containerClass.indexOf("select2-dropdown-open") === -1;
                var chosenEl = null;
                chosenEl = s2container.querySelector('.select2-chosen');
                if (!chosenEl) {
                    chosenEl = s2container.querySelector('[id^="select2-chosen-"]');
                }
                var chosenText = "";
                if (chosenEl) {
                    chosenText = (chosenEl.textContent + "").trim();
                }
                var notSearch = chosenText.trim().toLowerCase() !== "search";
                if (hasAllow && notOpen && notSearch) {
                    selectionConfirmed = true;
                    var volIdParsed = extractVolunteerIdFromChosenText(chosenText);
                    if (volIdParsed) {
                        currentVolunteerId = volIdParsed;
                        appendSelectedVolunteerId(currentVolunteerId);
                        log("Appended selected id=" + String(currentVolunteerId) + "; listLen=" + String(getSelectedVolunteerIds().length));
                    }
                    log("Selection confirmed; chosenText=" + chosenText + "; parsedId=" + String(currentVolunteerId));
                    break;
                }
                confirmWait = confirmWait + 400;
            }
            if (!selectionConfirmed) {
                var cancelBtn0 = getModalCancelButton(modal);
                if (cancelBtn0) {
                    cancelBtn0.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                attempts = attempts + 1;
                log("Selection not confirmed; trying next random letter");
                continue;
            }
            var saveBtn = await waitForSelector('button#actionButton.btn.green[type="button"]', 5000);
            if (!saveBtn) {
                saveBtn = document.querySelector('button#actionButton');
            }
            if (!saveBtn) {
                log("Save button not found");
                return;
            }
            saveBtn.click();
            log("Save clicked; current volunteerId=" + String(currentVolunteerId));
            var successImmediate = hasSuccessAlert();
            if (successImmediate) {
                savedOk = true;
                log("Immediate success alert");
                var stillOpen = document.querySelector('#ajaxModal, .modal');
                if (stillOpen) {
                    var style = window.getComputedStyle(stillOpen);
                    var visible = style.display !== "none" && style.visibility !== "hidden";
                    if (visible) {
                        var cancelBtnNow = getModalCancelButton(stillOpen);
                        if (cancelBtnNow) {
                            cancelBtnNow.click();
                            await waitUntilHidden("#ajaxModal, .modal", 5000);
                        }
                    }
                }
                var lastSelected = getLastSelectedVolunteerId();
                if (lastSelected) {
                    setLastVolunteerId(lastSelected);
                }
                setCohortGuard("postsave");
                break;
            }
            await sleep(600);
            var failedImmediate = hasValidationError();
            if (failedImmediate) {
                var cancelBtn = getModalCancelButton(modal);
                if (cancelBtn) {
                    cancelBtn.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                attempts = attempts + 1;
                log("Immediate validation error; trying next random letter");
                continue;
            }
            var waitedAlerts = 0;
            var maxAlerts = 12000;
            var successDetected = false;
            var failureDetected = false;
            while (waitedAlerts < maxAlerts) {
                await sleep(400);
                var sOk = hasSuccessAlert();
                var sFail = hasValidationError();
                if (sOk) {
                    successDetected = true;
                    break;
                }
                if (sFail) {
                    failureDetected = true;
                    break;
                }
                waitedAlerts = waitedAlerts + 400;
            }
            if (successDetected) {
                savedOk = true;
                log("Delayed success alert");
                var stillOpen2 = document.querySelector('#ajaxModal, .modal');
                if (stillOpen2) {
                    var style2 = window.getComputedStyle(stillOpen2);
                    var visible2 = style2.display !== "none" && style2.visibility !== "hidden";
                    if (visible2) {
                        var cancelBtn2 = getModalCancelButton(stillOpen2);
                        if (cancelBtn2) {
                            cancelBtn2.click();
                            await waitUntilHidden("#ajaxModal, .modal", 5000);
                        }
                    }
                }
                var lastSelected2 = getLastSelectedVolunteerId();
                if (lastSelected2) {
                    setLastVolunteerId(lastSelected2);
                }
                setCohortGuard("postsave");
                break;
            }
            if (failureDetected) {
                var cancelBtn3 = getModalCancelButton(modal);
                if (cancelBtn3) {
                    cancelBtn3.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                idx = idx + 1;
                log("Validation error after wait; next letter index=" + String(idx));
                continue;
            }
            var cancelBtn4 = getModalCancelButton(modal);
            if (cancelBtn4) {
                cancelBtn4.click();
                await waitUntilHidden("#ajaxModal, .modal", 5000);
            }
            idx = idx + 1;
            log("No alert; closing modal; next letter index=" + String(idx));
        }
        if (!savedOk) {
            log("No volunteer saved; exiting");
            return;
        }
        var listReady = await waitForListTable(15000);
        if (!listReady) {
            await sleep(800);
        }
        var targetVolId = getLastVolunteerId();
        if (!targetVolId) {
            targetVolId = getLastSelectedVolunteerId();
        }
        log("Proceeding to activation for id=" + String(targetVolId));
        var targetRow = null;
        var waitedRow = 0;
        var maxWaitRow = 30000;
        while (waitedRow < maxWaitRow) {
            if (isPaused()) {
                log("Activation: paused; stopping row search for volunteerId=" + String(targetVolId));
                clearRunMode();
                clearContinueEpoch();
                clearCohortGuard();
                return;
            }
            targetRow = findCohortRowByVolunteerId(targetVolId);
            if (targetRow) {
                break;
            }
            await sleep(300);
            waitedRow = waitedRow + 300;
            if (waitedRow >= 1000 && waitedRow % 1000 < 300) {
                log("Activation: still searching for volunteerId=" + String(targetVolId) + " waited=" + String(waitedRow) + "ms");
            }
        }
        if (!targetRow) {
            log("Activation row not found for id=" + String(targetVolId));
            return;
        }
        var actionBtn = getRowActionButton(targetRow);
        if (!actionBtn) {
            log("Row action button not found");
            return;
        }
        actionBtn.click();
        await sleep(300);
        var planLink = getMenuLinkActivatePlan(targetRow);
        if (!planLink) {
            log("Activate Plan link not found");
            return;
        }
        planLink.click();
        var ok1 = await clickBootboxOk(5000);
        if (!ok1) {
            await sleep(500);
        }
        var found = await waitForActivateVolunteerById(targetVolId, 45000);
        if (!found) {
            if (isPaused()) {
                log("Activation: paused during waitForActivateVolunteerById; clearing state");
            }
            clearRunMode();
            clearContinueEpoch();
            clearCohortGuard();
            return;
        }
        found.link.click();
        var ok2 = await clickBootboxOk(5000);
        if (!ok2) {
            await sleep(500);
        }
        if (getRunMode() === "all") {
            setRunMode("consent");
            updateRunAllPopupStatus("Running ICF Barcode");
            await sleep(3000);
            location.href = STUDY_SHOW_URL + "?autoconsent=1";
            log("Continuing ALL to consent after pause");
            return;
        }
        clearContinueEpoch();
        clearRunMode();
        setCohortGuard("done");
        clearSelectedVolunteerIds();
        clearUsedLetters();
        // Clear editDoneMap when program completely finishes
        try {
            localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
            log("Cleared editDoneMap on program completion");
        } catch (e) {}
        log("Activation completed and run state cleared");
    }

    // Find the modal cancel/dismiss button within a root or globally.
    function getModalCancelButton(root) {
        var btn = null;
        if (root) {
            btn = root.querySelector('button.btn.default[data-dismiss="modal"]');
        } else {
            btn = document.querySelector('#ajaxModal button.btn.default[data-dismiss="modal"], .modal button.btn.default[data-dismiss="modal"]');
        }
        return btn;
    }

    // Detect a validation error alert on the page.
    function hasValidationError() {
        var el = document.querySelector('div.alert.alert-danger.alert-dismissable');
        if (!el) {
            return false;
        }
        var t = el.textContent + "";
        var s = t.trim().toLowerCase();
        var match = s.indexOf("please correct the validation errors") !== -1;
        if (match) {
            return true;
        }
        return false;
    }

    function getCohortEditDoneMap() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_IMPORT_COHORT_EDIT_DONE);
        } catch (e) {}
        var map = {};
        if (raw) {
            try {
                map = JSON.parse(raw);
            } catch (e2) {}
        }
        if (typeof map !== "object" || map === null) {
            map = {};
        }
        return map;
    }


    function setCohortEditDone(cohortId) {
        log("setCohortEditDone: set cohortId=" + String(cohortId));
        var map = getCohortEditDoneMap();
        map[String(cohortId)] = 1;
        try {
            localStorage.setItem(STORAGE_IMPORT_COHORT_EDIT_DONE, JSON.stringify(map));
            log("setCohortEditDone: stored");
        } catch (e) {}
    }

    function isCohortEditDone(cohortId) {
        log("isCohortEditDone: check cohortId=" + String(cohortId));
        var map = getCohortEditDoneMap();
        var flag = !!map[String(cohortId)];
        log("isCohortEditDone: result=" + String(flag));
        return flag;
    }

    function hasNoAssignmentsInImportModal(modalRoot) {
        log("hasNoAssignmentsInImportModal: start");
        if (!modalRoot) {
            log("hasNoAssignmentsInImportModal: modalRoot null");
            return false;
        }
        var tbody = modalRoot.querySelector('tbody#cohortAssignmentTableBody');
        var hasTbody = !!tbody;
        log("hasNoAssignmentsInImportModal: hasTbody=" + String(hasTbody));
        if (!tbody) {
            return false;
        }
        var cell = tbody.querySelector('tr > td[colspan="9"]');
        var hasCell = !!cell;
        log("hasNoAssignmentsInImportModal: hasCell=" + String(hasCell));
        if (!cell) {
            return false;
        }
        var text = (cell.textContent + "").trim().toLowerCase();
        var match = text.indexOf("there are no assignments that match the filter criteria.") !== -1;
        log("hasNoAssignmentsInImportModal: match=" + String(match) + " text='" + String(text) + "'");
        return match;
    }


    async function tryConfirmSelect2Inputs(modalRoot) {
        log("tryConfirmSelect2Inputs: start");
        let inputs = [];
        let i1 = modalRoot.querySelector('input#s2id_autogen34_search');
        if (i1) {
            inputs.push(i1);
        }
        let i2 = modalRoot.querySelector('input#s2id_autogen33');
        if (i2) {
            inputs.push(i2);
        }
        let generic = modalRoot.querySelectorAll('input.select2-input');
        if (generic && generic.length > 0) {
            let gIdx = 0;
            while (gIdx < generic.length) {
                inputs.push(generic[gIdx]);
                gIdx = gIdx + 1;
            }
        }
        if (inputs.length === 0) {
            log("tryConfirmSelect2Inputs: no select2 inputs found");
            return false;
        }
        let idx = 0;
        while (idx < inputs.length) {
            let el = inputs[idx];
            el.focus();
            await sleep(150);
            let kd = new KeyboardEvent("keydown", { bubbles: true, key: "Enter", keyCode: 13 });
            el.dispatchEvent(kd);
            let ku = new KeyboardEvent("keyup", { bubbles: true, key: "Enter", keyCode: 13 });
            el.dispatchEvent(ku);
            log("tryConfirmSelect2Inputs: pressed Enter on input index=" + String(idx));
            await sleep(300);
            idx = idx + 1;
        }
        log("tryConfirmSelect2Inputs: done");
        return true;
    }

    async function processCohortShowPageImportNonScrn() {
        log("processCohortShowPageImportNonScrn: start");
        if (isPaused()) {
            log("Paused; exiting processCohortShowPageImportNonScrn");
            return;
        }
        var auto = getQueryParam("autocohortimport");
        var go = (auto === "1"
                  || getRunMode() === "epochImport");
        log("processCohortShowPageImportNonScrn: auto=" + String(auto) + " runMode=" + String(getRunMode()) + " go=" + String(go));
        if (!go) {
            log("processCohortShowPageImportNonScrn: not in import mode; exiting");
            return;
        }
        log("Step 1/12: Opening Actions dropdown on cohort show");
        var opened = await clickActionsDropdownIfNeeded();
        log("Step 1/12: Actions dropdown opened=" + String(!!opened));
        if (!opened) {
            log("processCohortShowPageImportNonScrn: Actions dropdown not found");
            return;
        }
        log("Step 2/12: Locating Edit link");
        var editLink = document.querySelector('a[href^="/secure/administration/manage/studies/cohort/update/"][data-toggle="modal"]');
        var editHref = "";
        var cohortId = "";
        if (editLink) {
            editHref = editLink.getAttribute("href") + "";
            log("Step 2/12: Edit link found href=" + String(editHref));
            cohortId = parseCohortIdFromUpdateHref(editHref);
            log("Step 2/12: Parsed cohortId=" + String(cohortId));
        } else {
            log("Step 2/12: Cohort Edit link not found");
            return;
        }
        var alreadyEdited = false;
        if (cohortId && cohortId.length > 0) {
            alreadyEdited = isCohortEditDone(cohortId);
        }
        log("Guard: alreadyEdited=" + String(alreadyEdited) + " cohortId=" + String(cohortId));
        var successMarked = markImportDoneIfSuccessOnLoad(cohortId);
        log("Guard: successMarkedFromAlert=" + String(successMarked));
        var importDoneEarly = isImportDone(cohortId);
        log("Guard: importDoneEarly=" + String(importDoneEarly) + " cohortId=" + String(cohortId));
        if (!alreadyEdited) {
            log("Step 3/12: Clicking Edit link");
            editLink.click();
            var modal = await waitForSelector("#ajaxModal, .modal", 6000);
            if (!modal) {
                log("processCohortShowPageImportNonScrn: Edit modal did not open");
                return;
            }
            var groups = modal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
            var groupCount = groups ? groups.length : 0;
            if (!groups
                || groups.length === 0) {
                var waited = 0;
                var maxWait = 6000;
                var step = 300;
                while (waited < maxWait) {
                    if (isPaused()) {
                        log("Paused; exiting inside form-group wait");
                        return;
                    }
                    groups = modal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
                    groupCount = groups ? groups.length : 0;
                    if (groupCount > 0) {
                        break;
                    }
                    await sleep(step);
                    waited = waited + step;
                }
            }
            if (!groups
                || groups.length === 0) {
                var idsList = [
                    "subjectInitiation",
                    "sourceVolunteerDatabase",
                    "sourceAppointments",
                    "sourceAppointmentsCohort",
                    "sourceScreeningCohorts",
                    "sourceLeadInCohorts",
                    "sourceRandomizationCohorts",
                    "allowSubjectsActiveInCohorts",
                    "allowSubjectsActiveInStudies",
                    "requireVolunteerRecruitment",
                    "allowRecruitmentEligible",
                    "allowRecruitmentIdentified",
                    "allowRecruitmentIneligible",
                    "allowRecruitmentRemoved",
                    "allowEligibilityEligible",
                    "allowEligibilityPending",
                    "allowEligibilityIneligible",
                    "allowEligibilityUnspecified",
                    "allowStatusActive",
                    "allowStatusComplete",
                    "allowStatusTerminated",
                    "allowStatusWithdrawn",
                    "requireInformedConsent",
                    "requireOverVolunteeringCheck"
                ];
                var foundAny = false;
                var checkWait = 0;
                var checkMax = 6000;
                while (checkWait < checkMax) {
                    if (isPaused()) {
                        log("Paused; exiting inside checkbox presence wait");
                        return;
                    }
                    var k = 0;
                    while (k < idsList.length) {
                        var el = document.getElementById(idsList[k]);
                        if (el) {
                            foundAny = true;
                            break;
                        }
                        k = k + 1;
                    }
                    if (foundAny) {
                        break;
                    }
                    await sleep(300);
                    checkWait = checkWait + 300;
                }
                if (!foundAny) {
                    log("processCohortShowPageImportNonScrn: Edit modal inputs not present");
                    return;
                }
            }
            var ok6 = true;
            ok6 = ok6 && setCheckboxStateById("subjectInitiation", false);
            ok6 = ok6 && setCheckboxStateById("sourceVolunteerDatabase", true);
            ok6 = ok6 && setCheckboxStateById("sourceAppointments", true);
            ok6 = ok6 && setCheckboxStateById("sourceAppointmentsCohort", true);
            ok6 = ok6 && setCheckboxStateById("sourceScreeningCohorts", true);
            ok6 = ok6 && setCheckboxStateById("sourceLeadInCohorts", true);
            ok6 = ok6 && setCheckboxStateById("sourceRandomizationCohorts", true);
            var ok7 = true;
            ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInCohorts", true);
            ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInStudies", true);
            var ok8 = true;
            ok8 = ok8 && setCheckboxStateById("requireVolunteerRecruitment", false);
            ok8 = ok8 && setCheckboxStateById("allowRecruitmentEligible", false);
            ok8 = ok8 && setCheckboxStateById("allowRecruitmentIdentified", false);
            ok8 = ok8 && setCheckboxStateById("allowRecruitmentIneligible", false);
            ok8 = ok8 && setCheckboxStateById("allowRecruitmentRemoved", false);
            var ok9 = true;
            ok9 = ok9 && setCheckboxStateById("allowEligibilityEligible", true);
            ok9 = ok9 && setCheckboxStateById("allowEligibilityPending", true);
            ok9 = ok9 && setCheckboxStateById("allowEligibilityIneligible", true);
            ok9 = ok9 && setCheckboxStateById("allowEligibilityUnspecified", true);
            var ok10 = true;
            ok10 = ok10 && setCheckboxStateById("allowStatusActive", true);
            ok10 = ok10 && setCheckboxStateById("allowStatusComplete", false);
            ok10 = ok10 && setCheckboxStateById("allowStatusTerminated", false);
            ok10 = ok10 && setCheckboxStateById("allowStatusWithdrawn", false);
            var ok11 = true;
            ok11 = ok11 && setCheckboxStateById("requireInformedConsent", false);
            ok11 = ok11 && setCheckboxStateById("requireOverVolunteeringCheck", false);
            var reason = modal.querySelector('textarea#reasonForChange');
            var reasonExists = !!reason;
            if (reason) {
                reason.value = "Test";
                var evtR = new Event("input", { bubbles: true });
                reason.dispatchEvent(evtR);
            }
            var saveBtn = await waitForSelector("#actionButton", 5000);
            if (!saveBtn) {
                log("processCohortShowPageImportNonScrn: Edit modal Save button not found");
                return;
            }
            if (cohortId && cohortId.length > 0) {
                log("Step 12/12: Pre-marking cohort edit done id=" + String(cohortId));
                setCohortEditDone(cohortId);
            }
            saveBtn.click();
            log("Edit modal Save clicked");
            await sleep(2000);
            log("Edit modal post-save wait done");
        } else {
            log("Guard: skipping Edit modal for cohortId=" + String(cohortId));
        }
        var importDone = isImportDone(cohortId);
        log("Import Guard: importDone=" + String(importDone) + " cohortId=" + String(cohortId));
        if (!importDone) {
            log("Import Step 1/9: Opening Actions dropdown to find Import");
            var opened2 = await clickActionsDropdownIfNeeded();
            log("Import Step 1/9: Actions dropdown opened=" + String(!!opened2));
            if (!opened2) {
                log("processCohortShowPageImportNonScrn: Assignments actions dropdown not found");
                return;
            }
            log("Import Step 2/9: Locating Import link");
            var importLink = document.querySelector('a[href^="/secure/study/cohortassign/manage/import/"][data-toggle="modal"]');
            if (!importLink) {
                var menu = document.querySelector("ul.dropdown-menu");
                var menuExists = !!menu;
                log("Import Step 2/9: direct import link missing; menu exists=" + String(menuExists));
                if (!menu) {
                    log("processCohortShowPageImportNonScrn: Import link menu not found");
                    return;
                }
                importLink = menu.querySelector('a[href^="/secure/study/cohortassign/manage/import/"][data-toggle="modal"]');
            }
            if (importLink) {
                var impHref = importLink.getAttribute("href") + "";
                log("Import Step 2/9: Import link found href=" + String(impHref));
            } else {
                log("processCohortShowPageImportNonScrn: Import link not found");
                return;
            }
            log("Import Step 3/9: Clicking Import link");
            importLink.click();
            var modalImp = await waitForSelector("#ajaxModal, .modal", 6000);
            log("Import Step 3/9: Import modal present=" + String(!!modalImp));
            if (!modalImp) {
                log("processCohortShowPageImportNonScrn: Import modal did not open");
                return;
            }
            log("Import Step 4/9: Confirming Select2 inputs via Enter");
            var confirmedS2 = await tryConfirmSelect2Inputs(modalImp);
            log("Import Step 4/9: select2 confirmed=" + String(confirmedS2));
            log("Import Step 5/9: Waiting for select#epoch and selecting first non-empty option");
            var epochSel = null;
            var waitedEpoch = 0;
            var maxWaitEpoch = 8000;
            while (waitedEpoch < maxWaitEpoch) {
                if (isPaused()) {
                    log("Paused; exiting during epoch select wait");
                    return;
                }
                epochSel = document.querySelector('#ajaxModal select#epoch, .modal select#epoch, select#epoch');
                var foundEpoch = !!epochSel;
                log("Import Step 5/9: epoch select present=" + String(foundEpoch) + " at t=" + String(waitedEpoch) + "ms");
                if (foundEpoch) {
                    break;
                }
                await sleep(300);
                waitedEpoch = waitedEpoch + 300;
            }
            var epochChosen = false;
            if (epochSel) {
                var eOpts = epochSel.querySelectorAll("option");
                var iE = 0;
                while (iE < eOpts.length) {
                    if (isPaused()) {
                        log("Paused; exiting during epoch option scan");
                        return;
                    }
                    var v = (eOpts[iE].value + "").trim();
                    var txt = (eOpts[iE].textContent + "").trim();
                    log("Import Step 5/9: epoch option idx=" + String(iE) + " value='" + String(v) + "' text='" + String(txt) + "'");
                    if (v.length > 0) {
                        epochSel.value = v;
                        var evtE = new Event("change", { bubbles: true });
                        epochSel.dispatchEvent(evtE);
                        epochChosen = true;
                        log("Import Step 5/9: epoch selected value=" + String(v) + " text='" + String(txt) + "'");
                        break;
                    }
                    iE = iE + 1;
                }
            }
            if (!epochChosen) {
                log("Import Step 5/9: epoch not chosen; exiting");
                return;
            }
            log("Import Step 6/9: Waiting for select#cohort to populate and selecting first non-empty option");
            var cohortSelPicked = null;
            var waitedC = 0;
            var maxWaitC = 12000;
            while (waitedC < maxWaitC) {
                if (isPaused()) {
                    log("Paused; exiting during cohort select wait");
                    return;
                }
                var cohortSels = document.querySelectorAll('#ajaxModal select#cohort, .modal select#cohort, select#cohort');
                var countSels = cohortSels ? cohortSels.length : 0;
                log("Import Step 6/9: cohort selects count=" + String(countSels) + " at t=" + String(waitedC) + "ms");
                if (countSels > 0) {
                    var sIdx = 0;
                    var foundNonEmpty = false;
                    while (sIdx < cohortSels.length) {
                        var cOptsScan = cohortSels[sIdx].querySelectorAll("option");
                        var cCount = cOptsScan ? cOptsScan.length : 0;
                        log("Import Step 6/9: cohort select idx=" + String(sIdx) + " options=" + String(cCount));
                        var ka = 0;
                        while (ka < cCount) {
                            var v2 = (cOptsScan[ka].value + "").trim();
                            var t2 = (cOptsScan[ka].textContent + "").trim();
                            log("Import Step 6/9: option idx=" + String(ka) + " value='" + String(v2) + "' text='" + String(t2) + "'");
                            if (v2.length > 0) {
                                cohortSelPicked = cohortSels[sIdx];
                                foundNonEmpty = true;
                                break;
                            }
                            ka = ka + 1;
                        }
                        if (foundNonEmpty) {
                            break;
                        }
                        sIdx = sIdx + 1;
                    }
                    if (foundNonEmpty) {
                        break;
                    }
                }
                await sleep(400);
                waitedC = waitedC + 400;
            }
            var cohortChosen = false;
            if (cohortSelPicked) {
                var cOpts2 = cohortSelPicked.querySelectorAll("option");
                var iC = 0;
                while (iC < cOpts2.length) {
                    var valC = (cOpts2[iC].value + "").trim();
                    var txtC = (cOpts2[iC].textContent + "").trim();
                    if (valC.length > 0) {
                        cOpts2[iC].selected = true;
                        var evtC = new Event("change", { bubbles: true });
                        cohortSelPicked.dispatchEvent(evtC);
                        cohortChosen = true;
                        log("Import Step 6/9: cohort selected value=" + String(valC) + " text='" + String(txtC) + "'");
                        break;
                    }
                    iC = iC + 1;
                }
            }
            if (!cohortChosen) {
                log("Import Step 6/9: cohort not chosen; exiting");
                return;
            }
            log("Import Step 7/9: Opening datepicker and selecting today");

            // Check if manual selection is enabled
            var manualSelect = false;
            try {
                var saved = localStorage.getItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME);
                manualSelect = (saved === "1");
            } catch (e) {}
            log("Import Step 7/9: manualSelectInitialRefTime=" + String(manualSelect));

            var picker = document.querySelector('#ajaxModal span#firstInitialSegmentReferencePicker, .modal span#firstInitialSegmentReferencePicker, span#firstInitialSegmentReferencePicker');
            var addon = null;
            if (picker) {
                addon = picker.querySelector('span.input-group-addon');
            }
            var openedCal = false;
            if (addon) {
                addon.click();
                log("Import Step 7/9: calendar addon clicked");
                await sleep(400);
                openedCal = true;
            } else {
                log("Import Step 7/9: calendar addon not found");
            }

            var dateInput = document.querySelector('#ajaxModal input#firstInitialSegmentReference, .modal input#firstInitialSegmentReference, input#firstInitialSegmentReference');
            var initialDateVal = dateInput ? (dateInput.value + "") : "";

            if (manualSelect && openedCal) {
                // Wait for user to manually select a date - show confirmation popup
                log("Import Step 7/9: Waiting for user to manually select date (manual mode enabled)");

                var confirmPopup = null;
                var dateConfirmed = false;

                // Create confirmation popup
                var confirmContent = document.createElement("div");
                confirmContent.style.display = "flex";
                confirmContent.style.flexDirection = "column";
                confirmContent.style.gap = "12px";

                var confirmText = document.createElement("div");
                confirmText.style.color = "#fff";
                confirmText.style.fontSize = "14px";
                confirmText.style.lineHeight = "1.5";
                confirmText.textContent = "Please select a date from the calendar, then click 'Confirm Date Selection' to continue.";
                confirmContent.appendChild(confirmText);

                var confirmBtn = document.createElement("button");
                confirmBtn.textContent = "Confirm Date Selection";
                confirmBtn.style.padding = "10px";
                confirmBtn.style.cursor = "pointer";
                confirmBtn.style.background = "#2d7";
                confirmBtn.style.color = "#000";
                confirmBtn.style.border = "none";
                confirmBtn.style.borderRadius = "6px";
                confirmBtn.style.fontSize = "14px";
                confirmBtn.style.fontWeight = "500";
                confirmBtn.style.transition = "background 0.2s";
                confirmBtn.style.width = "100%";

                confirmBtn.addEventListener("mouseenter", function() {
                    confirmBtn.style.background = "#3e8";
                });
                confirmBtn.addEventListener("mouseleave", function() {
                    confirmBtn.style.background = "#2d7";
                });

                confirmBtn.addEventListener("click", function() {
                    dateConfirmed = true;
                    if (confirmPopup) {
                        confirmPopup.close();
                    }
                    log("Import Step 7/9: Date selection confirmed; resuming automation");
                });

                confirmContent.appendChild(confirmBtn);

                confirmPopup = createPopup({
                    title: "Confirm Date Selection",
                    content: confirmContent,
                    width: "400px",
                    height: "auto"
                });

                // Wait for confirmation
                while (!dateConfirmed) {
                    if (isPaused()) {
                        log("Paused; exiting during manual datepicker wait");
                        if (confirmPopup) {
                            confirmPopup.close();
                        }
                        return;
                    }
                    await sleep(300);
                }

                await sleep(300);
            } else if (!manualSelect && openedCal) {
                // Auto-select today (original behavior)
                var todayClicked = false;
                var waitedD = 0;
                var maxWaitD = 8000;
                var todayCell = null;
                while (waitedD < maxWaitD) {
                    if (isPaused()) {
                        log("Paused; exiting during datepicker wait");
                        return;
                    }
                    var daysPanel = document.querySelector('div.datepicker-days');
                    var hasPanel = !!daysPanel;
                    log("Import Step 7/9: datepicker-days present=" + String(hasPanel) + " at t=" + String(waitedD) + "ms");
                    if (daysPanel) {
                        todayCell = daysPanel.querySelector('td.day.active.today');
                        if (!todayCell) {
                            todayCell = daysPanel.querySelector('td.day.today.active');
                        }
                        if (todayCell) {
                            break;
                        }
                    }
                    await sleep(300);
                    waitedD = waitedD + 300;
                }
                var hasToday = !!todayCell;
                log("Import Step 7/9: today cell exists=" + String(hasToday));
                if (todayCell) {
                    todayCell.click();
                    log("Import Step 7/9: today clicked");
                    await sleep(300);
                    todayClicked = true;
                }
            }

            var dateInputFinal = document.querySelector('#ajaxModal input#firstInitialSegmentReference, .modal input#firstInitialSegmentReference, input#firstInitialSegmentReference');
            var dateVal = dateInputFinal ? (dateInputFinal.value + "") : "";
            log("Import Step 7/9: date input exists=" + String(!!dateInputFinal) + " value='" + String(dateVal) + "'");
            log("Import Step 8/9: Checking for 'no assignments' message");
            var noneMsg = hasNoAssignmentsInImportModal(modalImp);
            log("Import Step 8/9: noneMsg=" + String(noneMsg));
            if (noneMsg) {
                log("Import Step 8/9: closing Import modal and advancing to next epoch");
                var cancelBtn0 = getModalCancelButton(modalImp);
                if (cancelBtn0) {
                    cancelBtn0.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }

                var idxAdv0 = getNonScrnEpochIndex();
                var nextIdx = idxAdv0 + 1;
                var messageContent = document.createElement("div");
                messageContent.style.display = "flex";
                messageContent.style.flexDirection = "column";
                messageContent.style.gap = "12px";

                var messageText = document.createElement("div");
                messageText.style.color = "#fff";
                messageText.style.fontSize = "14px";
                messageText.style.lineHeight = "1.5";
                messageText.innerHTML = "No assignments found for the selected criteria.";
                messageContent.appendChild(messageText);

                var okBtn = document.createElement("button");
                okBtn.textContent = "OK";
                okBtn.style.padding = "10px";
                okBtn.style.cursor = "pointer";
                okBtn.style.background = "#2d7";
                okBtn.style.color = "#000";
                okBtn.style.border = "none";
                okBtn.style.borderRadius = "6px";
                okBtn.style.fontSize = "14px";
                okBtn.style.fontWeight = "500";
                okBtn.style.transition = "background 0.2s";
                okBtn.style.width = "100%";

                okBtn.addEventListener("mouseenter", function() {
                    okBtn.style.background = "#3e8";
                });
                okBtn.addEventListener("mouseleave", function() {
                    okBtn.style.background = "#2d7";
                });

                var noAssignmentsPopup = null;
                okBtn.addEventListener("click", function() {
                    if (noAssignmentsPopup) {
                        noAssignmentsPopup.close();
                    }
                    if (cohortId && cohortId.length > 0) {
                        setImportDone(cohortId);
                        log("Import Step 8/9: marked import done for cohortId=" + String(cohortId) + " (none assignments)");
                    }
                    setNonScrnEpochIndex(nextIdx);
                    clearImportSubjectIds();
                    log("Import Step 8/9: advanced epoch idx=" + String(nextIdx) + "; navigating to Study Show");
                    clearRunMode();
                    try {
                        localStorage.removeItem(STORAGE_IMPORT_DONE_MAP);
                        localStorage.removeItem(STORAGE_IMPORT_SUBJECT_IDS);
                        localStorage.removeItem(STORAGE_NON_SCRN_EPOCH_INDEX);
                        localStorage.removeItem(STORAGE_NON_SCRN_SELECTED_EPOCH);
                        localStorage.removeItem(STORAGE_IMPORT_COHORT_EDIT_DONE);
                        log("Cleanup: cleared non-SCRN import state");
                    } catch (e) {
                        log("Cleanup error: " + String(e));
                    }
                    location.href = STUDY_SHOW_URL;
                });

                messageContent.appendChild(okBtn);

                noAssignmentsPopup = createPopup({
                    title: "No Assignments Found",
                    content: messageContent,
                    width: "400px",
                    height: "auto"
                });

                return;
            }
            log("Import Step 9/9: Collecting volunteer IDs");
            var tbodyImp = null;
            var waitedT = 0;
            var maxWaitT = 10000;
            while (waitedT < maxWaitT) {
                if (isPaused()) {
                    log("Paused; exiting during volunteer ID wait");
                    return;
                }
                tbodyImp = document.querySelector('#ajaxModal tbody#cohortAssignmentTableBody, .modal tbody#cohortAssignmentTableBody, tbody#cohortAssignmentTableBody');
                var tbodyReady = !!tbodyImp;
                log("Import Step 9/9: cohortAssignmentTableBody present=" + String(tbodyReady) + " at t=" + String(waitedT) + "ms");
                if (tbodyReady) {
                    break;
                }
                await sleep(300);
                waitedT = waitedT + 300;
            }
            var ids = [];
            if (tbodyImp) {
                var scanNodes = tbodyImp.querySelectorAll("tr, td, span, a, div");
                log("Import Step 9/9: scanNodes count=" + String(scanNodes.length));
                var iScan = 0;
                while (iScan < scanNodes.length) {
                    var node = scanNodes[iScan];
                    var text = (node.textContent + "").trim();
                    var idFromText = extractVolunteerIdFromFreeText(text);
                    if (idFromText && idFromText.length > 0) {
                        ids.push(idFromText);
                    }
                    if (node.tagName && node.tagName.toLowerCase() === "a") {
                        var href = node.getAttribute("href") + "";
                        var mV = href.match(/\/secure\/volunteers\/manage\/show\/(\d+)\//);
                        if (mV && mV[1]) {
                            ids.push(mV[1]);
                        }
                    }
                    iScan = iScan + 1;
                }
            } else {
                log("Import Step 9/9: cohortAssignmentTableBody not found");
            }
            var uniq = {};
            var dedup = [];
            var iD = 0;
            while (iD < ids.length) {
                var vId = String(ids[iD]);
                if (vId.length > 0) {
                    if (!uniq[vId]) {
                        uniq[vId] = 1;
                        dedup.push(vId);
                    }
                }
                iD = iD + 1;
            }
            setImportSubjectIds(dedup);
            log("Import Step 9/9: collected volunteerIds len=" + String(dedup.length) + " ids=" + JSON.stringify(dedup));
            var saveImpBtn = document.querySelector('#ajaxModal button#actionButton.btn.green, .modal button#actionButton.btn.green, button#actionButton.btn.green');
            if (!saveImpBtn) {
                saveImpBtn = document.querySelector('#ajaxModal button#actionButton, .modal button#actionButton, button#actionButton');
            }
            if (!saveImpBtn) {
                saveImpBtn = document.querySelector('#ajaxModal button.btn.green[type="button"], .modal button.btn.green[type="button"], button.btn.green[type="button"]');
            }
            var saveImpExists = !!saveImpBtn;
            log("Import Step 9/9: save button exists=" + String(saveImpExists));
            if (!saveImpBtn) {
                log("processCohortShowPageImportNonScrn: Import modal Save button not found");
                return;
            }
            if (cohortId && cohortId.length > 0) {
                setImportDone(cohortId);
                log("Import Guard: pre-marked import done for cohortId=" + String(cohortId));
            }
            saveImpBtn.click();
            log("Import modal Save clicked");
            await sleep(1200);
            log("Import modal post-save wait done");
            if (cohortId && cohortId.length > 0) {
                setImportDone(cohortId);
                log("Import Guard: marked import done for cohortId=" + String(cohortId));
            }
            log("Import completed: page will refresh, activation will continue on next load");
            return;
        } else {
            log("Import Guard: skipping Import modal for cohortId=" + String(cohortId));
        }

        var volunteerIds = getImportSubjectIds();
        if (volunteerIds.length === 0) {
            log("Activation: no stored volunteer IDs found; exiting");
            clearRunMode();
            try {
                localStorage.removeItem(STORAGE_IMPORT_DONE_MAP);
                localStorage.removeItem(STORAGE_IMPORT_SUBJECT_IDS);
                localStorage.removeItem(STORAGE_NON_SCRN_EPOCH_INDEX);
                localStorage.removeItem(STORAGE_NON_SCRN_SELECTED_EPOCH);
                localStorage.removeItem(STORAGE_IMPORT_COHORT_EDIT_DONE);
                log("Cleanup: cleared non-SCRN import state");
            } catch (e) {
                log("Cleanup error: " + String(e));
            }
            return;
        }
        log("Activation: starting for volunteerIds count=" + String(volunteerIds.length));
        var listReady = await waitForListTable(15000);
        if (!listReady) {
            await sleep(800);
        }
        log("Activation: list table ready, starting volunteer processing");

        var i = 0;
        while (i < volunteerIds.length) {
            var volId = volunteerIds[i];
            log("Activation: processing volunteerId=" + String(volId));

            var row = findCohortRowByVolunteerId(volId);
            var waitedRow = 0;
            if (row) {
                log("Activation: row found immediately for volunteerId=" + String(volId));
            } else {
                var maxWaitRow = 5000;
                var checkInterval = 200;
                while (waitedRow < maxWaitRow) {
                    await sleep(checkInterval);
                    waitedRow += checkInterval;
                    row = findCohortRowByVolunteerId(volId);
                    if (row) {
                        log("Activation: row found for volunteerId=" + String(volId) + " after " + String(waitedRow) + "ms");
                        break;
                    }
                    if (waitedRow % 1000 === 0) {
                        log("Activation: still searching for volunteerId=" + String(volId) + " waited=" + String(waitedRow) + "ms");
                    }
                }
            }
            if (!row) {
                log("Activation: row not found for volunteerId=" + String(volId) + " after " + String(waitedRow) + "ms; skipping");
                i++;
                continue;
            }

            var actionBtn = getRowActionButton(row);
            if (!actionBtn) {
                log("Activation: action button not found for volunteerId=" + String(volId));
                i++;
                continue;
            }
            actionBtn.click();
            await sleep(300);

            var planLink = getMenuLinkActivatePlan(row);
            if (!planLink) {
                log("Activation: Activate Plan link not found for volunteerId=" + String(volId));
                i++;
                continue;
            }
            planLink.click();
            var okPlan = await clickBootboxOk(5000);
            if (!okPlan) {
                await sleep(500);
            }

            var found = await waitForActivateVolunteerById(volId, 45000);
            if (!found) {
                log("Activation: Activate Volunteer link not found for volunteerId=" + String(volId));
                i++;
                continue;
            }
            found.link.click();
            var okVol = await clickBootboxOk(5000);
            if (!okVol) {
                await sleep(500);
            }

            log("Activation: completed for volunteerId=" + String(volId));
            i++;
        }


        log("Activation Complete: clearing collected IDs and moving to next epoch");
        clearImportSubjectIds();
        var idxNext = getNonScrnEpochIndex();
        idxNext = idxNext + 1;
        setNonScrnEpochIndex(idxNext);
        log("Activation Complete: next epoch idx=" + String(idxNext) + "; navigating to Study Show");
        clearRunMode();
        try {
            localStorage.removeItem(STORAGE_IMPORT_DONE_MAP);
            localStorage.removeItem(STORAGE_IMPORT_SUBJECT_IDS);
            localStorage.removeItem(STORAGE_NON_SCRN_EPOCH_INDEX);
            localStorage.removeItem(STORAGE_NON_SCRN_SELECTED_EPOCH);
            localStorage.removeItem(STORAGE_IMPORT_COHORT_EDIT_DONE);
            log("Cleanup: cleared non-SCRN import state");
        } catch (e) {
            log("Cleanup error: " + String(e));
        }
        location.href = STUDY_SHOW_URL;
    }

    async function processStudyShowPageForNonScrn() {
        var autoNonScrn = getQueryParam("autononscrn");
        var mode = getRunMode();
        if (!(mode === "nonscrn" && autoNonScrn === "1")) {
            return;
        }

        var tbody = await waitForSelector('tbody#epochTableBody', 5000);
        if (!tbody) {
            log("Epoch table not found");
            return;
        }

        var anchors = tbody.querySelectorAll('a[href^="/secure/administration/studies/epoch/show/"]');
        if (anchors.length === 0) {
            log("No epochs found");
            return;
        }

        var contentDiv = document.createElement("div");
        contentDiv.style.display = "flex";
        contentDiv.style.flexDirection = "column";
        contentDiv.style.gap = "8px";

        var popup = null;

        anchors.forEach(function (a) {
            var btn = document.createElement("button");
            btn.textContent = a.textContent.trim();
            btn.style.display = "block";
            btn.style.width = "100%";
            btn.style.padding = "10px";
            btn.style.cursor = "pointer";
            btn.style.background = "#4a90e2";
            btn.style.color = "#fff";
            btn.style.border = "none";
            btn.style.borderRadius = "6px";
            btn.style.fontSize = "14px";
            btn.style.fontWeight = "500";
            btn.style.transition = "background 0.2s";

            btn.addEventListener("mouseenter", function() {
                btn.style.background = "#357abd";
            });
            btn.addEventListener("mouseleave", function() {
                btn.style.background = "#4a90e2";
            });

            btn.addEventListener("click", async function () {
                var label = a.textContent.trim();
                var href = a.getAttribute("href");
                localStorage.setItem(STORAGE_NON_SCRN_SELECTED_EPOCH, href);

                // Save checkbox preference
                var checkbox = document.getElementById("manualSelectInitialRefTimeImport");
                if (checkbox) {
                    try {
                        localStorage.setItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME, checkbox.checked ? "1" : "0");
                    } catch (e) {}
                }

                if (popup) {
                    popup.close();
                }
                clearNonScrnEpochIndex();

                if (isScreeningLabel(label)) {
                    log("Selected Screening epoch; starting Cohort Add automation");
                    setRunMode("epoch");
                    location.href = location.origin + href + "?autocohort=1";
                } else {
                    log("Selected non-SCRN epoch; starting Import automation");
                    setRunMode("epochImport");
                    location.href = location.origin + href + "?autoepochimport=1";
                }
            });

            contentDiv.appendChild(btn);
        });

        // Add checkbox for Manual Select Initial Ref Time
        var checkboxRow = document.createElement("div");
        checkboxRow.style.display = "flex";
        checkboxRow.style.alignItems = "center";
        checkboxRow.style.gap = "8px";
        checkboxRow.style.marginTop = "8px";
        checkboxRow.style.padding = "8px";
        checkboxRow.style.borderTop = "1px solid #444";

        var checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.id = "manualSelectInitialRefTimeImport";
        checkbox.style.cursor = "pointer";

        // Load saved preference
        try {
            var saved = localStorage.getItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME);
            if (saved === "1") {
                checkbox.checked = true;
            }
        } catch (e) {}

        var label = document.createElement("label");
        label.htmlFor = "manualSelectInitialRefTimeImport";
        label.textContent = "Manual Select Initial Ref Time";
        label.style.cursor = "pointer";
        label.style.color = "#fff";
        label.style.fontSize = "14px";
        label.style.userSelect = "none";

        checkboxRow.appendChild(checkbox);
        checkboxRow.appendChild(label);
        contentDiv.appendChild(checkboxRow);

        popup = createPopup({
            title: "Select Epoch to Import Subject",
            content: contentDiv,
            width: "350px",
            height: "auto",
            maxHeight: "80%"
        });
    }

    //==========================
    // RUN ICF BARCODE (INFORMED CONSENT) FEATURE
    //==========================
    // This section contains all functions related to informed consent barcode collection.
    // This feature automates the process of collecting informed consent barcodes for subjects.
    // Functions handle barcode retrieval from study show pages, subject list scanning,
    // consent tab navigation, and barcode entry into consent forms.
    //==========================
    function InformedConsentFunctions() {}
    // Persist last volunteer id used for activation flows.
    function setLastVolunteerId(id) {
        try {
            localStorage.setItem("activityPlanState.lastVolunteerId", String(id));
            log("Stored lastVolunteerId=" + String(id));
        } catch (e) {}
    }

    // Retrieve persisted last volunteer id.
    function getLastVolunteerId() {
        var raw = null;
        try {
            raw = localStorage.getItem("activityPlanState.lastVolunteerId");
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    async function processSubjectsListPageForConsent() {
        if (isPaused()) {
            log("Paused; skipping subjects list consent processing");
            return;
        }
        var auto = getQueryParam("autoconsent");
        var mode = getRunMode();
        var go = (mode === "consent" || mode === "allconsent") || (auto === "1" && mode && mode.length > 0);
        if (!go) {
            return;
        }
        var codeExisting = getIcBarcode();
        if (!codeExisting || codeExisting.length === 0) {
            log("IC barcode not cached; routing to Study Show to pull barcode");
            location.href = STUDY_SHOW_URL + "?autoconsent=1";
            return;
        }
        var bodyReady = await waitForSubjectsBody(8000);
        if (!bodyReady) {
            log("processSubjectsListPageForConsent: subjectTableBody not ready");
            return;
        }
        var rowsReady = await waitForSubjectRows(8000);
        if (!rowsReady) {
            log("processSubjectsListPageForConsent: subject rows not ready");
            return;
        }
        var rows = document.querySelectorAll('tbody#subjectTableBody > tr');
        log("processSubjectsListPageForConsent: actual subject rows=" + String(rows.length));
        var volId = getLastVolunteerId();
        if (volId && volId.length > 0) {
            var iTarget = 0;
            var targetRow = null;
            while (iTarget < rows.length) {
                var rT = rows[iTarget];
                var volA = rT.querySelector('a[href^="/secure/volunteers/manage/show/"]');
                var match = false;
                if (volA) {
                    var hrefVol = volA.getAttribute('href') + "";
                    var mv = hrefVol.match(/\/show\/(\d+)/);
                    var idVol = "";
                    if (mv && mv[1]) {
                        idVol = mv[1];
                    }
                    match = idVol === String(volId);
                    log("processSubjectsListPageForConsent: row " + String(iTarget) + " volunteer href=" + String(hrefVol) + " parsed=" + String(idVol) + " match=" + String(match));
                } else {
                    log("processSubjectsListPageForConsent: row " + String(iTarget) + " volunteer link missing");
                }
                if (match) {
                    targetRow = rT;
                    break;
                }
                iTarget = iTarget + 1;
            }
            if (targetRow) {
                var noConsentT = rowHasNoInformedConsentInFourthCell(targetRow);
                log("processSubjectsListPageForConsent: targeted row hasNoConsent=" + String(noConsentT));
                if (noConsentT) {
                    var idT = extractSubjectShowIdFromRow(targetRow);
                    if (!idT || idT.length === 0) {
                        log("processSubjectsListPageForConsent: targeted row show id missing");
                    } else {
                        var urlT = location.origin + "/secure/study/subjects/show/" + idT + "?autoconsent=1";
                        log("processSubjectsListPageForConsent: navigating to targeted " + String(urlT));
                        location.href = urlT;
                        return;
                    }
                } else {
                    log("processSubjectsListPageForConsent: targeted volunteer already has consent; scanning for next row without consent");
                }
            } else {
                log("processSubjectsListPageForConsent: targeted volunteer row not found; scanning");
            }
        }
        var idxRaw = getConsentScanIndex();
        var idx = 0;
        if (idxRaw && idxRaw.length > 0) {
            var n = parseInt(idxRaw, 10);
            if (!isNaN(n)) {
                idx = n;
            }
        }
        var i = idx;
        var pickedRow = null;
        while (i < rows.length) {
            var r = rows[i];
            var noC = rowHasNoInformedConsentInFourthCell(r);
            log("processSubjectsListPageForConsent: row " + String(i) + " hasNoConsent=" + String(noC));
            if (noC) {
                pickedRow = r;
                log("processSubjectsListPageForConsent: picked row index=" + String(i));
                break;
            }
            i = i + 1;
        }
        if (!pickedRow) {
            log("processSubjectsListPageForConsent: no row without consent from index=" + String(idx) + "; falling back to last Show link on subjects table");
            var showsReady = await waitForSubjectShowAnchors(6000);
            if (!showsReady) {
                log("processSubjectsListPageForConsent: show anchors not ready");
                clearConsentScanIndex();
                return;
            }
            var allShows = document.querySelectorAll('tbody#subjectTableBody a[href^="/secure/study/subjects/show/"]');
            if (!allShows || allShows.length === 0) {
                log("processSubjectsListPageForConsent: no show anchors found");
                clearConsentScanIndex();
                return;
            }
            var lastHref = allShows[allShows.length - 1].getAttribute('href') + "";
            var ml = lastHref.match(/\/show\/(\d+)/);
            if (!ml || !ml[1]) {
                log("processSubjectsListPageForConsent: last show href parse failed href=" + String(lastHref));
                clearConsentScanIndex();
                return;
            }
            var lastId = ml[1];
            var urlL = location.origin + "/secure/study/subjects/show/" + lastId + "?autoconsent=1";
            log("processSubjectsListPageForConsent: navigating to fallback " + String(urlL));
            location.href = urlL;
            return;
        }
        setConsentScanIndex(String(i));
        var pickedId = extractSubjectShowIdFromRow(pickedRow);
        if (!pickedId || pickedId.length === 0) {
            log("processSubjectsListPageForConsent: picked row show id missing; advancing idx and retry");
            i = i + 1;
            setConsentScanIndex(String(i));
            location.href = "/secure/study/subjects/list?autoconsent=1";
            return;
        }
        var targetUrl = location.origin + "/secure/study/subjects/show/" + pickedId + "?autoconsent=1";
        log("processSubjectsListPageForConsent: navigating to " + String(targetUrl));
        location.href = targetUrl;
    }
    // Append an id to the persisted selected volunteers list.
    function appendSelectedVolunteerId(id) {
        var arr = getSelectedVolunteerIds();
        arr.push(String(id));
        setSelectedVolunteerIds(arr);
    }

    // Return most recently appended selected volunteer id.
    function getLastSelectedVolunteerId() {
        var arr = getSelectedVolunteerIds();
        if (arr.length === 0) {
            return "";
        }
        return String(arr[arr.length - 1]);
    }

    // Clear persisted informed-consent barcode.
    function clearIcBarcode() {
        try {
            localStorage.removeItem(STORAGE_IC_BARCODE);
            log("IC barcode cleared");
        } catch (e) {}
    }
    // Persist array of selected volunteer ids.
    function setSelectedVolunteerIds(ids) {
        try {
            localStorage.setItem(STORAGE_SELECTED_IDS, JSON.stringify(ids));
            log("SelectedIds set len=" + String(ids.length));
        } catch (e) {}
    }

    // Read persisted selected volunteer ids array.
    function getSelectedVolunteerIds() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_SELECTED_IDS);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        var arr = [];
        try {
            arr = JSON.parse(raw);
        } catch (e2) {}
        if (Array.isArray(arr)) {
            return arr;
        }
        return [];
    }

    // Read persisted consent scan index.
    function getConsentScanIndex() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_CONSENT_SCAN_INDEX);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    // Persist consent scan index used when scanning subject list.
    function setConsentScanIndex(idx) {
        try {
            localStorage.setItem(STORAGE_CONSENT_SCAN_INDEX, String(idx));
        } catch (e) {}
    }
    // Automate consent collection on a subject show page when required.
    async function processSubjectShowPageForConsent() {
        if (isPaused()) {
            log("Paused; skipping subject show consent processing");
            return;
        }
        var auto = getQueryParam("autoconsent");
        var mode = getRunMode();
        var go = (mode === "consent" || mode === "allconsent") || (auto === "1" && mode && mode.length > 0);
        if (!go) {
            return;
        }
        var codeExisting = getIcBarcode();
        if (!codeExisting || codeExisting.length === 0) {
            log("IC barcode missing on subject page; routing to Study Show to pull barcode");
            location.href = STUDY_SHOW_URL + "?autoconsent=1";
            return;
        }
        var filled = subjectShowHasConsentFilled();
        if (filled) {
            var idxRaw = getConsentScanIndex();
            if (idxRaw && idxRaw.length > 0) {
                var n = parseInt(idxRaw, 10);
                if (isNaN(n)) {
                    n = 0;
                }
                n = n + 1;
                setConsentScanIndex(String(n));
                location.href = "/secure/study/subjects/list?autoconsent=1";
                log("Consent already filled; advancing scan index=" + String(n));
                return;
            }
            log("Consent already filled for targeted subject");
            return;
        }
        var ok = await collectConsentOnSubjectShow();
        if (!ok) {
            log("Consent collection failed");
            return;
        }
        clearConsentScanIndex();
        var wasAllMode = getRunMode() === "consent" && localStorage.getItem(STORAGE_RUN_ALL_POPUP) === "1";
        clearRunMode();
        // If this was the final step of Run All, close the popup
        if (wasAllMode) {
            try {
                localStorage.removeItem(STORAGE_RUN_ALL_POPUP);
                if (RUN_ALL_POPUP_REF && RUN_ALL_POPUP_REF.close) {
                    RUN_ALL_POPUP_REF.close();
                }
                RUN_ALL_POPUP_REF = null;
            } catch (e) {}
        }
        log("Consent collected and run state cleared");
    }
    // Detect whether consent is present on a subject show page.
    function subjectShowHasConsentFilled() {
        var blocks = document.querySelectorAll('tr.borderTop');
        var i = 0;
        while (i < blocks.length) {
            var tr = blocks[i];
            var label = tr.querySelector('td.subjectSortaBold');
            if (label) {
                var text = (label.textContent + "").trim().toLowerCase();
                var match = text.indexOf('informed consent:') !== -1;
                if (match) {
                    var valTd = tr.querySelectorAll('td')[1];
                    if (valTd) {
                        var hasSpan = !!valTd.querySelector('span.tooltips');
                        if (hasSpan) {
                            return true;
                        }
                        return false;
                    }
                }
            }
            i = i + 1;
        }
        return false;
    }


    // Automate the consent collection flow on a subject show page.
    async function collectConsentOnSubjectShow() {
        var tab = await waitForSelector('a#consentTabLink[href="#consentTab"]', 8000);
        if (!tab) {
            log("Consent tab link not found");
            return false;
        }
        tab.click();
        await sleep(400);
        var collectBtn = await waitForSelector('button[id^="consentCollectButton_"]', 8000);
        if (!collectBtn) {
            log("Consent collect button not found");
            return false;
        }
        collectBtn.click();
        await sleep(600);
        var barcodeIcon = await waitForSelector('i#icfBarcodeIcon.fa.fa-barcode', 2000);
        if (!barcodeIcon) {
            var addNewBtn = document.querySelector('a.btn.btn-default.blue.pull-right[onclick*="maybeAddNew"]');
            if (!addNewBtn) {
                addNewBtn = await waitForSelector('a.btn.btn-default.blue.pull-right[onclick*="maybeAddNew"]', 2000);
            }
            if (addNewBtn) {
                addNewBtn.click();
                await sleep(300);
                var okAdd = await clickBootboxOk(6000);
                if (!okAdd) {
                    await sleep(300);
                }
                barcodeIcon = await waitForSelector('i#icfBarcodeIcon.fa.fa-barcode', 8000);
            }
        }
        if (!barcodeIcon) {
            log("Barcode icon not found");
            return false;
        }
        barcodeIcon.click();
        await sleep(400);
        var inputBox = await waitForSelector('input.bootbox-input.bootbox-input-text.form-control', 8000);
        if (!inputBox) {
            log("Bootbox input not found");
            return false;
        }
        var code = getIcBarcode();
        if (!code || code.length === 0) {
            log("IC barcode missing");
            return false;
        }
        inputBox.value = code;
        var evt = new Event("input", { bubbles: true });
        inputBox.dispatchEvent(evt);
        var ok1 = await clickBootboxOk(6000);
        if (!ok1) {
            log("Bootbox OK for barcode not found");
            return false;
        }
        var nowBtn = await waitForSelector('button[id$="_currentTime"].btn.btn-xs.dark', 8000);
        if (!nowBtn) {
            log("Current Time button not found");
            return false;
        }
        nowBtn.click();
        await sleep(300);
        var saveBtn = await waitForSelector('button#saveAndCloseSubmit.btn.green', 8000);
        if (!saveBtn) {
            log("Save and Close button not found");
            return false;
        }
        saveBtn.click();
        await sleep(1200);
        return true;
    }

    // Retrieve persisted informed-consent barcode.
    function getIcBarcode() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_IC_BARCODE);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    // Return the subjects-list tbody element or fallback table tbody.
    function getSubjectsListTbody() {
        var tbody = document.querySelector('tbody#subjectTableBody');
        if (!tbody) {
            var table = document.querySelector('table.table.table-striped.table-hover.table-bordered');
            if (!table) {
                return null;
            }
            tbody = table.querySelector('tbody');
        }
        return tbody;
    }

    // Heuristic to determine whether the fourth cell lacks informed consent info.
    function rowHasNoInformedConsentInFourthCell(row) {
        if (!row) {
            log("rowHasNoInformedConsentInFourthCell: row is null");
            return false;
        }
        var tds = row.querySelectorAll('td');
        var tdCount = tds.length;
        log("rowHasNoInformedConsentInFourthCell: tdCount=" + String(tdCount));
        if (tdCount < 4) {
            log("rowHasNoInformedConsentInFourthCell: less than 4 tds");
            return false;
        }
        var fourth = tds[3];
        var inner = fourth.querySelector('table.clinSparkInnerTable');
        if (!inner) {
            log("rowHasNoInformedConsentInFourthCell: inner table missing");
            return false;
        }
        var body = inner.querySelector('tbody');
        if (!body) {
            log("rowHasNoInformedConsentInFourthCell: inner tbody missing");
            return false;
        }
        var rows = body.querySelectorAll('tr');
        log("rowHasNoInformedConsentInFourthCell: inner rows=" + String(rows.length));
        var i = 0;
        while (i < rows.length) {
            var tr = rows[i];
            var cells = tr.querySelectorAll('td');
            if (cells.length >= 2) {
                var label = (cells[0].textContent + "").trim().toLowerCase();
                var match = label.indexOf("informed consent:") !== -1;
                if (match) {
                    var valTd = cells[1];
                    var hasSpan = !!valTd.querySelector('span.tooltips');
                    var txt = (valTd.textContent + "").trim();
                    log("rowHasNoInformedConsentInFourthCell: found label; valueText='" + String(txt) + "' hasSpan=" + String(hasSpan));
                    if (txt.length === 0 && !hasSpan) {
                        log("rowHasNoInformedConsentInFourthCell: determined EMPTY");
                        return true;
                    }
                    log("rowHasNoInformedConsentInFourthCell: determined NOT EMPTY");
                    return false;
                }
            }
            i = i + 1;
        }
        log("rowHasNoInformedConsentInFourthCell: consent label not found");
        return false;
    }

    // Wait for subjects table body to appear.
    async function waitForSubjectsBody(timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var tbody = document.querySelector('tbody#subjectTableBody');
            if (tbody) {
                return tbody;
            }
            await sleep(200);
        }
        return null;
    }

    // Wait for subject rows to be available.
    async function waitForSubjectRows(timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var rows = document.querySelectorAll('tbody#subjectTableBody > tr');
            if (rows && rows.length > 0) {
                return rows;
            }
            await sleep(200);
        }
        return null;
    }

    // Wait for subject show anchors to be available.
    async function waitForSubjectShowAnchors(timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var anchors = document.querySelectorAll('tbody#subjectTableBody a[href^="/secure/study/subjects/show/"]');
            if (anchors && anchors.length > 0) {
                return anchors;
            }
            await sleep(200);
        }
        return null;
    }

    // Extract the subject show id from a row's show anchor.
    function extractSubjectShowIdFromRow(row) {
        if (!row) {
            log("extractSubjectShowIdFromRow: row is null");
            return "";
        }
        var tds = row.querySelectorAll('td');
        log("extractSubjectShowIdFromRow: tdCount=" + String(tds.length));
        var a = null;
        if (tds.length >= 7) {
            var tdShow = tds[6];
            a = tdShow.querySelector('a[href^="/secure/study/subjects/show/"]');
        }
        if (!a) {
            a = row.querySelector('a[href^="/secure/study/subjects/show/"]');
        }
        if (!a) {
            log("extractSubjectShowIdFromRow: no show anchor");
            return "";
        }
        var href = a.getAttribute('href') + "";
        log("extractSubjectShowIdFromRow: href=" + String(href));
        var m = href.match(/\/show\/(\d+)/);
        if (m && m[1]) {
            log("extractSubjectShowIdFromRow: id=" + String(m[1]));
            return m[1];
        }
        log("extractSubjectShowIdFromRow: regex failed");
        return "";
    }

    //==========================
    // RUN ADD COHORT SUBJECTS FEATURE
    //==========================
    // This section contains all functions related to adding cohort subjects.
    // This feature automates the process of editing cohort settings, adding volunteers
    // to cohorts, and activating them. Functions handle epoch selection, cohort editing,
    // volunteer selection via random letter search, and activation workflows.
    //==========================
    function AddCohortSubjectsFunctions() {}

    async function processEpochShowPageForAddCohort() {
        if (isPaused()) {
            log("Paused; exiting processEpochShowPageForAddCohort");
            return;
        }
        var auto = getQueryParam("autoepochaddcohort");
        var go = auto === "1" || getRunMode() === "epochAddCohort";
        if (!go) {
            return;
        }
        var anchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        if (anchors.length === 0) {
            log("Epoch has no cohorts");
            return;
        }
        var target = anchors[0];
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            log("Cohort href missing");
            return;
        }
        location.href = location.origin + href + "?autocohortadd=1";
        log("Routing to first cohort show for add");
    }

    // Check for duplicate volunteer error in help-block
    function hasDuplicateVolunteerError(modal) {
        if (!modal) {
            modal = document.querySelector("#ajaxModal, .modal");
        }
        if (!modal) {
            return false;
        }
        var helpBlocks = modal.querySelectorAll('span.help-block');
        var i = 0;
        while (i < helpBlocks.length) {
            var text = (helpBlocks[i].textContent + "").trim().toLowerCase();
            if (text.indexOf("volunteers must be unique per cohort") !== -1) {
                return true;
            }
            i = i + 1;
        }
        return false;
    }

    async function processCohortShowPageAddCohort() {
        log("processCohortShowPageAddCohort: start");
        if (isPaused()) {
            log("Paused; exiting processCohortShowPageAddCohort");
            return;
        }
        var auto = getQueryParam("autocohortadd");
        var go = (auto === "1" || getRunMode() === "epochAddCohort");
        log("processCohortShowPageAddCohort: auto=" + String(auto) + " runMode=" + String(getRunMode()) + " go=" + String(go));
        if (!go) {
            log("processCohortShowPageAddCohort: not in add cohort mode; exiting");
            return;
        }
        var g = getCohortGuard();
        log("CohortGuard at start=" + String(g));
        var successOnLoad = hasSuccessAlert();
        if (successOnLoad && g !== "postsave") {
            log("Success alert on load; proceeding to activation flow");
            var listReadyX = await waitForListTable(12000);
            if (!listReadyX) {
                await sleep(800);
            }
            var targetVolX = getLastVolunteerId();
            if (!targetVolX) {
                targetVolX = getLastSelectedVolunteerId();
            }
            log("Using id for activation=" + String(targetVolX));
            var targetRowX = null;
            var waitedRowX = 0;
            var maxWaitRowX = 30000;
            while (waitedRowX < maxWaitRowX) {
                targetRowX = findCohortRowByVolunteerId(targetVolX);
                if (targetRowX) {
                    break;
                }
                await sleep(300);
                waitedRowX = waitedRowX + 300;
            }
            if (!targetRowX) {
                log("Target row not found for id=" + String(targetVolX));
                clearRunMode();
                clearCohortGuard();
                return;
            }
            var actionBtnX = getRowActionButton(targetRowX);
            if (!actionBtnX) {
                log("Action button not found in row");
                clearRunMode();
                clearCohortGuard();
                return;
            }
            actionBtnX.click();
            await sleep(300);
            var planLinkX = getMenuLinkActivatePlan(targetRowX);
            if (!planLinkX) {
                log("Activate Plan link not found");
                clearRunMode();
                clearCohortGuard();
                return;
            }
            planLinkX.click();
            var ok1X = await clickBootboxOk(5000);
            if (!ok1X) {
                await sleep(500);
            }
            var foundX = await waitForActivateVolunteerById(targetVolX, 45000);
            if (!foundX) {
                clearRunMode();
                clearCohortGuard();
                return;
            }
            foundX.link.click();
            var ok2X = await clickBootboxOk(5000);
            if (!ok2X) {
                await sleep(500);
            }
            clearRunMode();
            setCohortGuard("done");
            clearSelectedVolunteerIds();
            // Clear editDoneMap when program completely finishes
            try {
                localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
                log("Cleared editDoneMap on program completion");
            } catch (e) {}
            log("Activation flow completed");
            return;
        }
        log("Edit Prep: Opening Actions dropdown on cohort show");
        var openedEdit = await clickActionsDropdownIfNeeded();
        log("Edit Prep: Actions dropdown opened=" + String(!!openedEdit));
        if (!openedEdit) {
            log("Edit Prep: Actions dropdown not found");
            return;
        }
        log("Edit Prep: Locating Edit link");
        var editLink = document.querySelector('a[href^="/secure/administration/manage/studies/cohort/update/"][data-toggle="modal"]');
        var editHref = "";
        var cohortId = "";
        if (editLink) {
            editHref = editLink.getAttribute("href") + "";
            log("Edit Prep: Edit link found href=" + String(editHref));
            cohortId = parseCohortIdFromUpdateHref(editHref);
        } else {
            log("Edit Prep: Cohort Edit link not found");
        }
        var storageKeyAdd = "activityPlanState.cohortAdd.editDoneMap";
        var addMapRaw = null;
        try {
            addMapRaw = localStorage.getItem(storageKeyAdd);
        } catch (eMapRead) {}
        var addMap = {};
        if (addMapRaw) {
            try {
                addMap = JSON.parse(addMapRaw);
            } catch (eMapParse) {}
        }
        if (typeof addMap !== "object" || addMap === null) {
            addMap = {};
        }
        var alreadyEditedAdd = false;
        if (cohortId && cohortId.length > 0) {
            alreadyEditedAdd = !!addMap[String(cohortId)];
        }
        log("Edit Prep Guard: alreadyEditedAdd=" + String(alreadyEditedAdd) + " cohortId=" + String(cohortId));
        if (editLink && !alreadyEditedAdd) {
            log("Edit Prep: Clicking Edit link");
            editLink.click();
            var editModal = await waitForSelector("#ajaxModal, .modal", 6000);
            log("Edit Prep: Edit modal present=" + String(!!editModal));
            if (!editModal) {
                log("Edit Prep: Edit modal did not open");
            } else {
                log("Edit Prep: Checking form-group structure");
                var groups = editModal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
                var groupCount = groups ? groups.length : 0;
                log("Edit Prep: form-group count=" + String(groupCount));
                if (!groups || groups.length === 0) {
                    var waited = 0;
                    var maxWait = 6000;
                    var step = 300;
                    while (waited < maxWait) {
                        if (isPaused()) {
                            log("Paused; exiting inside form-group wait");
                            return;
                        }
                        groups = editModal.querySelectorAll("#modalbody .form-group, .modal-body .form-group, form .form-group");
                        groupCount = groups ? groups.length : 0;
                        if (groupCount > 0) {
                            break;
                        }
                        await sleep(step);
                        waited = waited + step;
                    }
                    log("Edit Prep: post-wait form-group count=" + String(groupCount));
                }
                log("Edit Prep: Applying Subject/Volunteer Source checkboxes");
                var ok6 = true;
                ok6 = ok6 && setCheckboxStateById("subjectInitiation", true);
                ok6 = ok6 && setCheckboxStateById("sourceVolunteerDatabase", true);
                ok6 = ok6 && setCheckboxStateById("sourceAppointments", true);
                ok6 = ok6 && setCheckboxStateById("sourceAppointmentsCohort", true);
                ok6 = ok6 && setCheckboxStateById("sourceScreeningCohorts", true);
                ok6 = ok6 && setCheckboxStateById("sourceLeadInCohorts", true);
                ok6 = ok6 && setCheckboxStateById("sourceRandomizationCohorts", true);
                var ok7 = true;
                ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInCohorts", true);
                ok7 = ok7 && setCheckboxStateById("allowSubjectsActiveInStudies", true);
                var ok8 = true;
                ok8 = ok8 && setCheckboxStateById("requireVolunteerRecruitment", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentEligible", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentIdentified", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentIneligible", false);
                ok8 = ok8 && setCheckboxStateById("allowRecruitmentRemoved", false);
                var ok9 = true;
                ok9 = ok9 && setCheckboxStateById("allowEligibilityEligible", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityPending", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityIneligible", true);
                ok9 = ok9 && setCheckboxStateById("allowEligibilityUnspecified", true);
                var ok10 = true;
                ok10 = ok10 && setCheckboxStateById("allowStatusActive", true);
                ok10 = ok10 && setCheckboxStateById("allowStatusComplete", false);
                ok10 = ok10 && setCheckboxStateById("allowStatusTerminated", false);
                ok10 = ok10 && setCheckboxStateById("allowStatusWithdrawn", false);
                var ok11 = true;
                ok11 = ok11 && setCheckboxStateById("requireInformedConsent", false);
                ok11 = ok11 && setCheckboxStateById("requireOverVolunteeringCheck", false);
                var reason = editModal.querySelector('textarea#reasonForChange');
                if (reason) {
                    reason.value = "Test";
                    var evtR = new Event("input", { bubbles: true });
                    reason.dispatchEvent(evtR);
                }
                log("Edit Prep: Saving Edit modal");
                var saveBtnEdit = await waitForSelector("#actionButton", 5000);
                if (saveBtnEdit) {
                    if (cohortId && cohortId.length > 0) {
                        addMap[String(cohortId)] = 1;
                        try {
                            localStorage.setItem(storageKeyAdd, JSON.stringify(addMap));
                        } catch (eMapStore) {}
                    }
                    saveBtnEdit.click();
                    log("Edit Prep: Edit modal Save clicked");
                    await sleep(2000);
                    log("Edit Prep: Edit modal post-save wait done");
                } else {
                    log("Edit Prep: Edit modal Save button not found");
                }
            }
        } else {
            log("Edit Prep Guard: skipping Edit modal for cohortId=" + String(cohortId));
        }
        setCohortGuard("inprogress");
        clearLastVolunteerId();
        clearSelectedVolunteerIds();
        clearUsedLetters();
        var letters = "abcdefghijklmnopqrstuvwxyz".split("");
        var usedLetters = getUsedLetters();
        var savedOk = false;
        var currentVolunteerId = "";
        var attempts = 0;
        var maxAttempts = 26;
        while (attempts < maxAttempts && usedLetters.length < letters.length) {
            var opened = await clickActionsDropdownIfNeeded();
            if (!opened) {
                return;
            }
            var addLink = await waitForSelector('a#addCohortAssignmentButton[data-toggle="modal"]', 3000);
            if (!addLink) {
                addLink = document.querySelector('a[href^="/secure/study/cohortassign/manage/save/"][data-toggle="modal"]');
            }
            if (!addLink) {
                log("Add Cohort Assignment link not found");
                return;
            }
            addLink.click();
            log("Add Cohort Assignment clicked");
            var modal = await waitForSelector("#ajaxModal, .modal", 5000);
            if (!modal) {
                log("Modal not found");
                return;
            }
            var planSel = await waitForSelector('select#activityPlan', 5000);
            if (!planSel) {
                log("ActivityPlan select not found");
                return;
            }
            var opts = planSel.querySelectorAll("option");
            var chosen = null;
            var j = 0;
            while (j < opts.length) {
                var op = opts[j];
                var val = op.value + "";
                if (val && val.length > 0) {
                    chosen = op;
                    break;
                }
                j = j + 1;
            }
            if (chosen) {
                planSel.value = chosen.value;
                var evt1 = new Event("change", { bubbles: true });
                planSel.dispatchEvent(evt1);
                log("ActivityPlan chosen value=" + String(chosen.value));
            }
            var searchSel = await waitForSelector('select#cohortAssignmentSearch', 5000);
            if (!searchSel) {
                log("cohortAssignmentSearch not found");
                return;
            }
            var needAll = searchSel.value !== "AllVolunteers";
            if (needAll) {
                searchSel.value = "AllVolunteers";
                var evt2 = new Event("change", { bubbles: true });
                searchSel.dispatchEvent(evt2);
                log("Search set to AllVolunteers");
            }
            log("Add Step: Checking for Initial Reference Time datepicker");
            var initialSegmentRefDiv = modal.querySelector('div#initialSegmentReferenceDiv');
            var hasInitialRef = !!initialSegmentRefDiv;
            if (hasInitialRef) {
                // Check if the element is actually visible (not display: none)
                var style = window.getComputedStyle(initialSegmentRefDiv);
                var isVisible = style.display !== "none" && style.visibility !== "hidden";
                if (!isVisible) {
                    hasInitialRef = false;
                    log("Add Step: initialSegmentReferenceDiv exists but is hidden (display: none); skipping datepicker");
                } else {
                    log("Add Step: initialSegmentReferenceDiv present and visible=" + String(hasInitialRef));
                }
            } else {
                log("Add Step: initialSegmentReferenceDiv not present; skipping datepicker");
            }
            if (hasInitialRef) {
                // Check if manual selection is enabled (only relevant if datepicker exists)
                var manualSelect = false;
                try {
                    var saved = localStorage.getItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME);
                    manualSelect = (saved === "1");
                } catch (e) {}
                log("Add Step: manualSelectInitialRefTime=" + String(manualSelect));

                var picker = modal.querySelector('span#initialSegmentReferencePicker');
                var addon = null;
                if (picker) {
                    addon = picker.querySelector('span.input-group-addon');
                }
                var openedCal = false;
                if (addon) {
                    addon.click();
                    log("Add Step: calendar addon clicked");
                    await sleep(400);
                    openedCal = true;
                } else {
                    log("Add Step: calendar addon not found");
                }

                var dateInput = modal.querySelector('input#initialSegmentReference');
                var initialDateVal = dateInput ? (dateInput.value + "") : "";

                if (manualSelect && openedCal) {
                    // Wait for user to manually select a date - show confirmation popup
                    log("Add Step: Waiting for user to manually select date (manual mode enabled)");

                    var confirmPopup = null;
                    var dateConfirmed = false;

                    // Create confirmation popup
                    var confirmContent = document.createElement("div");
                    confirmContent.style.display = "flex";
                    confirmContent.style.flexDirection = "column";
                    confirmContent.style.gap = "12px";

                    var confirmText = document.createElement("div");
                    confirmText.style.color = "#fff";
                    confirmText.style.fontSize = "14px";
                    confirmText.style.lineHeight = "1.5";
                    confirmText.textContent = "Please select a date from the calendar, then click 'Confirm Date Selection' to continue.";
                    confirmContent.appendChild(confirmText);

                    var confirmBtn = document.createElement("button");
                    confirmBtn.textContent = "Confirm Date Selection";
                    confirmBtn.style.padding = "10px";
                    confirmBtn.style.cursor = "pointer";
                    confirmBtn.style.background = "#2d7";
                    confirmBtn.style.color = "#000";
                    confirmBtn.style.border = "none";
                    confirmBtn.style.borderRadius = "6px";
                    confirmBtn.style.fontSize = "14px";
                    confirmBtn.style.fontWeight = "500";
                    confirmBtn.style.transition = "background 0.2s";
                    confirmBtn.style.width = "100%";

                    confirmBtn.addEventListener("mouseenter", function() {
                        confirmBtn.style.background = "#3e8";
                    });
                    confirmBtn.addEventListener("mouseleave", function() {
                        confirmBtn.style.background = "#2d7";
                    });

                    confirmBtn.addEventListener("click", function() {
                        dateConfirmed = true;
                        if (confirmPopup) {
                            confirmPopup.close();
                        }
                        log("Add Step: Date selection confirmed; resuming automation");
                    });

                    confirmContent.appendChild(confirmBtn);

                    confirmPopup = createPopup({
                        title: "Confirm Date Selection",
                        content: confirmContent,
                        width: "400px",
                        height: "auto"
                    });

                    // Wait for confirmation
                    while (!dateConfirmed) {
                        if (isPaused()) {
                            log("Paused; exiting during manual datepicker wait");
                            if (confirmPopup) {
                                confirmPopup.close();
                            }
                            return;
                        }
                        await sleep(300);
                    }

                    await sleep(300);
                } else if (!manualSelect) {
                    // Auto-select today (original behavior)
                    var todayClicked = false;
                    if (openedCal) {
                        var waitedD = 0;
                        var maxWaitD = 8000;
                        var todayCell = null;
                        while (waitedD < maxWaitD) {
                            if (isPaused()) {
                                log("Paused; exiting during datepicker wait");
                                return;
                            }
                            var daysPanel = document.querySelector('div.datepicker-days');
                            var hasPanel = !!daysPanel;
                            log("Add Step: datepicker-days present=" + String(hasPanel) + " at t=" + String(waitedD) + "ms");
                            if (daysPanel) {
                                todayCell = daysPanel.querySelector('td.day.active.today');
                                if (!todayCell) {
                                    todayCell = daysPanel.querySelector('td.day.today.active');
                                }
                                if (todayCell) {
                                    break;
                                }
                            }
                            await sleep(300);
                            waitedD = waitedD + 300;
                        }
                        var hasToday = !!todayCell;
                        log("Add Step: today cell exists=" + String(hasToday));
                        if (todayCell) {
                            todayCell.click();
                            log("Add Step: today clicked");
                            await sleep(300);
                            todayClicked = true;
                        }
                    }
                }

                var dateInputFinal = modal.querySelector('input#initialSegmentReference');
                var dateVal = dateInputFinal ? (dateInputFinal.value + "") : "";
                log("Add Step: date input exists=" + String(!!dateInputFinal) + " value='" + String(dateVal) + "'");
            } else {
                log("Add Step: Initial Reference Time datepicker not present; skipping (ignoring datepicker even if manual checkbox is checked)");
            }
            var s2container = modal.querySelector('#s2id_volunteer');
            if (!s2container) {
                s2container = modal.querySelector('.select2-container.form-control.select2');
            }
            if (!s2container) {
                log("Select2 container not found");
                return;
            }
            var s2choice = s2container.querySelector('a.select2-choice');
            if (s2choice) {
                s2choice.click();
                await sleep(150);
            }
            var focusser = s2container.querySelector('input.select2-focusser');
            if (focusser) {
                focusser.focus();
                var kd = new KeyboardEvent("keydown", { bubbles: true, key: "ArrowDown", keyCode: 40 });
                focusser.dispatchEvent(kd);
                focusser.click();
                await sleep(150);
            }
            var s2drop = document.querySelector('#select2-drop.select2-drop-active');
            if (!s2drop) {
                s2drop = s2container.querySelector('.select2-drop');
            }
            if (!s2drop) {
                s2choice = s2container.querySelector('a.select2-choice');
                if (s2choice) {
                    s2choice.click();
                }
                s2drop = await waitForSelector('#select2-drop.select2-drop-active', 2000);
                if (!s2drop) {
                    s2drop = s2container.querySelector('.select2-drop');
                }
            }
            if (!s2drop) {
                log("Select2 drop not found");
                return;
            }
            var s2input = s2drop.querySelector('input.select2-input');
            if (!s2input) {
                s2input = await waitForSelector('#select2-drop.select2-drop-active input.select2-input', 2000);
            }
            if (!s2input) {
                s2input = s2container.querySelector('input.select2-input');
            }
            if (!s2input) {
                log("Select2 input not found");
                return;
            }
            var letter = pickRandomUnusedLetter(letters, usedLetters);
            if (!letter) {
                log("No unused letters remaining; exiting");
                break;
            }
            usedLetters.push(letter);
            setUsedLetters(usedLetters);
            s2input.value = letter;
            var inpEvt = new Event("input", { bubbles: true });
            s2input.dispatchEvent(inpEvt);
            var keyEvt = new KeyboardEvent("keyup", { bubbles: true, key: letter, keyCode: letter.toUpperCase().charCodeAt(0) });
            s2input.dispatchEvent(keyEvt);
            log("Typed random letter=" + String(letter) + " (used: " + JSON.stringify(usedLetters) + ")");
            var selectionConfirmed = false;
            var confirmWait = 0;
            var confirmMax = 12000;
            while (confirmWait < confirmMax) {
                var enterDown = new KeyboardEvent("keydown", { bubbles: true, key: "Enter", keyCode: 13 });
                s2input.dispatchEvent(enterDown);
                var enterUp = new KeyboardEvent("keyup", { bubbles: true, key: "Enter", keyCode: 13 });
                s2input.dispatchEvent(enterUp);
                await sleep(400);
                var containerClass = s2container.getAttribute("class") + "";
                var hasAllow = containerClass.indexOf("select2-allowclear") !== -1;
                var notOpen = containerClass.indexOf("select2-dropdown-open") === -1;
                var chosenEl = null;
                chosenEl = s2container.querySelector('.select2-chosen');
                if (!chosenEl) {
                    chosenEl = s2container.querySelector('[id^="select2-chosen-"]');
                }
                var chosenText = "";
                if (chosenEl) {
                    chosenText = (chosenEl.textContent + "").trim();
                }
                var notSearch = chosenText.trim().toLowerCase() !== "search";
                if (hasAllow && notOpen && notSearch) {
                    selectionConfirmed = true;
                    var volIdParsed = extractVolunteerIdFromChosenText(chosenText);
                    if (volIdParsed) {
                        currentVolunteerId = volIdParsed;
                        appendSelectedVolunteerId(currentVolunteerId);
                        log("Appended selected id=" + String(currentVolunteerId) + "; listLen=" + String(getSelectedVolunteerIds().length));
                    }
                    log("Selection confirmed; chosenText=" + chosenText + "; parsedId=" + String(currentVolunteerId));
                    break;
                }
                confirmWait = confirmWait + 400;
            }
            if (!selectionConfirmed) {
                var cancelBtn0 = getModalCancelButton(modal);
                if (cancelBtn0) {
                    cancelBtn0.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                attempts = attempts + 1;
                log("Selection not confirmed; incrementing attempts to " + String(attempts) + "; will try next unused letter");
                continue;
            }
            var saveBtn = await waitForSelector('button#actionButton.btn.green[type="button"]', 5000);
            if (!saveBtn) {
                saveBtn = document.querySelector('button#actionButton');
            }
            if (!saveBtn) {
                log("Save button not found");
                return;
            }
            saveBtn.click();
            log("Save clicked; current volunteerId=" + String(currentVolunteerId));
            var successImmediate = hasSuccessAlert();
            if (successImmediate) {
                savedOk = true;
                log("Immediate success alert");
                var stillOpen = document.querySelector('#ajaxModal, .modal');
                if (stillOpen) {
                    var sstyle = window.getComputedStyle(stillOpen);
                    var visible = sstyle.display !== "none" && sstyle.visibility !== "hidden";
                    if (visible) {
                        var cancelBtnNow = getModalCancelButton(stillOpen);
                        if (cancelBtnNow) {
                            cancelBtnNow.click();
                            await waitUntilHidden("#ajaxModal, .modal", 5000);
                        }
                    }
                }
                var lastSelected = getLastSelectedVolunteerId();
                if (lastSelected) {
                    setLastVolunteerId(lastSelected);
                }
                setCohortGuard("postsave");
                break;
            }
            await sleep(600);
            var duplicateError = hasDuplicateVolunteerError(modal);
            if (duplicateError) {
                log("Duplicate volunteer error detected; closing modal and trying next letter");
                var cancelBtn = getModalCancelButton(modal);
                if (cancelBtn) {
                    cancelBtn.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                // Letter is already marked as used, so just increment attempts and continue
                attempts = attempts + 1;
                log("Duplicate volunteer error; incrementing attempts to " + String(attempts) + "; will try next unused letter");
                continue;
            }
            var failedImmediate = hasValidationError();
            if (failedImmediate) {
                var ccancelBtn = getModalCancelButton(modal);
                if (ccancelBtn) {
                    ccancelBtn.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                attempts = attempts + 1;
                log("Immediate validation error; incrementing attempts to " + String(attempts) + "; will try next unused letter");
                continue;
            }
            var waitedAlerts = 0;
            var maxAlerts = 12000;
            var successDetected = false;
            var failureDetected = false;
            var duplicateDetected = false;
            while (waitedAlerts < maxAlerts) {
                await sleep(400);
                var sOk = hasSuccessAlert();
                var sFail = hasValidationError();
                var dupErr = hasDuplicateVolunteerError(modal);
                if (sOk) {
                    successDetected = true;
                    break;
                }
                if (dupErr) {
                    duplicateDetected = true;
                    break;
                }
                if (sFail) {
                    failureDetected = true;
                    break;
                }
                waitedAlerts = waitedAlerts + 400;
            }
            if (successDetected) {
                savedOk = true;
                log("Delayed success alert");
                var stillOpen2 = document.querySelector('#ajaxModal, .modal');
                if (stillOpen2) {
                    var style2 = window.getComputedStyle(stillOpen2);
                    var visible2 = style2.display !== "none" && style2.visibility !== "hidden";
                    if (visible2) {
                        var cancelBtn2 = getModalCancelButton(stillOpen2);
                        if (cancelBtn2) {
                            cancelBtn2.click();
                            await waitUntilHidden("#ajaxModal, .modal", 5000);
                        }
                    }
                }
                var lastSelected2 = getLastSelectedVolunteerId();
                if (lastSelected2) {
                    setLastVolunteerId(lastSelected2);
                }
                setCohortGuard("postsave");
                break;
            }
            if (duplicateDetected) {
                log("Duplicate volunteer error detected after wait; closing modal and trying next letter");
                var cancelBtn3 = getModalCancelButton(modal);
                if (cancelBtn3) {
                    cancelBtn3.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                // Letter is already marked as used, so just increment attempts and continue
                attempts = attempts + 1;
                log("Duplicate volunteer error after wait; incrementing attempts to " + String(attempts) + "; will try next unused letter");
                continue;
            }
            if (failureDetected) {
                var ccancelBtn3 = getModalCancelButton(modal);
                if (ccancelBtn3) {
                    ccancelBtn3.click();
                    await waitUntilHidden("#ajaxModal, .modal", 5000);
                }
                attempts = attempts + 1;
                log("Validation error after wait; incrementing attempts to " + String(attempts) + "; will try next unused letter");
                continue;
            }
            var cancelBtn4 = getModalCancelButton(modal);
            if (cancelBtn4) {
                cancelBtn4.click();
                await waitUntilHidden("#ajaxModal, .modal", 5000);
            }
            attempts = attempts + 1;
            log("No alert; closing modal; incrementing attempts to " + String(attempts) + "; will try next unused letter");
        }
        if (!savedOk) {
            log("No volunteer saved; exiting");
            return;
        }
        var listReady = await waitForListTable(15000);
        if (!listReady) {
            await sleep(800);
        }
        var targetVolId = getLastVolunteerId();
        if (!targetVolId) {
            targetVolId = getLastSelectedVolunteerId();
        }
        log("Proceeding to activation for id=" + String(targetVolId));
        var targetRow = null;
        var waitedRow = 0;
        var maxWaitRow = 30000;
        while (waitedRow < maxWaitRow) {
            targetRow = findCohortRowByVolunteerId(targetVolId);
            if (targetRow) {
                break;
            }
            await sleep(300);
            waitedRow = waitedRow + 300;
        }
        if (!targetRow) {
            log("Activation row not found for id=" + String(targetVolId));
            return;
        }
        var actionBtn = getRowActionButton(targetRow);
        if (!actionBtn) {
            log("Row action button not found");
            return;
        }
        actionBtn.click();
        await sleep(300);
        var planLink = getMenuLinkActivatePlan(targetRow);
        if (!planLink) {
            log("Activate Plan link not found");
            return;
        }
        planLink.click();
        var ok1 = await clickBootboxOk(5000);
        if (!ok1) {
            await sleep(500);
        }
        var found = await waitForActivateVolunteerById(targetVolId, 45000);
        if (!found) {
            return;
        }
        found.link.click();
        var ok2 = await clickBootboxOk(5000);
        if (!ok2) {
            await sleep(500);
        }
        clearRunMode();
        setCohortGuard("done");
        clearSelectedVolunteerIds();
        // Clear editDoneMap when program completely finishes
        try {
            localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
            log("Cleared editDoneMap on program completion");
        } catch (e) {}
        log("Activation completed and run state cleared");
    }

    function clearNonScrnEpochIndex() {
        try {
            localStorage.removeItem(STORAGE_NON_SCRN_EPOCH_INDEX);
        } catch (e) {}
    }
    function setNonScrnEpochIndex(n) {
        try {
            localStorage.setItem(STORAGE_NON_SCRN_EPOCH_INDEX, String(n));
        } catch (e) {}
    }
    async function processStudyShowPageForAddCohort() {
        var autoAddCohort = getQueryParam("autoaddcohort");
        var mode = getRunMode();
        if (!(mode === "epochAddCohort" && autoAddCohort === "1")) {
            return;
        }

        var tbody = await waitForSelector('tbody#epochTableBody', 5000);
        if (!tbody) {
            log("Epoch table not found");
            return;
        }

        var anchors = tbody.querySelectorAll('a[href^="/secure/administration/studies/epoch/show/"]');
        if (anchors.length === 0) {
            log("No epochs found");
            return;
        }

        var contentDiv = document.createElement("div");
        contentDiv.style.display = "flex";
        contentDiv.style.flexDirection = "column";
        contentDiv.style.gap = "8px";

        var popup = null;

        anchors.forEach(function (a) {
            var btn = document.createElement("button");
            btn.textContent = a.textContent.trim();
            btn.style.display = "block";
            btn.style.width = "100%";
            btn.style.padding = "10px";
            btn.style.cursor = "pointer";
            btn.style.background = "#4a90e2";
            btn.style.color = "#fff";
            btn.style.border = "none";
            btn.style.borderRadius = "6px";
            btn.style.fontSize = "14px";
            btn.style.fontWeight = "500";
            btn.style.transition = "background 0.2s";

            btn.addEventListener("mouseenter", function() {
                btn.style.background = "#357abd";
            });
            btn.addEventListener("mouseleave", function() {
                btn.style.background = "#4a90e2";
            });

            btn.addEventListener("click", async function () {
                var label = a.textContent.trim();
                var href = a.getAttribute("href");
                localStorage.setItem(STORAGE_NON_SCRN_SELECTED_EPOCH, href);

                // Save checkbox preference
                var checkbox = document.getElementById("manualSelectInitialRefTimeAddCohort");
                if (checkbox) {
                    try {
                        localStorage.setItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME, checkbox.checked ? "1" : "0");
                    } catch (e) {}
                }

                if (popup) {
                    popup.close();
                }
                clearNonScrnEpochIndex();

                log("Selected epoch; starting Add Cohort Subjects automation");
                setRunMode("epochAddCohort");
                updateRunAllPopupStatus("Running Add Cohort Subjects");
                location.href = location.origin + href + "?autoepochaddcohort=1";
            });

            contentDiv.appendChild(btn);
        });

        // Add checkbox for Manual Select Initial Ref Time
        var checkboxRow = document.createElement("div");
        checkboxRow.style.display = "flex";
        checkboxRow.style.alignItems = "center";
        checkboxRow.style.gap = "8px";
        checkboxRow.style.marginTop = "8px";
        checkboxRow.style.padding = "8px";
        checkboxRow.style.borderTop = "1px solid #444";

        var checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.id = "manualSelectInitialRefTimeAddCohort";
        checkbox.style.cursor = "pointer";

        // Load saved preference
        try {
            var saved = localStorage.getItem(STORAGE_MANUAL_SELECT_INITIAL_REF_TIME);
            if (saved === "1") {
                checkbox.checked = true;
            }
        } catch (e) {}

        var label = document.createElement("label");
        label.htmlFor = "manualSelectInitialRefTimeAddCohort";
        label.textContent = "Manual Select Initial Ref Time";
        label.style.cursor = "pointer";
        label.style.color = "#fff";
        label.style.fontSize = "14px";
        label.style.userSelect = "none";

        checkboxRow.appendChild(checkbox);
        checkboxRow.appendChild(label);
        contentDiv.appendChild(checkboxRow);

        popup = createPopup({
            title: "Select Epoch for Add Cohort Subjects",
            content: contentDiv,
            width: "350px",
            height: "auto",
            maxHeight: "80%"
        });
    }
    //==========================
    // RUN STUDY UPDATES FEATURE
    //==========================
    // This section contains all functions related to study update automation.
    // This feature automates updating study status to ACTIVE and locking eligibility forms.
    // Functions handle study show page routing, edit basics page processing, study metadata
    // navigation, and eligibility form locking workflows.
    //==========================

    // Orchestrate study show page actions for consent, update, or epoch routing.

    function StudyUpdateFunctions() {}
    // Persist informed-consent barcode value.
    function setIcBarcode(code) {
        try {
            localStorage.setItem(STORAGE_IC_BARCODE, String(code));
            log("IC barcode pulled=" + String(code));
        } catch (e) {}
    }

    // Extract primary informed-consent barcode from Study show informedConsent panel.
    function getPrimaryIcBarcodeFromStudyShow() {
        var panel = document.querySelector('div#informedConsent.panel-collapse');
        if (!panel) {
            return "";
        }
        var tbody = panel.querySelector('tbody#informedConsentTbody');
        if (!tbody) {
            return "";
        }
        var rows = tbody.querySelectorAll('tr');
        var i = 0;
        while (i < rows.length) {
            var row = rows[i];
            var tds = row.querySelectorAll('td');
            var j = 0;
            var barcode = "";
            var isPrimary = false;
            while (j < tds.length) {
                var td = tds[j];
                var spans = td.querySelectorAll('span.tooltips');
                var k = 0;
                while (k < spans.length) {
                    var sp = spans[k];
                    var title = sp.getAttribute('data-original-title') + "";
                    var val = (sp.textContent + "").trim();
                    if (title === "Barcode") {
                        barcode = val;
                    }
                    if (title === "Type") {
                        if ((val + "").trim().toLowerCase() === "primary") {
                            isPrimary = true;
                        }
                    }
                    k = k + 1;
                }
                j = j + 1;
            }
            if (isPrimary && barcode) {
                return barcode;
            }
            i = i + 1;
        }
        return "";
    }

    // Return whether to continue epoch processing.
    function getContinueEpoch() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_CONTINUE_EPOCH);
        } catch (e) {}
        var mode = getRunMode();
        if (raw === "1" && (mode === "epoch" || mode === "all")) {
            return true;
        }
        return false;
    }

    // Mark continue-epoch flag in storage.
    function setContinueEpoch() {
        try {
            localStorage.setItem(STORAGE_CONTINUE_EPOCH, "1");
            log("ContinueEpoch set");
        } catch (e) {}
    }

    // Open IC panel and pull primary barcode into storage.
    async function ensureIcPanelOpenAndPullBarcode() {
        var toggle = document.querySelector('a[href="#informedConsent"]');
        if (toggle) {
            toggle.click();
            await sleep(500);
        }
        var code = getPrimaryIcBarcodeFromStudyShow();
        if (code && code.length > 0) {
            setIcBarcode(code);
            return true;
        }
        log("Primary IC barcode not found");
        return false;
    }
    async function processStudyShowPage() {
        if (isPaused()) {
            log("Paused; skipping Study Show automation");
            return;
        }
        var autoConsent = getQueryParam("autoconsent");
        var mode = getRunMode();
        if ((mode === "consent" || mode === "allconsent") || (autoConsent === "1" && mode && mode.length > 0)) {
            var ok = await ensureIcPanelOpenAndPullBarcode();
            location.href = "/secure/study/subjects/list?autoconsent=1";
            return;
        }
        var autoUpdate = getQueryParam("autoupdate");
        var ok2 = (autoUpdate === "1" && mode && mode.length > 0);
        if (ok2) {
            try {
                localStorage.setItem(STORAGE_EDIT_STUDY, "1");
                log("Study edit flag set");
            } catch (e) {}
            var editLink = await waitForSelector('a[href^="/secure/administration/manage/studies/update/"][href$="/basics"]', 5000);
            if (!editLink) {
                return;
            }
            editLink.click();
            log("Study edit basics clicked");
            return;
        }
        var goEpoch = getContinueEpoch();
        var needCheck = null;
        try {
            needCheck = localStorage.getItem(STORAGE_CHECK_ELIG_LOCK);
        } catch (e2) {}
        if (needCheck === "1") {
            try {
                localStorage.removeItem(STORAGE_CHECK_ELIG_LOCK);
            } catch (e3) {}
            log("Routing to Study Metadata for eligibility lock");
            location.href = STUDY_METADATA_URL + "?autoeliglock=1";
            return;
        }
        if (mode === "epoch" || mode === "all" || goEpoch) {
            await processStudyShowRouting();
            return;
        }
    }


    // Choose an epoch (prefer screening) and route to its show page for autoepoch.
    async function processStudyShowRouting() {
        var tbody = await waitForSelector('tbody#epochTableBody', 5000);
        if (!tbody) {
            return;
        }
        var anchors = tbody.querySelectorAll('a[href^="/secure/administration/studies/epoch/show/"]');
        var target = null;
        var i = 0;
        while (i < anchors.length) {
            var a = anchors[i];
            var text = a.textContent + "";
            var match = isScreeningLabel(text);
            if (match) {
                target = a;
                break;
            }
            i = i + 1;
        }
        if (!target && anchors.length > 0) {
            target = anchors[0];
        }
        if (!target) {
            return;
        }
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            return;
        }
        var mode = getRunMode();
        if (mode === "all") {
            updateRunAllPopupStatus("Running Add Cohort Subjects");
        }
        location.href = location.origin + href + "?autoepoch=1";
        log("Routing to epoch show");
    }

    // If edit-study flag present, set study to ACTIVE and save reason.
    async function processStudyEditBasicsPageIfFlag() {
        var flag = null;
        try {
            flag = localStorage.getItem(STORAGE_EDIT_STUDY);
        } catch (e) {}
        if (!flag) {
            return;
        }
        try {
            localStorage.removeItem(STORAGE_EDIT_STUDY);
            log("Cleared study edit flag");
        } catch (e2) {}
        var selectEl = await waitForSelector('select#studyState', 5000);
        if (!selectEl) {
            return;
        }
        var current = selectEl.value + "";
        var mode = getRunMode();
        var alreadyActive = current === "ACTIVE";
        if (alreadyActive) {
            var cancelBtn = document.querySelector('button.btn.default[onclick*="/secure/administration/studies/show"]');
            if (cancelBtn) {
                if (mode === "all") {
                    setContinueEpoch();
                } else {
                    try {
                        localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
                    } catch (e4) {}
                }
                cancelBtn.click();
                log("Study already ACTIVE; returning to show");
            }
            return;
        }
        var need = current !== "ACTIVE";
        if (need) {
            selectEl.value = "ACTIVE";
            var evt = new Event("change", { bubbles: true });
            selectEl.dispatchEvent(evt);
            log("Study state set to ACTIVE");
        }
        var reasonEl = await waitForSelector('textarea#reasonForChange', 5000);
        if (!reasonEl) {
            return;
        }
        reasonEl.value = "Status update";
        var inputEvt = new Event("input", { bubbles: true });
        reasonEl.dispatchEvent(inputEvt);
        var saveBtn = await waitForSelector('button.btn.green[type="submit"]', 5000);
        if (!saveBtn) {
            return;
        }
        if (mode === "all") {
            setContinueEpoch();
        } else {
            try {
                localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
            } catch (e6) {}
        }
        saveBtn.click();
        log("Study basics saved");
    }

    // Route from epoch to cohort show when running epoch/all flows.
    async function processEpochShowPage() {
        var auto = getQueryParam("autoepoch");
        var mode = getRunMode();
        var go = (mode === "epoch" || mode === "all") || (auto === "1" && mode && mode.length > 0);
        if (!go) {
            return;
        }

        var anchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        if (anchors.length === 0) {
            return;
        }
        var target = null;
        var i = 0;
        while (i < anchors.length) {
            var a = anchors[i];
            var text = a.textContent + "";
            var match = isScreeningLabel(text);
            if (match) {
                target = a;
                break;
            }
            i = i + 1;
        }
        if (!target) {
            target = anchors[0];
        }
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            return;
        }
        location.href = location.origin + href + "?autocohort=1";
        log("Routing to cohort show");
    }

    //==========================
    // RUN ACTIVITY PLANS FEATURE
    //==========================
    // This section contains all functions related to running activity plans automation.
    // This feature automates the process of activating activity plans by navigating
    // through plan lists, opening plan show pages, and updating plan states.
    // Functions handle list page processing, plan link extraction, pending ID management,
    // and automatic state updates.
    //==========================

    function RunActivityPlansFunctions() {}
    // Read the after-refresh action string.
    function getAfterRefresh() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_AFTER_REFRESH);
        } catch (e) {}
        if (!raw) {
            return null;
        }
        return raw;
    }


    // Clear the automation run flag from storage.
    function clearRunFlag() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_KEY);
        } catch (e) {}
        if (raw) {
            try {
                localStorage.removeItem(STORAGE_KEY);
                log("Run flag cleared");
            } catch (e2) {}
        }
    }
    // Collect links to plan show pages for running autostate.
    function getPlanLinks() {
        var links = [];
        var rows = document.querySelectorAll("table.table.table-striped.table-bordered.table-hover tbody tr");
        var i = 0;
        while (i < rows.length) {
            var row = rows[i];
            var a = row.querySelector('td.highlight a[href^="/secure/crfdesign/activityplans/show/"]');
            if (a) {
                var href = a.getAttribute("href") + "";
                if (href.length > 0) {
                    var id = extractPlanIdFromHref(href);
                    if (id) {
                        var url = location.origin + href + "?autostate=" + String(id);
                        links.push(url);
                    }
                }
            }
            i = i + 1;
        }
        return links;
    }
    // Extract plan id from a show href.
    function extractPlanIdFromHref(href) {
        var m = href.match(/\/show\/(\d+)\//);
        if (m && m[1]) {
            return m[1];
        }
        m = href.match(/\/show\/(\d+)$/);
        if (m && m[1]) {
            return m[1];
        }
        return null;
    }
    // Parse an autostate param value from a URL.
    function extractAutostateIdFromUrl(url) {
        var u = null;
        try {
            u = new URL(url);
        } catch (e) {}
        if (!u) {
            return null;
        }
        var v = u.searchParams.get("autostate");
        if (v) {
            return v;
        }
        return null;
    }

    // Update Run All popup status
    function updateRunAllPopupStatus(statusText) {
        // Store status in localStorage for persistence
        try {
            localStorage.setItem(STORAGE_RUN_ALL_STATUS, statusText);
        } catch (e) {}

        if (!RUN_ALL_POPUP_REF) {
            return;
        }
        try {
            var statusDiv = RUN_ALL_POPUP_REF.element.querySelector("#runAllStatus");
            if (statusDiv) {
                statusDiv.textContent = statusText;
            }
        } catch (e) {
            log("Error updating Run All popup status: " + e);
        }
    }

    // Store pending autostate ids list.
    function setPendingIds(ids) {
        var payload = JSON.stringify(ids);
        try {
            localStorage.setItem(STORAGE_PENDING, payload);
            log("Pending IDs=" + String(ids.length));
        } catch (e) {}
    }

    // Persist an after-refresh action string.
    function setAfterRefresh(action) {
        try {
            localStorage.setItem(STORAGE_AFTER_REFRESH, String(action));
            log("AfterRefresh=" + String(action));
        } catch (e) {}
    }

    // After activity plan processing, route to next step (study/show) as required.
    function monitorCompletionThenAdvance() {
        var mode = getRunMode();
        if (mode === "activity") {
            location.href = LIST_URL;
            return;
        }
        if (mode === "all") {
            setAfterRefresh("goLockSamplePaths");
            location.href = LIST_URL;
            return;
        }
        setAfterRefresh("goStudyShow");
        location.href = LIST_URL;
    }

    // Background function to lock a single activity plan using HTTP requests
    async function fetchAndLockActivityPlan(planUrl, planName) {
        try {
            log("Locking Activity Plan: " + planName);

            var detailHtml = await fetchPage(planUrl);
            var detailDoc = parseHtml(detailHtml);

            var editLink = detailDoc.querySelector('a[href^="/secure/crfdesign/activityplans/updatestate/"]');
            if (!editLink) {
                log("Edit state link not found for: " + planName);
                return { success: false, message: "Edit state link not found" };
            }

            var updatePath = editLink.getAttribute("href");
            var updateUrl = location.origin + updatePath;

            log("Fetching modal content from: " + updateUrl);
            var modalHtml = await fetchPage(updateUrl);
            var modalDoc = parseHtml(modalHtml);

            log("Modal HTML length: " + String(modalHtml.length));
            var saveButton = modalDoc.querySelector('button#actionButton');
            if (saveButton) {
                log("Save button found: " + saveButton.outerHTML);
            } else {
                log("Save button NOT found in modal");
            }

            var formGroups = modalDoc.querySelectorAll('.form-group');
            if (!formGroups || formGroups.length < 2) {
                log("Form groups not found for: " + planName);
                return { success: false, message: "Form groups not found" };
            }

            var secondFormGroup = formGroups[1];
            var formControlStatic = secondFormGroup.querySelector('p.form-control-static');

            var hasReadyState = false;
            if (formControlStatic) {
                var staticText = (formControlStatic.textContent + "").trim();
                var hasLockIcon = formControlStatic.querySelector('i.fa.fa-lock');
                if (staticText.indexOf("Ready") !== -1 || hasLockIcon) {
                    hasReadyState = true;
                    log("Activity Plan " + planName + " is in Ready state, will lock");
                }
            }

            if (!hasReadyState) {
                log("Activity Plan " + planName + " is not in Ready state, skipping");
                return { success: true, message: "Not in Ready state, skipped" };
            }

            var formElement = modalDoc.querySelector('form');
            if (!formElement) {
                log("Form not found for: " + planName);
                return { success: false, message: "Form not found" };
            }

            var allInputs = formElement.querySelectorAll('input, textarea, select');
            log("Found " + String(allInputs.length) + " form inputs");

            var formData = "";
            var i = 0;
            while (i < allInputs.length) {
                var input = allInputs[i];
                var name = input.getAttribute("name");
                var type = input.getAttribute("type") || "";
                var tagName = input.tagName.toLowerCase();
                var value = "";

                log("Input " + String(i) + ": name=" + String(name) + ", type=" + String(type) + ", tag=" + tagName + ", value=" + String(input.value || input.getAttribute("value") || ""));

                if (name) {
                    if (type === "checkbox" || type === "radio") {
                        if (input.checked || input.hasAttribute("checked")) {
                            value = input.value || input.getAttribute("value") || "on";
                        } else {
                            i = i + 1;
                            continue;
                        }
                    } else if (tagName === "select") {
                        var selectedOption = input.querySelector("option[selected]");
                        if (selectedOption) {
                            value = selectedOption.value || selectedOption.getAttribute("value") || "";
                        } else {
                            var firstOption = input.querySelector("option");
                            value = firstOption ? (firstOption.value || firstOption.getAttribute("value") || "") : "";
                        }
                    } else {
                        value = input.value || input.getAttribute("value") || "";
                    }

                    if (formData.length > 0) {
                        formData += "&";
                    }
                    formData += encodeURIComponent(name) + "=" + encodeURIComponent(value);
                }
                i = i + 1;
            }

            var formAction = formElement.getAttribute("action");
            var formMethod = formElement.getAttribute("method") || "POST";
            log("Form action (ignored): " + formAction);
            log("Form method: " + formMethod);
            log("Form data: " + formData);
            log("Submitting lock for: " + planName);

            var submitUrl = updateUrl;

            log("Submit URL: " + submitUrl);
            var resultHtml = await submitForm(submitUrl, formData);
            var resultDoc = parseHtml(resultHtml);

            var errorAlert = resultDoc.querySelector('div.alert.alert-danger.alert-dismissable');
            if (errorAlert) {
                var errorText = (errorAlert.textContent + "").trim();
                if (errorText.indexOf("Each segment must contain at least one visible activity") !== -1) {
                    log("Activity Plan " + planName + " has validation error (no visible activities), skipping");
                    return { success: true, message: "Validation error, skipped" };
                }
            }

            log("Waiting for page refresh after submission...");
            await sleep(1000);

            var verifyHtml = await fetchPage(planUrl);
            var verifyDoc = parseHtml(verifyHtml);

            var verifyAlert = verifyDoc.querySelector('div.alert.alert-success.alert-dismissable');
            if (verifyAlert) {
                var verifyText = (verifyAlert.textContent + "").trim();
                if (verifyText.indexOf("activity plan has been updated") !== -1 || verifyText.indexOf("Activity plan has been updated") !== -1) {
                    log("Successfully locked: " + planName);
                    return { success: true, message: "Locked successfully" };
                }
            }

            log("Lock submitted but success not confirmed for: " + planName);
            return { success: true, message: "Submitted (unconfirmed)" };

        } catch (error) {
            log("Error locking " + planName + ": " + String(error));
            return { success: false, message: String(error) };
        }
    }

    // Orchestrate opening plan show pages and queuing pending ids when run flag present.
    async function processListPage() {
        log("processListPage start");
        var flag = null;
        try {
            flag = localStorage.getItem(STORAGE_KEY);
        } catch (e) {}
        if (!flag) {
            var after = getAfterRefresh();
            var mode = getRunMode();
            if (after === "goLockSamplePaths" && mode === "all") {
                clearAfterRefresh();
                try {
                    localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1");
                } catch (e) {}
                log("Go Lock Sample Paths after Activity Plans");
                updateRunAllPopupStatus("Running Lock Sample Paths");
                location.href = "https://cenexeltest.clinspark.com/secure/samples/configure/paths";
                return;
            }
            if (after === "goStudyShow" && mode !== "activity") {
                clearAfterRefresh();
                location.href = STUDY_SHOW_URL + "?autoupdate=1";
                log("Go Study Show after refresh");
            }
            return;
        }
        clearRunFlag();

        var mode = getRunMode();
        var showPopup = mode === "activity" || mode === "all";

        var planData = [];
        var rows = document.querySelectorAll("table.table.table-striped.table-bordered.table-hover tbody tr");
        var i = 0;
        while (i < rows.length) {
            var row = rows[i];
            var a = row.querySelector('td.highlight a[href^="/secure/crfdesign/activityplans/show/"]');
            if (a) {
                var href = a.getAttribute("href") + "";
                var planName = (a.textContent + "").trim();
                if (href.length > 0) {
                    var fullUrl = location.origin + href;
                    planData.push({ url: fullUrl, name: planName });
                }
            }
            i = i + 1;
        }

        log("Found " + String(planData.length) + " activity plans to lock");

        if (planData.length === 0) {
            log("No activity plans to lock");
            var mode2 = getRunMode();
            if (mode2 === "all") {
                await sleep(1000);
                log("Continuing to Lock Sample Paths for ALL mode");
                try {
                    localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1");
                } catch (e) {}
                updateRunAllPopupStatus("Running Lock Sample Paths");
                location.href = "https://cenexeltest.clinspark.com/secure/samples/configure/paths";
            }
            return;
        }

        var statusDiv = null;
        var progressDiv = null;
        var loadingAnimation = null;
        var countsDiv = null;
        var loadingInterval = null;

        if (showPopup) {
            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            statusDiv = document.createElement("div");
            statusDiv.id = "lockActivityPlansStatus";
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Locking Activity Plans";

            progressDiv = document.createElement("div");
            progressDiv.id = "lockActivityPlansProgress";
            progressDiv.style.textAlign = "center";
            progressDiv.style.fontSize = "16px";
            progressDiv.style.color = "#9df";
            progressDiv.textContent = "Processing 0/" + String(planData.length);

            loadingAnimation = document.createElement("div");
            loadingAnimation.id = "lockActivityPlansLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            countsDiv = document.createElement("div");
            countsDiv.id = "lockActivityPlansCounts";
            countsDiv.style.textAlign = "center";
            countsDiv.style.fontSize = "14px";
            countsDiv.style.color = "#ccc";
            countsDiv.innerHTML = "<span style='color:#9f9'>Success: 0</span> | <span style='color:#f99'>Failed: 0</span>";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(progressDiv);
            popupContainer.appendChild(loadingAnimation);
            popupContainer.appendChild(countsDiv);

            LOCK_ACTIVITY_PLANS_POPUP_REF = createPopup({
                title: "Lock Activity Plans",
                content: popupContainer,
                width: "400px",
                height: "auto",
                onClose: function() {
                    log("Lock Activity Plans: cancelled by user");
                    try {
                        localStorage.removeItem(STORAGE_KEY);
                        localStorage.removeItem(STORAGE_LOCK_ACTIVITY_PLANS_POPUP);
                    } catch (e) {}
                    LOCK_ACTIVITY_PLANS_POPUP_REF = null;
                }
            });

            try {
                localStorage.setItem(STORAGE_LOCK_ACTIVITY_PLANS_POPUP, "1");
                localStorage.setItem("activityPlanState.lockActivityPlans.successCount", "0");
                localStorage.setItem("activityPlanState.lockActivityPlans.failCount", "0");
            } catch (e) {}

            var dots = 1;
            loadingInterval = setInterval(function() {
                if (!LOCK_ACTIVITY_PLANS_POPUP_REF || !document.body.contains(LOCK_ACTIVITY_PLANS_POPUP_REF.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var d = 0;
                while (d < dots) {
                    text = text + ".";
                    d = d + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);
        }

        var successCount = 0;
        var failCount = 0;
        var j = 0;
        while (j < planData.length) {
            var plan = planData[j];
            log("Processing (" + String(j + 1) + "/" + String(planData.length) + "): " + plan.name);

            if (showPopup && progressDiv) {
                progressDiv.textContent = "Processing " + String(j + 1) + "/" + String(planData.length) + ": " + plan.name;
            }

            var result = await fetchAndLockActivityPlan(plan.url, plan.name);
            if (result.success) {
                successCount = successCount + 1;
            } else {
                failCount = failCount + 1;
            }

            if (showPopup && countsDiv) {
                countsDiv.innerHTML = "<span style='color:#9f9'>Success: " + String(successCount) + "</span> | <span style='color:#f99'>Failed: " + String(failCount) + "</span>";
            }

            await sleep(500);
            j = j + 1;
        }

        if (showPopup && loadingInterval) {
            clearInterval(loadingInterval);
        }

        log("Lock Activity Plans completed. Success=" + String(successCount) + " Failed=" + String(failCount));

        if (showPopup && statusDiv) {
            statusDiv.textContent = "Completed";
            statusDiv.style.color = "#9f9";
        }
        if (showPopup && loadingAnimation) {
            loadingAnimation.textContent = "All activity plans processed";
        }
        if (showPopup && progressDiv) {
            progressDiv.textContent = "Processed " + String(planData.length) + "/" + String(planData.length);
        }

        try {
            localStorage.removeItem(STORAGE_LOCK_ACTIVITY_PLANS_POPUP);
            log("Lock Activity Plans: popup flag cleared");
        } catch (e) {}

        var mode3 = getRunMode();
        if (mode3 === "all") {
            await sleep(2000);
            if (LOCK_ACTIVITY_PLANS_POPUP_REF && LOCK_ACTIVITY_PLANS_POPUP_REF.close) {
                LOCK_ACTIVITY_PLANS_POPUP_REF.close();
            }
            LOCK_ACTIVITY_PLANS_POPUP_REF = null;
            log("Continuing to Lock Sample Paths for ALL mode");
            try {
                localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1");
            } catch (e) {}
            updateRunAllPopupStatus("Running Lock Sample Paths");
            location.href = "https://cenexeltest.clinspark.com/secure/samples/configure/paths";
        }
    }


    //==========================
    // RUN BARCODE FEATURE
    //==========================
    // This section contains all functions related to barcode lookup and data entry.
    // This feature automates finding subject barcodes by searching through epochs,
    // cohorts, and subjects. Functions handle subject identification, barcode retrieval,
    // and populating barcode input modals.
    //==========================

    // Persist barcode subject id for cross-tab lookup.
    function BarcodeFunctions() {}

    // Attempt to infer subject text/id from breadcrumb or tooltip elements.
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
                var m = href.match(/\/show\/subject\/(\d+)/);
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

    function setBarcodeSubjectId(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_SUBJECT_ID, String(t));
        } catch (e) { }
    }

    // Read persisted barcode subject id.
    function getBarcodeSubjectId() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_SUBJECT_ID);
        } catch (e) { }
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    // Clear barcode subject id.
    function clearBarcodeSubjectId() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_SUBJECT_ID);
        } catch (e) { }
    }

    // Persist the barcode result for communication between tabs.
    function setBarcodeResult(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_RESULT, String(t));
        } catch (e) { }
    }

    // Read the persisted barcode result.
    function getBarcodeResult() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_RESULT);
        } catch (e) { }
        if (!raw) {
            return "";
        }
        return String(raw);
    }


    // Persist barcode subject display text for cross-tab lookup.
    function setBarcodeSubjectText(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_SUBJECT_TEXT, String(t));
        } catch (e) { }
    }

    // Read persisted barcode subject display text.
    function getBarcodeSubjectText() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_BARCODE_SUBJECT_TEXT);
        } catch (e) { }
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    // Normalize subject text by trimming and collapsing whitespace.
    function normalizeSubjectString(t) {
        if (typeof t !== "string") {
            return "";
        }
        var s = t.replace(/\u00A0/g, " ");
        s = s.trim();
        s = s.replace(/\s+/g, " ");
        return s;
    }

    // Clear barcode subject text from storage.
    function clearBarcodeSubjectText() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_SUBJECT_TEXT);
        } catch (e) { }
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

    // Return true if currently on the barcode printing subjects page.
    function isBarcodeSubjectsPage() {
        var path = location.pathname;
        if (path === "/secure/barcodeprinting/subjects") {
            return true;
        }
        return false;
    }

    //==========================
    // SHARED UTILITY FUNCTIONS
    //==========================
    // This section contains utility functions used by multiple features across the automation.
    // These include pause/resume control, logging, async helpers (sleep, waitForSelector),
    // storage management, page detection, query parameter parsing, and other common utilities.
    // Functions in this section are shared dependencies used by multiple feature sections.
    //==========================

    function SharedUtilityFunctions() {}


    // Recreate popups on page load if they should be active
    function recreatePopupsIfNeeded() {
        try {
            var runMode = getRunMode();

            // Recreate Run All popup (also for consent phase which is part of Run All flow)
            if (runMode === "all" || runMode === "consent") {
                var runAllPopupActive = localStorage.getItem(STORAGE_RUN_ALL_POPUP);
                if (runAllPopupActive === "1" && (!RUN_ALL_POPUP_REF || !document.body.contains(RUN_ALL_POPUP_REF.element))) {
                    var popupContainer = document.createElement("div");
                    popupContainer.style.display = "flex";
                    popupContainer.style.flexDirection = "column";
                    popupContainer.style.gap = "16px";
                    popupContainer.style.padding = "8px";

                    // Restore status from localStorage
                    var savedStatus = "Running Lock Activity Plans";
                    try {
                        var storedStatus = localStorage.getItem(STORAGE_RUN_ALL_STATUS);
                        if (storedStatus && storedStatus.length > 0) {
                            savedStatus = storedStatus;
                        }
                    } catch (e) {}

                    var statusDiv = document.createElement("div");
                    statusDiv.id = "runAllStatus";
                    statusDiv.style.textAlign = "center";
                    statusDiv.style.fontSize = "18px";
                    statusDiv.style.color = "#fff";
                    statusDiv.style.fontWeight = "500";
                    statusDiv.textContent = savedStatus;

                    var loadingAnimation = document.createElement("div");
                    loadingAnimation.id = "runAllLoading";
                    loadingAnimation.style.textAlign = "center";
                    loadingAnimation.style.fontSize = "14px";
                    loadingAnimation.style.color = "#9df";
                    loadingAnimation.textContent = "Running.";

                    popupContainer.appendChild(statusDiv);
                    popupContainer.appendChild(loadingAnimation);

                    RUN_ALL_POPUP_REF = createPopup({
                        title: "Run Button (1-5) Progress",
                        content: popupContainer,
                        width: "400px",
                        height: "auto",
                        onClose: function() {
                            log("Run All: cancelled by user (close button)");
                            clearAllRunState();
                            clearCohortGuard();
                            try {
                                localStorage.removeItem(STORAGE_RUN_MODE);
                                localStorage.removeItem(STORAGE_KEY);
                                localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
                                localStorage.removeItem(STORAGE_RUN_ALL_POPUP);
                                localStorage.removeItem(STORAGE_RUN_ALL_STATUS);
                            } catch (e) {}
                            RUN_ALL_POPUP_REF = null;
                        }
                    });

                    var dots = 1;
                    var loadingInterval = setInterval(function() {
                        if (!RUN_ALL_POPUP_REF || !document.body.contains(RUN_ALL_POPUP_REF.element)) {
                            clearInterval(loadingInterval);
                            return;
                        }
                        dots = dots + 1;
                        if (dots > 3) {
                            dots = 1;
                        }
                        var text = "Running";
                        var i = 0;
                        while (i < dots) {
                            text = text + ".";
                            i = i + 1;
                        }
                        if (loadingAnimation) {
                            loadingAnimation.textContent = text;
                        }
                    }, 500);
                }
            }

            // Recreate Clear Mapping popup
            if (runMode === RUNMODE_CLEAR_MAPPING) {
                var clearMappingPopupActive = localStorage.getItem(STORAGE_CLEAR_MAPPING_POPUP);
                if (clearMappingPopupActive === "1" && (!CLEAR_MAPPING_POPUP_REF || !document.body.contains(CLEAR_MAPPING_POPUP_REF.element))) {
                    var popupContainer = document.createElement("div");
                    popupContainer.style.display = "flex";
                    popupContainer.style.flexDirection = "column";
                    popupContainer.style.gap = "16px";
                    popupContainer.style.padding = "8px";

                    var statusDiv = document.createElement("div");
                    statusDiv.style.textAlign = "center";
                    statusDiv.style.fontSize = "18px";
                    statusDiv.style.color = "#fff";
                    statusDiv.style.fontWeight = "500";
                    statusDiv.textContent = "Running Clear Mapping";

                    var loadingAnimation = document.createElement("div");
                    loadingAnimation.id = "clearMappingLoading";
                    loadingAnimation.style.textAlign = "center";
                    loadingAnimation.style.fontSize = "14px";
                    loadingAnimation.style.color = "#9df";
                    loadingAnimation.textContent = "Running.";

                    popupContainer.appendChild(statusDiv);
                    popupContainer.appendChild(loadingAnimation);

                    CLEAR_MAPPING_POPUP_REF = createPopup({
                        title: "Clear Mapping",

                        content: popupContainer,
                        width: "350px",
                        height: "auto",
                        onClose: function() {
                            log("ClearMapping: cancelled by user (close button)");
                            try {
                                localStorage.removeItem(STORAGE_RUN_MODE);
                                localStorage.removeItem(STORAGE_CLEAR_MAPPING_POPUP);
                            } catch (e) {}
                            CLEAR_MAPPING_POPUP_REF = null;
                        }
                    });

                    var dots = 1;
                    var loadingInterval = setInterval(function() {
                        if (!CLEAR_MAPPING_POPUP_REF || !document.body.contains(CLEAR_MAPPING_POPUP_REF.element)) {
                            clearInterval(loadingInterval);
                            return;
                        }
                        dots = dots + 1;
                        if (dots > 3) {
                            dots = 1;
                        }
                        var text = "Running";
                        var i = 0;
                        while (i < dots) {
                            text = text + ".";
                            i = i + 1;
                        }
                        if (loadingAnimation) {
                            loadingAnimation.textContent = text;
                        }
                    }, 500);
                }
            }

            // Recreate Import Cohort Subject popup
            if (runMode === "nonscrn" || runMode === "epochImport") {
                var importCohortPopupActive = localStorage.getItem(STORAGE_IMPORT_COHORT_POPUP);
                if (importCohortPopupActive === "1" && (!IMPORT_COHORT_POPUP_REF || !document.body.contains(IMPORT_COHORT_POPUP_REF.element))) {
                    var popupContainer = document.createElement("div");
                    popupContainer.style.display = "flex";
                    popupContainer.style.flexDirection = "column";
                    popupContainer.style.gap = "16px";
                    popupContainer.style.padding = "8px";

                    var statusDiv = document.createElement("div");
                    statusDiv.id = "importCohortStatus";
                    statusDiv.style.textAlign = "center";
                    statusDiv.style.fontSize = "18px";
                    statusDiv.style.color = "#fff";
                    statusDiv.style.fontWeight = "500";
                    statusDiv.textContent = "Running Import Cohort Subject";

                    var loadingAnimation = document.createElement("div");
                    loadingAnimation.id = "importCohortLoading";
                    loadingAnimation.style.textAlign = "center";
                    loadingAnimation.style.fontSize = "14px";
                    loadingAnimation.style.color = "#9df";
                    loadingAnimation.textContent = "Running.";

                    popupContainer.appendChild(statusDiv);
                    popupContainer.appendChild(loadingAnimation);

                    IMPORT_COHORT_POPUP_REF = createPopup({
                        title: "Import Cohort Subject",
                        content: popupContainer,
                        width: "400px",
                        height: "auto",
                        onClose: function() {
                            log("Import Cohort: cancelled by user (close button)");
                            clearAllRunState();
                            clearCohortGuard();
                            try {
                                localStorage.removeItem(STORAGE_RUN_MODE);
                                localStorage.removeItem(STORAGE_IMPORT_COHORT_POPUP);
                            } catch (e) {}
                            IMPORT_COHORT_POPUP_REF = null;
                        }
                    });

                    var dots = 1;
                    var loadingInterval = setInterval(function() {
                        if (!IMPORT_COHORT_POPUP_REF || !document.body.contains(IMPORT_COHORT_POPUP_REF.element)) {
                            clearInterval(loadingInterval);
                            return;
                        }
                        dots = dots + 1;
                        if (dots > 3) {
                            dots = 1;
                        }
                        var text = "Running";
                        var i = 0;
                        while (i < dots) {
                            text = text + ".";
                            i = i + 1;
                        }
                        if (loadingAnimation) {
                            loadingAnimation.textContent = text;
                        }
                    }, 500);
                }
            }

            // Recreate Add Cohort Subjects popup
            if (runMode === "epochAddCohort") {
                var addCohortPopupActive = localStorage.getItem(STORAGE_ADD_COHORT_POPUP);
                if (addCohortPopupActive === "1" && (!ADD_COHORT_POPUP_REF || !document.body.contains(ADD_COHORT_POPUP_REF.element))) {
                    var popupContainer = document.createElement("div");
                    popupContainer.style.display = "flex";
                    popupContainer.style.flexDirection = "column";
                    popupContainer.style.gap = "16px";
                    popupContainer.style.padding = "8px";

                    var statusDiv = document.createElement("div");
                    statusDiv.id = "addCohortStatus";
                    statusDiv.style.textAlign = "center";
                    statusDiv.style.fontSize = "18px";
                    statusDiv.style.color = "#fff";
                    statusDiv.style.fontWeight = "500";
                    statusDiv.textContent = "Running Add Cohort Subjects";

                    var loadingAnimation = document.createElement("div");
                    loadingAnimation.id = "addCohortLoading";
                    loadingAnimation.style.textAlign = "center";
                    loadingAnimation.style.fontSize = "14px";
                    loadingAnimation.style.color = "#9df";
                    loadingAnimation.textContent = "Running.";

                    popupContainer.appendChild(statusDiv);
                    popupContainer.appendChild(loadingAnimation);

                    ADD_COHORT_POPUP_REF = createPopup({
                        title: "Add Cohort Subjects",
                        content: popupContainer,
                        width: "400px",
                        height: "auto",
                        onClose: function() {
                            log("Add Cohort: cancelled by user (close button)");
                            clearAllRunState();
                            clearCohortGuard();
                            try {
                                localStorage.removeItem(STORAGE_RUN_MODE);
                                localStorage.removeItem(STORAGE_ADD_COHORT_POPUP);
                            } catch (e) {}
                            ADD_COHORT_POPUP_REF = null;
                        }
                    });

                    var dots = 1;
                    var loadingInterval = setInterval(function() {
                        if (!ADD_COHORT_POPUP_REF || !document.body.contains(ADD_COHORT_POPUP_REF.element)) {
                            clearInterval(loadingInterval);
                            return;
                        }
                        dots = dots + 1;
                        if (dots > 3) {
                            dots = 1;
                        }
                        var text = "Running";
                        var i = 0;
                        while (i < dots) {
                            text = text + ".";
                            i = i + 1;
                        }
                        if (loadingAnimation) {
                            loadingAnimation.textContent = text;
                        }
                    }, 500);
                }
            }

        } catch (e) {
            log("Error recreating popups: " + e);
        }
    }

    // Clear all Collect All related data
    function clearCollectAllData() {
        COLLECT_ALL_CANCELLED = false;
        COLLECT_ALL_POPUP_REF = null;
        log("CollectAll: data cleared");
    }

    
    // Process eligibility form page to perform the lock action
    async function processEligibilityFormPageForLocking() {
        if (isPaused()) {
            log("Paused; skipping eligibility form locking");
            return;
        }
        var autoLock = getQueryParam("autolock");
        var mode = getRunMode();
        
        // Check if lock was just completed (after page refresh from save)
        var lockCompleted = null;
        try {
            lockCompleted = localStorage.getItem("eligibilityLockCompleted");
        } catch (e) {}
        
        if (lockCompleted === "1") {
            try {
                localStorage.removeItem("eligibilityLockCompleted");
                localStorage.removeItem(STORAGE_CHECK_ELIG_LOCK);
            } catch (e) {}
            
            log("Eligibility lock completed; continuing to Update Study Status");
            if (mode === "study") {
                clearRunMode();
                log("Ending Study Update");
                return;
            }
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        
        if (autoLock !== "1") {
            return;
        }
        
        log("Eligibility form page: attempting to lock");
        
        // Find and click the Action button in breadcrumb
        var actionBtn = await waitForSelector('button.btn.btn-default.btn-sm.dropdown-toggle[data-toggle="dropdown"]', 5000);
        if (!actionBtn) {
            log("Action button not found; cannot lock eligibility");
            if (mode === "study") {
                clearRunMode();
                return;
            }
            log("Continuing to Update Study Status without lock");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        
        actionBtn.click();
        log("Action button clicked");
        await sleep(500);
        
        // Find and click the Lock link in dropdown
        var lockLink = await waitForSelector('a[href*="/secure/crfdesign/studylibrary/locking/form/"][data-toggle="modal"]', 3000);
        if (!lockLink) {
            log("Lock link not found in dropdown; cannot lock eligibility");
            if (mode === "study") {
                clearRunMode();
                return;
            }
            log("Continuing to Update Study Status without lock");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        
        lockLink.click();
        log("Lock link clicked; waiting for modal");
        await sleep(1000);
        
        // Wait for modal and find Save button
        var saveBtn = await waitForSelector('button#actionButton.btn.green', 5000);
        if (!saveBtn) {
            log("Save button not found in modal; cannot complete lock");
            if (mode === "study") {
                clearRunMode();
                return;
            }
            log("Continuing to Update Study Status without lock");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        
        // Set flag before clicking save (page will refresh after save)
        try {
            localStorage.setItem("eligibilityLockCompleted", "1");
        } catch (e) {}
        
        saveBtn.click();
        log("Save button clicked; eligibility lock initiated (page will refresh)");
    }
    
    // Find and navigate to eligibility locking form when required.
    async function processStudyMetadataPageForEligibilityLock() {
        if (isPaused()) {
            log("Paused; skipping study metadata processing");
            return;
        }
        var auto = getQueryParam("autoeliglock");
        var mode = getRunMode();
        var go = auto === "1" || mode === "study" || mode === "all";
        if (!go) {
            return;
        }
        var anchors = document.querySelectorAll('td.smdValue a[href^="/secure/crfdesign/studylibrary/show/form/"]');
        var target = null;
        var i = 0;
        while (i < anchors.length) {
            var a = anchors[i];
            var txt = (a.textContent + "").trim().toLowerCase();
            var hasElig = txt.indexOf("subject eligibility") !== -1;
            if (hasElig) {
                target = a;
                break;
            }
            i = i + 1;
        }
        if (!target) {
            if (mode === "study") {
                clearRunMode();
                log("Eligibility link not found; ending Study Update");
                return;
            }
            log("Eligibility link not found; continuing ALL to Study Show");
            location.href = "/secure/administration/studies/show";
            return;
        }
        var td = target.closest('td.smdValue');
        var hasLock = false;
        if (td) {
            var ico = td.querySelector('i.fa.fa-lock.tooltips');
            if (ico) {
                hasLock = true;
            }
        }
        if (hasLock) {
            if (mode === "study") {
                clearRunMode();
                log("Eligibility already locked; ending Study Update");
                return;
            }
            log("Eligibility locked; continuing to Update Study Status");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        log("Eligibility not locked; navigating to eligibility form to lock it");
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            if (mode === "study") {
                clearRunMode();
                log("Eligibility href missing; ending Study Update");
                return;
            }
            log("Eligibility href missing; continuing to Update Study Status");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
            return;
        }
        var url = location.origin + href;
        if (href.indexOf("?") === -1) {
            url = url + "?autolock=1";
        } else {
            url = url + "&autolock=1";
        }
        log("Routing to Eligibility form for locking");
        location.href = url;
    }

    // Normalize (lowercase, remove punctuation/spaces) a text string used for label matching.
    function normalizeText(t) {
        if (typeof t !== "string") {
            return "";
        }
        var s = t.toLowerCase();
        s = s.replace(/\s+/g, "");
        s = s.replace(/[\-\_\(\)\[\]]/g, "");
        return s;
    }

    // Heuristic to determine if a label denotes a screening/Screening epoch/cohort.
    function isScreeningLabel(t) {
        var s = normalizeText(t);
        var hasScreen = s.indexOf("screen") !== -1;
        var hasScrn = s.indexOf("scrn") !== -1;
        if (hasScreen) {
            return true;
        }
        if (hasScrn) {
            return true;
        }
        return false;
    }

    // Wait until an element disappears or becomes hidden.
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

    // Wait for and click a Bootbox confirmation OK button.
    async function clickBootboxOk(timeoutMs) {
        var okBtn = await waitForSelector('button[data-bb-handler="confirm"].btn.btn-primary', timeoutMs);
        if (!okBtn) {
            okBtn = await waitForSelector('button.btn.btn-primary[data-bb-handler="confirm"]', 2000);
        }
        if (!okBtn) {
            log("Bootbox OK not found");
            return false;
        }
        okBtn.click();
        log("Bootbox OK clicked");
        return true;
    }

    // Click the modal action/save button and wait for the modal to hide.
    async function clickSaveInModal() {
        var modal = await waitForSelector("#ajaxModal, .modal", 5000);
        if (!modal) {
            log("Modal not found for save");
            return false;
        }
        var btn = await waitForSelector("#actionButton", 5000);
        if (!btn) {
            log("ActionButton not found");
            return false;
        }
        btn.click();
        log("ActionButton clicked");
        var ok = await waitUntilHidden("#ajaxModal, .modal", 10000);
        if (!ok) {
            await sleep(1500);
        }
        return true;
    }

    // Extract current plan id from path.
    function getCurrentPlanId() {
        var path = location.pathname;
        var m = path.match(/\/show\/(\d+)\//);
        if (m && m[1]) {
            return m[1];
        }
        m = path.match(/\/show\/(\d+)$/);
        if (m && m[1]) {
            return m[1];
        }
        return null;
    }

    // Detect a success alert indicating cohort assignment saved.
    function hasSuccessAlert() {
        var el = document.querySelector('div.alert.alert-success.alert-dismissable');
        if (!el) {
            return false;
        }
        var t = el.textContent + "";
        var s = t.trim().toLowerCase();
        var match = s.indexOf("cohort assignment has been saved") !== -1;
        if (match) {
            return true;
        }
        return false;
    }

    // Read a query parameter value from the URL.
    function getQueryParam(name) {
        var qs = location.search + "";
        var params = new URLSearchParams(qs);
        var val = params.get(name);
        if (val) {
            return val;
        }
        return null;
    }

    // Click an 'Actions' dropdown if present to reveal action links.
    async function clickActionsDropdownIfNeeded() {
        var dropdowns = document.querySelectorAll("button.dropdown-toggle, a.dropdown-toggle");
        var i = 0;
        while (i < dropdowns.length) {
            var el = dropdowns[i];
            var text = (el.textContent + "").trim().toLowerCase();
            var match = text.indexOf("actions") !== -1 || text.indexOf("action") !== -1;
            if (match) {
                el.click();
                await sleep(200);
                log("Actions dropdown opened");
                return true;
            }
            i = i + 1;
        }
        log("Actions dropdown not found");
        return false;
    }

    // Locate and click the edit-state link for activity plans, opening its modal.
    async function findAndOpenEditStateModal() {
        var link = document.querySelector('a[href^="/secure/crfdesign/activityplans/updatestate/"]');
        if (!link) {
            var opened = await clickActionsDropdownIfNeeded();
            if (opened) {
                link = document.querySelector('a[href^="/secure/crfdesign/activityplans/updatestate/"]');
            }
        }
        if (!link) {
            log("Edit state link not found");
            return false;
        }
        link.click();
        log("Edit state link clicked");
        return true;
    }

    // Read pending autostate ids list.
    function getPendingIds() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_PENDING);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        var ids = [];
        try {
            ids = JSON.parse(raw);
        } catch (e2) {}
        if (Array.isArray(ids)) {
            return ids;
        }
        return [];
    }

    // Remove one id from pending ids list.
    function removePendingId(id) {
        var ids = getPendingIds();
        var out = [];
        var i = 0;
        while (i < ids.length) {
            var v = ids[i];
            if (String(v) !== String(id)) {
                out.push(v);
            }
            i = i + 1;
        }
        setPendingIds(out);
        log("Removed pending id=" + String(id));
    }

    async function processShowPageIfAuto() {
        log("processShowPageIfAuto: not needed with background approach");
    }

    // Read current run mode from storage.
    function getRunMode() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {}
        if (!raw) {
            return "";
        }
        return String(raw);
    }

    // Set run mode in storage.
    function setRunMode(s) {
        try {
            localStorage.setItem(STORAGE_RUN_MODE, String(s));
            log("RunMode=" + String(s));
        } catch (e) {}
    }

    // Open a URL in a tab using GM API if available, otherwise window.open.
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

    // Reliably close the current tab with fallback UI if window.close() fails.
    function closeTabWithFallback(message) {
        var msg = message || "Task completed. You may close this tab.";
        log("closeTabWithFallback: attempting to close tab");

        // Try to close the window
        try {
            window.close();
        } catch (e) {
            log("closeTabWithFallback: window.close() failed: " + String(e));
        }

        // Check if window is still open after a short delay
        setTimeout(function() {
            // If we're still here, the window didn't close - show fallback UI
            if (!window.closed) {
                log("closeTabWithFallback: window still open, showing fallback UI");
                showCloseTabFallbackUI(msg);
            }
        }, 500);

        // Additional close attempts
        setTimeout(function() {
            try { window.close(); } catch (e) {}
        }, 1000);
        setTimeout(function() {
            try { window.close(); } catch (e) {}
        }, 2000);
    }

    // Show a fallback UI when window.close() fails
    function showCloseTabFallbackUI(message) {
        // Check if fallback UI already exists
        if (document.getElementById("closeTabFallbackUI")) {
            return;
        }

        var overlay = document.createElement("div");
        overlay.id = "closeTabFallbackUI";
        overlay.style.cssText = "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.85);z-index:999999;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:20px;";

        var msgDiv = document.createElement("div");
        msgDiv.style.cssText = "color:#fff;font-size:24px;font-weight:600;text-align:center;padding:20px;";
        msgDiv.textContent = message;

        var subMsg = document.createElement("div");
        subMsg.style.cssText = "color:#9df;font-size:16px;text-align:center;";
        subMsg.textContent = "Click the button below or press Ctrl+W to close this tab.";

        var closeBtn = document.createElement("button");
        closeBtn.textContent = "Close This Tab";
        closeBtn.style.cssText = "background:#28a745;color:#fff;border:none;border-radius:8px;padding:15px 40px;font-size:18px;font-weight:600;cursor:pointer;";
        closeBtn.addEventListener("click", function() {
            try { window.close(); } catch (e) {}
            // If still here, try focus trick
            setTimeout(function() {
                try {
                    window.open("", "_self");
                    window.close();
                } catch (e2) {}
            }, 100);
        });

        overlay.appendChild(msgDiv);
        overlay.appendChild(subMsg);
        overlay.appendChild(closeBtn);
        document.body.appendChild(overlay);
    }

    // Promise-based sleep helper.
    function sleep(ms) {
        return new Promise(function (resolve) {
            setTimeout(function () {
                resolve();
            }, ms);
        });
    }

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

    function updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar) {
        var collapsed = getPanelCollapsed();
        if (collapsed) {
            // Collapse: shrink height but DO NOT overwrite stored expanded size
            panel.style.width = scale(PANEL_DEFAULT_WIDTH);
            panel.style.height = scale(PANEL_HEADER_HEIGHT_PX);
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
            // Expand: restore last stored size
            var size = getStoredPanelSize();
            panel.style.width = size.width;
            panel.style.height = size.height;
            panel.style.overflow = "hidden";

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

    // Store panel collapsed state to localStorage.
    function setPanelCollapsed(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_COLLAPSED, flag ? "1" : "0");
        } catch (e) {}
    }

    // Read panel collapsed state from localStorage.
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

    // Store log visibility state to localStorage.
    function setLogVisible(flag) {
        try {
            localStorage.setItem(STORAGE_LOG_VISIBLE, flag ? "1" : "0");
        } catch (e) {}
    }

    // Read log visibility state from localStorage. Defaults to true (visible).
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
            var runMode = getRunMode();
            localStorage.removeItem(STORAGE_RUN_MODE);
            log("RunMode cleared");
            // Clear popup flags
            if (runMode === "all") {
                localStorage.removeItem(STORAGE_RUN_ALL_POPUP);
                localStorage.removeItem(STORAGE_RUN_ALL_STATUS);
                if (RUN_ALL_POPUP_REF && RUN_ALL_POPUP_REF.close) {
                    try {
                        RUN_ALL_POPUP_REF.close();
                    } catch (e2) {}
                }
                RUN_ALL_POPUP_REF = null;
            } else if (runMode === RUNMODE_CLEAR_MAPPING) {
                localStorage.removeItem(STORAGE_CLEAR_MAPPING_POPUP);
                if (CLEAR_MAPPING_POPUP_REF && CLEAR_MAPPING_POPUP_REF.close) {
                    try {
                        CLEAR_MAPPING_POPUP_REF.close();
                    } catch (e3) {}
                }
                CLEAR_MAPPING_POPUP_REF = null;
            } else if (runMode === RUNMODE_ELIG_IMPORT) {
                localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                if (IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.close) {
                    try {
                        IMPORT_ELIG_POPUP_REF.close();
                    } catch (e4) {}
                }
                IMPORT_ELIG_POPUP_REF = null;
            } else if (runMode === "nonscrn" || runMode === "epochImport") {
                localStorage.removeItem(STORAGE_IMPORT_COHORT_POPUP);
                if (IMPORT_COHORT_POPUP_REF && IMPORT_COHORT_POPUP_REF.close) {
                    try {
                        IMPORT_COHORT_POPUP_REF.close();
                    } catch (e5) {}
                }
                IMPORT_COHORT_POPUP_REF = null;
            } else if (runMode === "epochAddCohort") {
                localStorage.removeItem(STORAGE_ADD_COHORT_POPUP);
                if (ADD_COHORT_POPUP_REF && ADD_COHORT_POPUP_REF.close) {
                    try {
                        ADD_COHORT_POPUP_REF.close();
                    } catch (e6) {}
                }
                ADD_COHORT_POPUP_REF = null;
            }
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
    function clearBarcodeResult() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_RESULT);
        } catch (e) {}
    }
    // Clear all run state from storage.
    function clearAllRunState() {
        var runMode = getRunMode();
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
            // Also clear popup flags
            if (runMode === "all") {
                localStorage.removeItem(STORAGE_RUN_ALL_POPUP);
                localStorage.removeItem(STORAGE_RUN_ALL_STATUS);
            } else if (runMode === RUNMODE_CLEAR_MAPPING) {
                localStorage.removeItem(STORAGE_CLEAR_MAPPING_POPUP);
            } else if (runMode === RUNMODE_ELIG_IMPORT) {
                localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
            }
        } catch (e) {}
    }

    // Return true if automation is paused via localStorage.
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

    // Set or clear the paused flag in localStorage.
    function setPaused(flag) {
        try {
            localStorage.setItem(STORAGE_PAUSED, flag ? "1" : "0");
            log("Paused=" + String(flag));
        } catch (e) {}
    }


    // Detect activity plans list page.
    function isListPage() {
        var path = location.pathname;
        var expected = "/secure/crfdesign/activityplans/list";
        if (path === expected) {
            return true;
        }
        return false;
    }

    // Detect activity plan show page.
    function isShowPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/crfdesign/activityplans/show/") !== -1;
        if (ok) {
            return true;
        }
        return false;
    }

    // Detect study show page.
    function isStudyShowPage() {
        var path = location.pathname;
        var expected = "/secure/administration/studies/show";
        if (path === expected) {
            return true;
        }
        return false;
    }

    // Detect study edit basics page.
    function isStudyEditBasicsPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/administration/manage/studies/update/") !== -1;
        if (!ok) {
            return false;
        }
        var end = path.endsWith("/basics");
        if (end) {
            return true;
        }
        return false;
    }

    // Detect epoch show page.
    function isEpochShowPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/administration/studies/epoch/show/") !== -1;
        if (ok) {
            return true;
        }
        return false;
    }

    // Detect cohort show page.
    function isCohortShowPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/administration/studies/cohort/show/") !== -1;
        if (ok) {
            return true;
        }
        return false;
    }

    // Detect subjects list page.
    function isSubjectsListPage() {
        var path = location.pathname;
        var expected = "/secure/study/subjects/list";
        if (path === expected) {
            return true;
        }
        return false;
    }

    // Detect subject show page.
    function isSubjectShowPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/study/subjects/show/") !== -1;
        if (ok) {
            return true;
        }
        return false;
    }

    // Detect if current path is the Study Metadata page.
    function isStudyMetadataPage() {
        var path = location.pathname;
        var expected = "/secure/crfdesign/studylibrary/show/studymetadata";
        if (path === expected) {
            return true;
        }
        return false;
    }
    
    // Detect if current path is an eligibility form page
    function isEligibilityFormPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/crfdesign/studylibrary/show/form/") !== -1;
        return ok;
    }
    //==========================
    // SHARED GUI AND PANEL FUNCTIONS
    //==========================
    // This section contains functions used by multiple features for panel management,
    // visibility control, hotkey handling, and UI interactions. These functions are
    // shared across all automation features and provide the common user interface.
    //==========================
    function SharedPanelFunctions() {}

    // Read stored position value (or return fallback).
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
    // Add an extensible button to the floating panel.
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
        btn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        btn.style.padding = scale(BUTTON_PADDING_PX);
        btn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
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


    // Create a draggable, closeable popup styled like the panel

    // Create a draggable, closeable popup styled like the panel
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
        closeBtn.addEventListener("mouseenter", function() {
            closeBtn.style.background = "#333";
        });
        closeBtn.addEventListener("mouseleave", function() {
            closeBtn.style.background = "transparent";
        });

        closeBtn.addEventListener("click", function() {
            try { clearAllRunState(); } catch (e) {}
            if (onClose) {
                try { onClose(); } catch (e2) { log("Popup onClose error: " + e2); }
            }
            document.removeEventListener("keydown", escapeHandler);
            popup.remove();
        });

        headerBar.appendChild(closeBtn);
        popup.appendChild(headerBar);

        var escapeHandler = function(e) {
            if (e.key === "Escape" && document.body.contains(popup)) {
                try { clearAllRunState(); } catch(e2){}
                if (onClose) {
                    try { onClose(); } catch (e3) { log("Popup onClose error (Escape): " + e3); }
                }
                if (popup && popup.remove) popup.remove();
                document.removeEventListener("keydown", escapeHandler);
            }
        };
        document.addEventListener("keydown", escapeHandler);

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
            } else if (right > vw) {
                newLeft = vw - rect.width;
            }

            if (top < 0) {
                newTop = 0;
            } else if (bottom > vh) {
                newTop = vh - rect.height;
            }

            if (newLeft !== left || newTop !== top) {
                popup.style.left = String(newLeft) + "px";
                popup.style.top = String(newTop) + "px";
                popup.style.transform = "none";
            }
        }

        headerBar.addEventListener("mousedown", function(e) {
            if (e.target === closeBtn || closeBtn.contains(e.target)) {
                return;
            }
            isDragging = true;
            var rect = popup.getBoundingClientRect();
            startX = e.clientX;
            startY = e.clientY;
            startLeft = rect.left;
            startTop = rect.top;
            popup.style.transform = "none";
            e.preventDefault();
        });

        var mouseMoveHandler = function(e) {
            if (!isDragging) {
                return;
            }
            var dx = e.clientX - startX;
            var dy = e.clientY - startY;
            var newLeft = startLeft + dx;
            var newTop = startTop + dy;
            popup.style.left = String(newLeft) + "px";
            popup.style.top = String(newTop) + "px";
            clampPopupPosition();
        };

        var mouseUpHandler = function() {
            if (!isDragging) {
                return;
            }
            isDragging = false;
            clampPopupPosition();
        };

        document.addEventListener("mousemove", mouseMoveHandler);
        document.addEventListener("mouseup", mouseUpHandler);

        document.body.appendChild(popup);

        setTimeout(function() {
            clampPopupPosition();
        }, 0);

        return {
            element: popup,
            close: function() {
                if (onClose) {
                    try {
                        onClose();
                    } catch (e) {
                        log("Popup onClose error: " + String(e));
                    }
                }
                document.removeEventListener("mousemove", mouseMoveHandler);
                document.removeEventListener("mouseup", mouseUpHandler);
                document.removeEventListener("keydown", escapeHandler);
                popup.remove();
            },
            setContent: function(newContent) {
                bodyContainer.innerHTML = "";
                if (typeof newContent === "string") {
                    bodyContainer.innerHTML = newContent;
                } else if (newContent && newContent.nodeType === 1) {
                    bodyContainer.appendChild(newContent);
                }
            }
        };
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
        }
        document.addEventListener("keydown", handler, true);
        window.addEventListener("keydown", handler, true);
        window.__APS_HOTKEY_BOUND = true;
        log("Hotkey: bound for " + String(getPanelHotkey()));
    }


    //==========================
    // MAKE PANEL FUNCTIONS
    //==========================
    // This section contains functions used to create and manage the panel UI.
    // These functions are used to create the panel UI and manage its state.
    //==========================
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
        panel.style.width = savedSize.width || scale(PANEL_DEFAULT_WIDTH);
        if (savedSize.height && savedSize.height !== PANEL_DEFAULT_HEIGHT) {
            panel.style.height = savedSize.height;
        } else {
            panel.style.height = PANEL_DEFAULT_HEIGHT;
        }
        panel.style.boxSizing = "border-box";
        panel.style.overflow = "hidden";

        var headerBar = document.createElement("div");
        headerBar.style.position = "relative";
        headerBar.style.display = "grid";
        headerBar.style.gridTemplateColumns = "auto 1fr auto";
        headerBar.style.alignItems = "center";
        headerBar.style.gap = String(scale(PANEL_HEADER_GAP_PX)) + "px";
        headerBar.style.height = String(scale(PANEL_HEADER_HEIGHT_PX)) + "px";
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
        bodyContainer.style.height = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
        bodyContainer.style.maxHeight = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.boxSizing = "border-box";
        bodyContainer.style.padding = scale(12) + "px";
        var btnRow = document.createElement("div");
        btnRow.style.display = "grid";
        btnRow.style.gridTemplateColumns = "1fr 1fr";
        btnRow.style.gap = "8px";
        btnRowRef = btnRow;
        var runPlansBtn = document.createElement("button");
        runPlansBtn.textContent = "Lock Activity Plans";
        runPlansBtn.style.background = "#4a90e2";
        runPlansBtn.style.color = "#fff";
        runPlansBtn.style.border = "none";
        runPlansBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runPlansBtn.style.padding = scale(BUTTON_PADDING_PX);
        runPlansBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runPlansBtn.style.cursor = "pointer";
        runPlansBtn.style.fontWeight = "500";
        runPlansBtn.style.transition = "background 0.2s";
        runPlansBtn.onmouseenter = () => { runPlansBtn.style.background = "#357abd"; };
        runPlansBtn.onmouseleave = () => { runPlansBtn.style.background = "#4a90e2"; };
        var runStudyBtn = document.createElement("button");
        runStudyBtn.textContent = "Update Study Status";
        runStudyBtn.style.background = "#4a90e2";
        runStudyBtn.style.color = "#fff";
        runStudyBtn.style.border = "none";
        runStudyBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runStudyBtn.style.padding = scale(BUTTON_PADDING_PX);
        runStudyBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runStudyBtn.style.cursor = "pointer";
        runStudyBtn.style.fontWeight = "500";
        runStudyBtn.style.transition = "background 0.2s";
        runStudyBtn.onmouseenter = () => { runStudyBtn.style.background = "#357abd"; };
        runStudyBtn.onmouseleave = () => { runStudyBtn.style.background = "#4a90e2"; };
        var runAddCohortBtn = document.createElement("button");
        runAddCohortBtn.textContent = "Add Cohort Subjects";
        runAddCohortBtn.style.background = "#4a90e2";
        runAddCohortBtn.style.color = "#fff";
        runAddCohortBtn.style.border = "none";
        runAddCohortBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runAddCohortBtn.style.padding = scale(BUTTON_PADDING_PX);
        runAddCohortBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runAddCohortBtn.style.cursor = "pointer";
        runAddCohortBtn.style.fontWeight = "500";
        runAddCohortBtn.style.transition = "background 0.2s";
        runAddCohortBtn.onmouseenter = () => { runAddCohortBtn.style.background = "#357abd"; };
        runAddCohortBtn.onmouseleave = () => { runAddCohortBtn.style.background = "#4a90e2"; };
        var runConsentBtn = document.createElement("button");
        runConsentBtn.textContent = "Run ICF Barcode";
        runConsentBtn.style.background = "#4a90e2";
        runConsentBtn.style.color = "#fff";
        runConsentBtn.style.border = "none";
        runConsentBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runConsentBtn.style.padding = scale(BUTTON_PADDING_PX);
        runConsentBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runConsentBtn.style.cursor = "pointer";
        runConsentBtn.style.fontWeight = "500";
        runConsentBtn.style.transition = "background 0.2s";
        runConsentBtn.onmouseenter = () => { runConsentBtn.style.background = "#357abd"; };
        runConsentBtn.onmouseleave = () => { runConsentBtn.style.background = "#4a90e2"; };
        var runAllBtn = document.createElement("button");
        runAllBtn.textContent = "Run Button (1-5)";
        runAllBtn.style.background = "#5cb85c";
        runAllBtn.style.color = "#fff";
        runAllBtn.style.border = "none";
        runAllBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runAllBtn.style.padding = scale(BUTTON_PADDING_PX);
        runAllBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runAllBtn.style.cursor = "pointer";
        runAllBtn.style.fontWeight = "600";
        runAllBtn.style.transition = "background 0.2s";
        runAllBtn.onmouseenter = () => { runAllBtn.style.background = "#449d44"; };
        runAllBtn.onmouseleave = () => { runAllBtn.style.background = "#5cb85c"; };
        var runNonScrnBtn = document.createElement("button");
        runNonScrnBtn.textContent = "Import Cohort Subject";
        runNonScrnBtn.style.background = "#5b43c7";
        runNonScrnBtn.style.color = "#fff";
        runNonScrnBtn.style.border = "none";
        runNonScrnBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runNonScrnBtn.style.padding = scale(BUTTON_PADDING_PX);
        runNonScrnBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runNonScrnBtn.style.cursor = "pointer";
        runNonScrnBtn.style.fontWeight = "500";
        runNonScrnBtn.style.transition = "background 0.2s";
        runNonScrnBtn.onmouseenter = () => { runNonScrnBtn.style.background = "#4a37a0"; };
        runNonScrnBtn.onmouseleave = () => { runNonScrnBtn.style.background = "#5b43c7"; };

        var addExistingSubjectBtn = document.createElement("button");
        addExistingSubjectBtn.textContent = "Add Existing Subject";
        addExistingSubjectBtn.style.background = "#5b43c7";
        addExistingSubjectBtn.style.color = "#fff";
        addExistingSubjectBtn.style.border = "none";
        addExistingSubjectBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        addExistingSubjectBtn.style.padding = scale(BUTTON_PADDING_PX);
        addExistingSubjectBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        addExistingSubjectBtn.style.cursor = "pointer";
        addExistingSubjectBtn.style.fontWeight = "500";
        addExistingSubjectBtn.style.transition = "background 0.2s";
        addExistingSubjectBtn.onmouseenter = () => { addExistingSubjectBtn.style.background = "#4a37a0"; };
        addExistingSubjectBtn.onmouseleave = () => { addExistingSubjectBtn.style.background = "#5b43c7"; };

        var saBuilderBtn = document.createElement("button");
        saBuilderBtn.textContent = "Scheduled Activities Builder";
        saBuilderBtn.style.background = "#5b43c7";
        saBuilderBtn.style.color = "#fff";
        saBuilderBtn.style.border = "none";
        saBuilderBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        saBuilderBtn.style.padding = scale(BUTTON_PADDING_PX);
        saBuilderBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        saBuilderBtn.style.cursor = "pointer";
        saBuilderBtn.style.fontWeight = "500";
        saBuilderBtn.style.transition = "background 0.2s";
        saBuilderBtn.onmouseenter = () => { saBuilderBtn.style.background = "#4a37a0"; };
        saBuilderBtn.onmouseleave = () => { saBuilderBtn.style.background = "#5b43c7"; };

        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Run Barcode";
        runBarcodeBtn.style.background = "#5b43c7";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        runBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.style.fontWeight = "500";
        runBarcodeBtn.style.transition = "background 0.2s";
        runBarcodeBtn.onmouseenter = () => { runBarcodeBtn.style.background = "#4a37a0"; };
        runBarcodeBtn.onmouseleave = () => { runBarcodeBtn.style.background = "#5b43c7"; };
        var runFormOORBtn = document.createElement("button");
        runFormOORBtn.textContent = "Run Form (OOR) Below Range";
        runFormOORBtn.style.background = "#f0ad4e";
        runFormOORBtn.style.color = "#fff";
        runFormOORBtn.style.border = "none";
        runFormOORBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormOORBtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormOORBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormOORBtn.style.cursor = "pointer";
        runFormOORBtn.style.fontWeight = "500";
        runFormOORBtn.style.transition = "background 0.2s";
        runFormOORBtn.onmouseenter = () => { runFormOORBtn.style.background = "#ec971f"; };
        runFormOORBtn.onmouseleave = () => { runFormOORBtn.style.background = "#f0ad4e"; };
        var runFormOORABtn = document.createElement("button");
        runFormOORABtn.textContent = "Run Form (OOR) Above Range";
        runFormOORABtn.style.background = "#f0ad4e";
        runFormOORABtn.style.color = "#fff";
        runFormOORABtn.style.border = "none";
        runFormOORABtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormOORABtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormOORABtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormOORABtn.style.cursor = "pointer";
        runFormOORABtn.style.fontWeight = "500";
        runFormOORABtn.style.transition = "background 0.2s";
        runFormOORABtn.onmouseenter = () => { runFormOORABtn.style.background = "#ec971f"; };
        runFormOORABtn.onmouseleave = () => { runFormOORABtn.style.background = "#f0ad4e"; };
        var runFormIRBtn = document.createElement("button");
        runFormIRBtn.textContent = "Run Form (In Range)";
        runFormIRBtn.style.background = "#f0ad4e";
        runFormIRBtn.style.color = "#fff";
        runFormIRBtn.style.border = "none";
        runFormIRBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormIRBtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormIRBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormIRBtn.style.cursor = "pointer";
        runFormIRBtn.style.fontWeight = "500";
        runFormIRBtn.style.transition = "background 0.2s";
        runFormIRBtn.onmouseenter = () => { runFormIRBtn.style.background = "#ec971f"; };
        runFormIRBtn.onmouseleave = () => { runFormIRBtn.style.background = "#f0ad4e"; };
        var parseMethodBtn = document.createElement("button");
        parseMethodBtn.textContent = "Item Method Forms";
        parseMethodBtn.style.background = "#4a90e2";
        parseMethodBtn.style.color = "#fff";
        parseMethodBtn.style.border = "none";
        parseMethodBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        parseMethodBtn.style.padding = scale(BUTTON_PADDING_PX);
        parseMethodBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        parseMethodBtn.style.cursor = "pointer";
        parseMethodBtn.onmouseenter = () => { parseMethodBtn.style.background = "#58a1f5"; };
        parseMethodBtn.onmouseleave = () => { parseMethodBtn.style.background = "#4a90e2"; };
        
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
        
        var archiveUpdateFormsBtn = document.createElement("button");
        archiveUpdateFormsBtn.textContent = "Archive/Update Forms";
        archiveUpdateFormsBtn.style.background = "#38dae6";
        archiveUpdateFormsBtn.style.color = "#fff";
        archiveUpdateFormsBtn.style.border = "none";
        archiveUpdateFormsBtn.style.borderRadius = "6px";
        archiveUpdateFormsBtn.style.padding = "8px";
        archiveUpdateFormsBtn.style.cursor = "pointer";
        archiveUpdateFormsBtn.style.fontWeight = "500";
        archiveUpdateFormsBtn.style.transition = "background 0.2s";
        archiveUpdateFormsBtn.onmouseenter = function() { this.style.background = "#2bb9c4"; };
        archiveUpdateFormsBtn.onmouseleave = function() { this.style.background = "#38dae6"; };
        archiveUpdateFormsBtn.addEventListener("click", async function () {
            ARCHIVE_UPDATE_FORMS_CANCELLED = false;
            log("Archive/Update Forms: button clicked");
            await runArchiveUpdateForms();
        });

        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#6c757d";
        pauseBtn.style.color = "#fff";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        pauseBtn.style.padding = scale(BUTTON_PADDING_PX);
        pauseBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        pauseBtn.style.cursor = "pointer";
        pauseBtn.style.fontWeight = "500";
        pauseBtn.style.transition = "background 0.2s";
        pauseBtn.onmouseenter = () => { pauseBtn.style.background = "#5a6268"; };
        pauseBtn.onmouseleave = () => { pauseBtn.style.background = "#6c757d"; };

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
        toggleLogsBtn.style.fontWeight = "500";
        toggleLogsBtn.style.transition = "background 0.2s";
        toggleLogsBtn.onmouseenter = () => { toggleLogsBtn.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = () => { toggleLogsBtn.style.background = "#6c757d"; };

        var runLockSamplePathsBtn = document.createElement("button");
        runLockSamplePathsBtn.textContent = "Lock Sample Paths";
        runLockSamplePathsBtn.style.background = "#4a90e2";
        runLockSamplePathsBtn.style.color = "#fff";
        runLockSamplePathsBtn.style.border = "none";
        runLockSamplePathsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runLockSamplePathsBtn.style.padding = scale(BUTTON_PADDING_PX);
        runLockSamplePathsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runLockSamplePathsBtn.style.cursor = "pointer";
        runLockSamplePathsBtn.style.fontWeight = "500";
        runLockSamplePathsBtn.style.transition = "background 0.2s";
        runLockSamplePathsBtn.onmouseenter = () => { runLockSamplePathsBtn.style.background = "#357abd"; };
        runLockSamplePathsBtn.onmouseleave = () => { runLockSamplePathsBtn.style.background = "#4a90e2"; };

        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import I/E";
        importEligBtn.style.background = "#38dae6";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        importEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        importEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        importEligBtn.style.cursor = "pointer";
        importEligBtn.style.fontWeight = "500";
        importEligBtn.style.transition = "background 0.2s";
        importEligBtn.onmouseenter = () => { importEligBtn.style.background = "#2bb9c4"; };
        importEligBtn.onmouseleave = () => { importEligBtn.style.background = "#38dae6"; };

        var findFormBtn = document.createElement("button");
        findFormBtn.textContent = "Find Form";
        findFormBtn.style.background = "#4a90e2";
        findFormBtn.style.color = "#fff";
        findFormBtn.style.border = "none";
        findFormBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findFormBtn.style.padding = scale(BUTTON_PADDING_PX);
        findFormBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findFormBtn.style.cursor = "pointer";
        findFormBtn.onmouseenter = () => { findFormBtn.style.background = "#58a1f5"; };
        findFormBtn.onmouseleave = () => { findFormBtn.style.background = "#4a90e2"; };

        var findStudyEventsBtn = document.createElement("button");
        findStudyEventsBtn.textContent = "Find Study Events";
        findStudyEventsBtn.style.background = "#4a90e2";
        findStudyEventsBtn.style.color = "#fff";
        findStudyEventsBtn.style.border = "none";
        findStudyEventsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findStudyEventsBtn.style.padding = scale(BUTTON_PADDING_PX);
        findStudyEventsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findStudyEventsBtn.style.cursor = "pointer";
        findStudyEventsBtn.onmouseenter = () => { findStudyEventsBtn.style.background = "#58a1f5"; };
        findStudyEventsBtn.onmouseleave = () => { findStudyEventsBtn.style.background = "#4a90e2"; };

        var clearMappingBtn = document.createElement("button");
        clearMappingBtn.textContent = "Clear Mapping";
        clearMappingBtn.style.background = "#38dae6";
        clearMappingBtn.style.color = "#fff";
        clearMappingBtn.style.border = "none";
        clearMappingBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        clearMappingBtn.style.padding = scale(BUTTON_PADDING_PX);
        clearMappingBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        clearMappingBtn.style.cursor = "pointer";
        clearMappingBtn.style.fontWeight = "500";
        clearMappingBtn.style.transition = "background 0.2s";
        clearMappingBtn.onmouseenter = () => { clearMappingBtn.style.background = "#2bb9c4"; };
        clearMappingBtn.onmouseleave = () => { clearMappingBtn.style.background = "#38dae6"; };

        var collectAllBtn = document.createElement("button");
        collectAllBtn.textContent = "Collect All";
        collectAllBtn.style.background = "#f0ad4e";
        collectAllBtn.style.color = "#fff";
        collectAllBtn.style.border = "none";
        collectAllBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        collectAllBtn.style.padding = scale(BUTTON_PADDING_PX);
        collectAllBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        collectAllBtn.style.cursor = "pointer";
        collectAllBtn.style.fontWeight = "500";
        collectAllBtn.style.transition = "background 0.2s";
        collectAllBtn.onmouseenter = () => { collectAllBtn.style.background = "#ec971f"; };
        collectAllBtn.onmouseleave = () => { collectAllBtn.style.background = "#f0ad4e"; };

        var panelButtons = [
            { el: runPlansBtn, label: "Lock Activity Plans" },
            { el: runLockSamplePathsBtn, label: "Lock Sample Paths" },
            { el: runStudyBtn, label: "Update Study Status" },
            { el: runAddCohortBtn, label: "Add Cohort Subjects" },
            { el: runConsentBtn, label: "Run ICF Consent" },
            { el: runAllBtn, label: "Run Button (1-5)" },
            { el: runNonScrnBtn, label: "Import Cohort Subjects" },
            { el: runBarcodeBtn, label: "Run Barcode" },
            { el: addExistingSubjectBtn, label: "Add Existing Subject" },
            { el: searchMethodsBtn, label: "Search Methods" },
            { el: saBuilderBtn, label: "Scheduled Activities Builder" },
            { el: runFormOORBtn, label: "Run Form (OOR) Below Range" },
            { el: runFormOORABtn, label: "Run Form (OOR) Above Range" },
            { el: runFormIRBtn, label: "Run Form (IR) In Range" },
            { el: collectAllBtn, label: "Collect All" },
            { el: importEligBtn, label: "Import I/E" },
            { el: clearMappingBtn, label: "Clear Mapping" },
            { el: archiveUpdateFormsBtn, label: "Archive/Update Forms" },
            { el: parseMethodBtn, label: "Item Method Forms" },
            { el: findFormBtn, label: "Find Form" },
            { el: findStudyEventsBtn, label: "Find Study Events" },
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
        
        runLockSamplePathsBtn.addEventListener("click", function () {
            try {
                localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1");
            } catch (e) { }

            status.textContent = "Navigating to Sample Paths…";
            log("Run Lock Sample Paths clicked");
            location.href = "https://cenexeltest.clinspark.com/secure/samples/configure/paths";
        });

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

        runNonScrnBtn.addEventListener("click", function () {
            status.textContent = "Preparing non-SCRN subject import...";
            log("Run Add non-SCRN Subject clicked");
            localStorage.setItem(STORAGE_RUN_MODE, "nonscrn");

            // Create progress popup
            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            var statusDiv = document.createElement("div");
            statusDiv.id = "importCohortStatus";
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Running Import Cohort Subject";

            var loadingAnimation = document.createElement("div");
            loadingAnimation.id = "importCohortLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(loadingAnimation);

            IMPORT_COHORT_POPUP_REF = createPopup({
                title: "Import Cohort Subject",
                content: popupContainer,
                width: "400px",
                height: "auto",
                onClose: function() {
                    log("Import Cohort: cancelled by user (close button)");
                    clearAllRunState();
                    clearCohortGuard();
                    try {
                        localStorage.removeItem(STORAGE_RUN_MODE);
                        localStorage.removeItem(STORAGE_IMPORT_COHORT_POPUP);
                    } catch (e) {}
                    IMPORT_COHORT_POPUP_REF = null;
                }
            });

            try {
                localStorage.setItem(STORAGE_IMPORT_COHORT_POPUP, "1");
            } catch (e) {}

            // Animate loading dots
            var dots = 1;
            var loadingInterval = setInterval(function() {
                if (!IMPORT_COHORT_POPUP_REF || !document.body.contains(IMPORT_COHORT_POPUP_REF.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var i = 0;
                while (i < dots) {
                    text = text + ".";
                    i = i + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);

            location.href = STUDY_SHOW_URL + "?autononscrn=1";
        });
        addExistingSubjectBtn.addEventListener("click", async function () {
            log("Add Existing Subject: button clicked");
            await startAddExistingSubject();
        });
        runAddCohortBtn.addEventListener("click", function () {
            status.textContent = "Preparing Add Cohort Subjects...";
            log("Run Add Cohort Subjects clicked");
            try {
                localStorage.setItem(STORAGE_RUN_MODE, "epochAddCohort");
                localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
                log("Cleared cohortAdd.editDoneMap for new run");
            } catch (e) {}

            // Create progress popup
            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            var statusDiv = document.createElement("div");
            statusDiv.id = "addCohortStatus";
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Running Add Cohort Subjects";

            var loadingAnimation = document.createElement("div");
            loadingAnimation.id = "addCohortLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(loadingAnimation);

            ADD_COHORT_POPUP_REF = createPopup({
                title: "Add Cohort Subjects",
                content: popupContainer,
                width: "400px",
                height: "auto",
                onClose: function() {
                    log("Add Cohort: cancelled by user (close button)");
                    clearAllRunState();
                    clearCohortGuard();
                    try {
                        localStorage.removeItem(STORAGE_RUN_MODE);
                        localStorage.removeItem(STORAGE_ADD_COHORT_POPUP);
                    } catch (e) {}
                    ADD_COHORT_POPUP_REF = null;
                }
            });

            try {
                localStorage.setItem(STORAGE_ADD_COHORT_POPUP, "1");
            } catch (e) {}

            // Animate loading dots
            var dots = 1;
            var loadingInterval = setInterval(function() {
                if (!ADD_COHORT_POPUP_REF || !document.body.contains(ADD_COHORT_POPUP_REF.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var i = 0;
                while (i < dots) {
                    text = text + ".";
                    i = i + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);

            location.href = STUDY_SHOW_URL + "?autoaddcohort=1";
        });
        try {
            var rawLogs = localStorage.getItem("activityPlanState.logs");
            if (rawLogs) {
                var logsArr = JSON.parse(rawLogs);
                if (Array.isArray(logsArr)) {
                    var i = 0;
                    while (i < logsArr.length) {
                        var line = document.createElement("div");
                        line.textContent = logsArr[i];
                        logBox.appendChild(line);
                        i = i + 1;
                    }
                    logBox.scrollTop = logBox.scrollHeight;
                }
            }
        } catch (e) {}
        panel.appendChild(bodyContainer);
        var resizeHandle = setupResizeHandle(panel, bodyContainer);
        panel.appendChild(resizeHandle);
        runPlansBtn.addEventListener("click", function () {
            try {
                localStorage.setItem(STORAGE_KEY, "1");
                localStorage.setItem(STORAGE_RUN_MODE, "activity");
                localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
                localStorage.removeItem(STORAGE_AFTER_REFRESH);
            } catch (e) {}
            status.textContent = "Opening Activity Plans…";
            log("Run Activity Plans clicked");
            location.href = LIST_URL;
        });
        runStudyBtn.addEventListener("click", function () {
            try {
                localStorage.setItem(STORAGE_RUN_MODE, "study");
                localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
            } catch (e) {}
            status.textContent = "Navigating to Study Show…";
            log("Run Study Update clicked");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
        });
        runConsentBtn.addEventListener("click", function () {
            try {
                localStorage.setItem(STORAGE_RUN_MODE, "consent");
            } catch (e) {}
            status.textContent = "Navigating to Study Show for Consent…";
            log("Run Informed Consent clicked");

            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            var statusDiv = document.createElement("div");
            statusDiv.id = "icfBarcodeStatus";
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Running ICF Barcode";

            var loadingAnimation = document.createElement("div");
            loadingAnimation.id = "icfBarcodeLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(loadingAnimation);

            ICF_BARCODE_POPUP_REF = createPopup({
                title: "Run ICF Barcode Progress",
                content: popupContainer,
                width: "400px",
                height: "auto",
                onClose: function() {
                    log("ICF Barcode: cancelled by user (close button)");
                    clearAllRunState();
                    try {
                        localStorage.removeItem(STORAGE_RUN_MODE);
                        localStorage.removeItem(STORAGE_ICF_BARCODE_POPUP);
                    } catch (e) {}
                    ICF_BARCODE_POPUP_REF = null;
                }
            });

            try {
                localStorage.setItem(STORAGE_ICF_BARCODE_POPUP, "1");
            } catch (e) {}

            // Animate loading dots
            var dots = 1;
            var loadingInterval = setInterval(function() {
                if (!ICF_BARCODE_POPUP_REF || !document.body.contains(ICF_BARCODE_POPUP_REF.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var i = 0;
                while (i < dots) {
                    text = text + ".";
                    i = i + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);

            location.href = STUDY_SHOW_URL + "?autoconsent=1";
        });

        runAllBtn.addEventListener("click", function () {
            try {
                localStorage.setItem(STORAGE_KEY, "1");
                localStorage.setItem(STORAGE_RUN_MODE, "all");
                localStorage.setItem(STORAGE_CONTINUE_EPOCH, "1");
                localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
            } catch (e) {}
            status.textContent = "Starting ALL: Activity Plans…";
            log("Run ALL clicked");

            // Create progress popup
            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            var statusDiv = document.createElement("div");
            statusDiv.id = "runAllStatus";
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Running Lock Activity Plans";

            var loadingAnimation = document.createElement("div");
            loadingAnimation.id = "runAllLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(loadingAnimation);

            RUN_ALL_POPUP_REF = createPopup({
                title: "Run Button (1-5) Progress",
                content: popupContainer,
                width: "400px",
                height: "auto",
                onClose: function() {
                    log("Run All: cancelled by user (close button)");
                    clearAllRunState();
                    clearCohortGuard();
                    try {
                        localStorage.removeItem(STORAGE_RUN_MODE);
                        localStorage.removeItem(STORAGE_KEY);
                        localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
                        localStorage.removeItem(STORAGE_RUN_ALL_POPUP);
                    } catch (e) {}
                    RUN_ALL_POPUP_REF = null;
                }
            });

            try {
                localStorage.setItem(STORAGE_RUN_ALL_POPUP, "1");
            } catch (e) {}

            // Animate loading dots
            var dots = 1;
            var loadingInterval = setInterval(function() {
                if (!RUN_ALL_POPUP_REF || !document.body.contains(RUN_ALL_POPUP_REF.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var i = 0;
                while (i < dots) {
                    text = text + ".";
                    i = i + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);

            location.href = LIST_URL;
            clearCohortGuard();
        });
        pauseBtn.addEventListener("click", function () {
            var nowPaused = isPaused();
            if (nowPaused) {
                setPaused(false);
                pauseBtn.textContent = "Pause";
                status.textContent = "Resumed";
                SA_BUILDER_PAUSE = false;
                CLEAR_MAPPING_PAUSE = false;
                log("Resumed");
            } else {
                setPaused(true);
                pauseBtn.textContent = "Resume";
                status.textContent = "Paused";
                log("Paused");
                
                // Update all active popup statuses to show "Paused"
                if (RUN_ALL_POPUP_REF && RUN_ALL_POPUP_REF.element) {
                    try {
                        var runAllStatusDiv = RUN_ALL_POPUP_REF.element.querySelector("#runAllStatus");
                        if (runAllStatusDiv) {
                            runAllStatusDiv.textContent = "⏸ Paused";
                            runAllStatusDiv.style.color = "#ffa500";
                        }
                        var runAllLoading = RUN_ALL_POPUP_REF.element.querySelector("#runAllLoading");
                        if (runAllLoading) {
                            runAllLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                if (LOCK_SAMPLE_PATHS_POPUP_REF && LOCK_SAMPLE_PATHS_POPUP_REF.element) {
                    try {
                        var lockStatusDiv = LOCK_SAMPLE_PATHS_POPUP_REF.element.querySelector("#lockSamplePathsStatus");
                        if (lockStatusDiv) {
                            lockStatusDiv.textContent = "⏸ Paused";
                            lockStatusDiv.style.color = "#ffa500";
                        }
                        var lockLoading = LOCK_SAMPLE_PATHS_POPUP_REF.element.querySelector("#lockSamplePathsLoading");
                        if (lockLoading) {
                            lockLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                if (IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.element) {
                    try {
                        var eligStatusDiv = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligStatus");
                        if (eligStatusDiv) {
                            eligStatusDiv.textContent = "⏸ Paused";
                            eligStatusDiv.style.color = "#ffa500";
                        }
                        var eligLoading = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligLoading");
                        if (eligLoading) {
                            eligLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                if (IMPORT_COHORT_POPUP_REF && IMPORT_COHORT_POPUP_REF.element) {
                    try {
                        var cohortStatusDiv = IMPORT_COHORT_POPUP_REF.element.querySelector("#importCohortStatus");
                        if (cohortStatusDiv) {
                            cohortStatusDiv.textContent = "⏸ Paused";
                            cohortStatusDiv.style.color = "#ffa500";
                        }
                        var cohortLoading = IMPORT_COHORT_POPUP_REF.element.querySelector("#importCohortLoading");
                        if (cohortLoading) {
                            cohortLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                if (COLLECT_ALL_POPUP_REF && COLLECT_ALL_POPUP_REF.element) {
                    try {
                        var collectStatusDiv = COLLECT_ALL_POPUP_REF.element.querySelector("#collectAllStatus");
                        if (collectStatusDiv) {
                            collectStatusDiv.textContent = "⏸ Paused";
                            collectStatusDiv.style.color = "#ffa500";
                        }
                        var collectLoading = COLLECT_ALL_POPUP_REF.element.querySelector("#collectAllLoading");
                        if (collectLoading) {
                            collectLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                if (CLEAR_MAPPING_POPUP_REF && CLEAR_MAPPING_POPUP_REF.element) {
                    try {
                        var clearStatusDiv = CLEAR_MAPPING_POPUP_REF.element.querySelector("#clearMappingStatus");
                        if (clearStatusDiv) {
                            clearStatusDiv.textContent = "⏸ Paused";
                            clearStatusDiv.style.color = "#ffa500";
                        }
                        var clearLoading = CLEAR_MAPPING_POPUP_REF.element.querySelector("#clearMappingLoading");
                        if (clearLoading) {
                            clearLoading.textContent = "";
                        }
                    } catch (e) {}
                }
                
                clearAllRunState();
                COLLECT_ALL_CANCELLED = true;
                SA_BUILDER_PAUSE = true;
                SA_BUILDER_CANCELLED = true;
                CLEAR_MAPPING_CANCELED = true;
                clearCollectAllData();
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
        runBarcodeBtn.addEventListener("click", async function () {
            log("Run Barcode: button clicked");
            await APS_RunBarcode();
        });
        runFormOORBtn.addEventListener("click", function () {
            RUN_FORM_V2_START_TS = Date.now();
            log("Run Form (OOR) A clicked");
            setFormValueMode("oorA");
            runFormAutomationV2();
        });
        runFormOORABtn.addEventListener("click", function () {
            RUN_FORM_V2_START_TS = Date.now();
            log("Run Form (OOR) B clicked");
            setFormValueMode("oorB");
            runFormAutomationV2();
        });
        runFormIRBtn.addEventListener("click", function () {
            RUN_FORM_V2_START_TS = Date.now();
            log("Run Form (IR) clicked");
            setFormValueMode("ir");
            runFormAutomationV2();
        });
        collectAllBtn.addEventListener("click", function () {
            log("CollectAll: button clicked");
            runCollectAll();
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
        importEligBtn.addEventListener("click", function () {
            log("ImportElig: button clicked");
            startImportEligibilityMapping();
        });
        parseMethodBtn.addEventListener("click", function () {
            openParseMethod();
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
        findFormBtn.addEventListener("click", function () {
            openFindForm();
        });
        findStudyEventsBtn.addEventListener("click", function () {
            openFindStudyEvents();
        });
        clearMappingBtn.addEventListener("click", function () {
            log("ClearMapping: button clicked");
            CLEAR_MAPPING_CANCELED = false;
            if (CLEAR_MAPPING_PAUSE == true) {
                log("ClearMapping: Paused");
                return;
            }
            // Create running animation popup
            var popupContainer = document.createElement("div");
            popupContainer.style.display = "flex";
            popupContainer.style.flexDirection = "column";
            popupContainer.style.gap = "16px";
            popupContainer.style.padding = "8px";

            var statusDiv = document.createElement("div");
            statusDiv.style.textAlign = "center";
            statusDiv.style.fontSize = "18px";
            statusDiv.style.color = "#fff";
            statusDiv.style.fontWeight = "500";
            statusDiv.textContent = "Running Clear Mapping";

            var loadingAnimation = document.createElement("div");
            loadingAnimation.id = "clearMappingLoading";
            loadingAnimation.style.textAlign = "center";
            loadingAnimation.style.fontSize = "14px";
            loadingAnimation.style.color = "#9df";
            loadingAnimation.textContent = "Running.";

            popupContainer.appendChild(statusDiv);
            popupContainer.appendChild(loadingAnimation);

            var clearMappingPopup = createPopup({
                title: "Clear Mapping",
                content: popupContainer,
                width: "350px",
                height: "auto",
                onClose: function() {
                    log("ClearMapping: cancelled by user (close button)");
                    try {
                        localStorage.removeItem(STORAGE_RUN_MODE);
                        localStorage.removeItem(STORAGE_CLEAR_MAPPING_POPUP);
                    } catch (e) {}
                    CLEAR_MAPPING_POPUP_REF = null;
                    CLEAR_MAPPING_CANCELED = true;
                }
            });

            CLEAR_MAPPING_POPUP_REF = clearMappingPopup;
            try {
                localStorage.setItem(STORAGE_CLEAR_MAPPING_POPUP, "1");
            } catch (e) {}

            // Animate loading dots
            var dots = 1;
            var loadingInterval = setInterval(function() {
                if (!clearMappingPopup || !document.body.contains(clearMappingPopup.element)) {
                    clearInterval(loadingInterval);
                    return;
                }
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Running";
                var i = 0;
                while (i < dots) {
                    text = text + ".";
                    i = i + 1;
                }
                if (loadingAnimation) {
                    loadingAnimation.textContent = text;
                }
            }, 500);

            // Close popup when done (check periodically)
            var checkInterval = setInterval(function() {
                try {
                    var runMode = localStorage.getItem(STORAGE_RUN_MODE);
                    if (runMode !== RUNMODE_CLEAR_MAPPING) {
                        clearInterval(loadingInterval);
                        clearInterval(checkInterval);
                        if (clearMappingPopup && clearMappingPopup.close) {
                            try {
                                localStorage.removeItem(STORAGE_CLEAR_MAPPING_POPUP);
                            } catch (e) {}
                            clearMappingPopup.close();
                        }
                        CLEAR_MAPPING_POPUP_REF = null;
                    }
                } catch (e) {}
            }, 1000);

            startClearMapping();
        });
        closeBtn.addEventListener("click", function () {
            panel.remove();
        });
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
            try {
                localStorage.setItem("activityPlanState.panel.top", panel.style.top);
                localStorage.setItem("activityPlanState.panel.right", panel.style.right);
            } catch (e) {}
        });
        document.body.appendChild(panel);
        applyPanelHiddenState(panel);

        updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar);

        var t = pxToInt(panel.style.top);
        var r2 = pxToInt(panel.style.right);
        var vh = window.innerHeight || 800;
        var vw = window.innerWidth || 1280;
        var ph = panel.offsetHeight || 40;
        var pw = panel.offsetWidth || 340;
        var minTop = 0;
        var maxTop = vh - ph;
        var minRight = 0;
        var maxRight = vw - pw;
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
            var i2 = 0;
            while (i2 < PENDING_BUTTONS.length) {
                var it = PENDING_BUTTONS[i2];
                addButtonToPanel(it.label, it.handler);
                i2 = i2 + 1;
            }
            PENDING_BUTTONS = [];
        }
        log("Panel ready");
        return panel;
    }


    // Initialize the panel, register APS_AddButton, and route to the appropriate processing function.
    function init() {
        makePanel();
        window.APS_AddButton = function (label, handler) {
            addButtonToPanel(label, handler);
        };
        bindPanelHotkeyOnce();

        // Recreate popups if needed
        recreatePopupsIfNeeded();

        if (isPaused()) {
            log("Paused; automation halted");
            return;
        }

        // Check for Add Existing Subject mode first
        var runModeRaw = null;
        try {
            runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            runModeRaw = null;
        }
        if (runModeRaw === RUNMODE_ADD_EXISTING_SUBJECT) {
            processAESOnPageLoad();
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
        var isSamplePathsPage = location.pathname === "/secure/samples/configure/paths";
        if (isSamplePathsPage) {
            processLockSamplePathsPage();
            return;
        }

        var isSamplePathDetail = location.pathname.indexOf("/secure/samples/configure/paths/show/") !== -1;
        if (isSamplePathDetail) {
            processLockSamplePathDetailPage();
            return;
        }

        var isSamplePathUpdate = location.pathname.indexOf("/secure/samples/configure/paths/update/") !== -1;
        if (isSamplePathUpdate) {
            processLockSamplePathUpdatePage();
            return;
        }

        var onList = isListPage();
        if (onList) {
            processListPage();
            return;
        }
        var onShow = isShowPage();
        if (onShow) {
            processShowPageIfAuto();
            return;
        }
        var onStudyShow = isStudyShowPage();
        if (onStudyShow) {
            processStudyShowPage();
            processStudyShowPageForNonScrn();
            processStudyShowPageForAddCohort();
            return;
        }
        var onEditBasics = isStudyEditBasicsPage();
        if (onEditBasics) {
            processStudyEditBasicsPageIfFlag();
            return;
        }

        var onEpoch = isEpochShowPage();
        if (onEpoch) {
            var imode = getRunMode();
            if (imode === "epochImport") {
                processEpochShowPageForImport();
                return;
            }
            if (imode === "epochAddCohort") {
                processEpochShowPageForAddCohort();
                return;
            }
            processEpochShowPage();
            return;
        }

        var onCohort = isCohortShowPage();
        if (onCohort) {
            var amode = getRunMode();
            var autoImport = getQueryParam("autocohortimport");
            var autoAdd = getQueryParam("autocohortadd");
            if (amode === "epochImport" || autoImport === "1") {
                processCohortShowPageImportNonScrn();
            } else if (amode === "epochAddCohort" || autoAdd === "1") {
                processCohortShowPageAddCohort();
            } else {
                processCohortShowPage();
            }
            return;
        }
        var onMetadata = isStudyMetadataPage();
        if (onMetadata) {
            processStudyMetadataPageForEligibilityLock();
            return;
        }
        var onEligibilityForm = isEligibilityFormPage();
        if (onEligibilityForm) {
            processEligibilityFormPageForLocking();
            return;
        }
        var onSubjectsList = isSubjectsListPage();
        if (onSubjectsList) {
            processSubjectsListPageForConsent();
            return;
        }
        var onSubjectShow = isSubjectShowPage();
        if (onSubjectShow) {
            processSubjectShowPageForConsent();
            return;
        }
        var onBarcodeSubjects = isBarcodeSubjectsPage();
        if (onBarcodeSubjects) {
            processBarcodeSubjectsPage();
            return;
        }

        // Check for manual navigation away from Import Eligibility process
        var pendingPopup = null;
        try {
            pendingPopup = localStorage.getItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
        } catch (e) {
            pendingPopup = null;
        }

        // Re-use runModeRaw from earlier check or get it again if needed
        if (!runModeRaw) {
            try {
                runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
            } catch (e) {
                runModeRaw = null;
            }
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

            // Recreate popup if it should be active and continue execution
            var importEligPopupActive = localStorage.getItem(STORAGE_IMPORT_ELIG_POPUP);
            if (importEligPopupActive === "1" && (!IMPORT_ELIG_POPUP_REF || !document.body.contains(IMPORT_ELIG_POPUP_REF.element))) {
                // Recreate the running popup
                var container = document.createElement("div");
                container.style.display = "flex";
                container.style.flexDirection = "column";
                container.style.gap = "12px";
                container.style.padding = "8px";

                var runningContainer = document.createElement("div");
                runningContainer.id = "importEligRunningContainer";
                runningContainer.style.display = "flex";
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

                // Animate loading dots
                var dots = 1;
                var loadingInterval = setInterval(function() {
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
                    if (runningBox && document.body.contains(runningBox)) {
                        runningBox.textContent = t;
                    } else {
                        clearInterval(loadingInterval);
                    }
                }, 350);

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

                IMPORT_ELIG_POPUP_REF = createPopup({
                    title: "Import I/E",
                    content: container,
                    width: "500px",
                    height: "auto",
                    maxHeight: "80vh",
                    onClose: function () {
                        log("ImportElig: popup X pressed → stopping automation");
                        clearAllRunState();
                        clearEligibilityWorkingState();
                        try {
                            localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
                            localStorage.removeItem(STORAGE_IMPORT_ELIG_POPUP);
                        } catch (e) {}
                        IMPORT_ELIG_POPUP_REF = null;
                    }
                });

                // Continue execution after recreating popup
                if (isEligibilityListPage()) {
                    log("ImportElig: popup recreated, continuing execution");
                    setTimeout(function () {
                        executeEligibilityMappingAutomation();
                    }, 1000);
                    return;
                }
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

        if (runModeRaw === RUNMODE_CLEAR_MAPPING) {
            if (CLEAR_MAPPING_CANCELED) {
                return;
            }
            if (!isEligibilityListPage()) {
                log("ClearMapping: run mode set but not on list page; redirecting");
                location.href = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
                return;
            }
            log("ClearMapping: run mode set on list page; resuming in 4s");
            setTimeout(function () {
                if (CLEAR_MAPPING_CANCELED) {
                    return;
                }
                logger("CLEAR_MAPPING_CANCELED: " + CLEAR_MAPPING_CANCELED)
                executeClearMappingAutomation();
            }, 4000);
            return;
        }
        if (location.pathname === "/secure/study/data/list") {
            processFindFormOnList();
        }

    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();