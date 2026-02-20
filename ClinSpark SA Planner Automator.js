
// ==UserScript==
// @name ClinSpark SA Builder Automator
// @namespace vinh.activity.plan.state
// @version 1.1.2
// @description
// @match https://cenexeltest.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Basic%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Basic%20Automator.js
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// ==/UserScript==

(function () {
    var STORAGE_PANEL_WIDTH = "activityPlanState.panel.width";
    var STORAGE_PANEL_HEIGHT = "activityPlanState.panel.height";
    var PANEL_DEFAULT_WIDTH = "340px";
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 48;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 60
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
    var STUDY_EVENT_POPUP_KEYWORD_LABEL = "Study Event Keyword";
    var STUDY_EVENT_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var STUDY_EVENT_POPUP_OK_TEXT = "Continue";
    var STUDY_EVENT_POPUP_CANCEL_TEXT = "Cancel";

    var STUDY_EVENT_NO_MATCH_TITLE = "Find Study Events";
    var STUDY_EVENT_NO_MATCH_MESSAGE = "No study event is found.";

    // Run Find Form
    var FORM_LIST_URL = "https://cenexeltest.clinspark.com/secure/study/data/list";
    var FORM_POPUP_TITLE = "Find Form";
    var FORM_POPUP_KEYWORD_LABEL = "Form Keyword";
    var FORM_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var FORM_POPUP_OK_TEXT = "Continue";
    var FORM_POPUP_CANCEL_TEXT = "Cancel";

    var FORM_NO_MATCH_TITLE = "Find Form";
    var FORM_NO_MATCH_MESSAGE = "No form is found.";

    var BARCODE_BG_TAB = null;
    const RUNMODE_CLEAR_MAPPING = "clearMapping";

    // Run Parse Method
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
    const STORAGE_RUN_ALL_POPUP = "activityPlanState.runAllPopup";
    const STORAGE_ICF_BARCODE_POPUP = "activityPlanState.icfBarcodePopup";
    const STORAGE_RUN_ALL_STATUS = "activityPlanState.runAllStatus";
    const STORAGE_CLEAR_MAPPING_POPUP = "activityPlanState.clearMappingPopup";
    const STORAGE_IMPORT_ELIG_POPUP = "activityPlanState.importEligPopup";
    const STORAGE_IMPORT_COHORT_POPUP = "activityPlanState.importCohortPopup";
    const STORAGE_ADD_COHORT_POPUP = "activityPlanState.addCohortPopup";

    // Run Subject Eligibility
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
            return { source: selectedSource, target: selectedTarget };
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
            offsetSeconds: ""
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

        var normalizedTarget = normalizeText(targetText);
        var opts = sel.querySelectorAll("option");
        var matchValue = null;

        for (var i = 0; i < opts.length; i++) {
            var opt = opts[i];
            var optText = normalizeSAText(opt.textContent);
            var optNorm = normalizeText(optText);
            if (optNorm === normalizedTarget || optText === targetText) {
                matchValue = opt.value;
                break;
            }
        }

        // Fuzzy match if exact not found
        if (!matchValue) {
            for (var j = 0; j < opts.length; j++) {
                var opt2 = opts[j];
                var optText2 = normalizeText(opt2.textContent);
                if (optText2.indexOf(normalizedTarget) !== -1 || normalizedTarget.indexOf(optText2) !== -1) {
                    matchValue = opt2.value;
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
    async function executeArchiveUpdateForms(sourceFormKey, targetFormValue, sourceFormsMap, targetFormsArray) {
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
                    await sleep(500);
                    row = findRowInDOM(occ.segmentText, occ.eventText, occ.formText);
                    if (row) {
                        await clickRowActionDropdown(row);
                        var visLink = row.querySelector('a[href*="visibility"]');
                        if (visLink) {
                            visLink.click();
                            var visModal = await waitForSAModal(10000);
                            if (visModal) {
                                visibilityProps = await collectVisibilityModalProperties();
                                var visCancelBtn = visModal.querySelector("button[data-dismiss='modal'], .btn-default, .close");
                                if (visCancelBtn) {
                                    visCancelBtn.click();
                                } else {
                                    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
                                }
                                await waitForSAModalClose(5000);
                            }
                        }
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
                                    reasonEl.value = "Update form";
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
                if (editProps && editProps.hidden && visibilityProps) {
                    // Find the newly added row
                    var newRow = findRowInDOM(occ.segmentText, occ.eventText, targetForm.text);
                    if (newRow) {
                        await clickRowActionDropdown(newRow);
                        var newVisLink = newRow.querySelector('a[href*="visibility"]');
                        if (newVisLink) {
                            newVisLink.click();
                            var newVisModal = await waitForSAModal(10000);
                            if (newVisModal) {
                                await sleep(500);

                                // Set visibility fields
                                if (visibilityProps.activityPlan) {
                                    await setSelect2ValueByText("visibleActivityPlan", visibilityProps.activityPlan);
                                    await waitForSelect2OptionsChange("visibleScheduledActivity", 3000);
                                }
                                if (visibilityProps.scheduledActivity) {
                                    await setSelect2ValueByText("visibleScheduledActivity", visibilityProps.scheduledActivity);
                                    await waitForSelect2OptionsChange("visibleItemRef", 3000);
                                }
                                if (visibilityProps.item) {
                                    await setSelect2ValueByText("visibleItemRef", visibilityProps.item);
                                    await waitForSelect2OptionsChange("visibleCodeListItem", 3000);
                                }
                                if (visibilityProps.itemValue) {
                                    await setSelect2ValueByText("visibleCodeListItem", visibilityProps.itemValue);
                                }

                                // Set reason for change
                                var visReasonEl = document.getElementById("reasonForChange");
                                if (visReasonEl) {
                                    visReasonEl.value = "Add visibility condition";
                                    visReasonEl.dispatchEvent(new Event("change", { bubbles: true }));
                                }

                                // Save visibility
                                var visSaveBtn = document.getElementById("actionButton");
                                if (visSaveBtn) {
                                    visSaveBtn.click();
                                    await waitForSAModalClose(10000);
                                }
                                await sleep(500);
                            }
                        }
                    } else {
                        log("Archive/Update Forms: could not find newly added row to set visibility");
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

            log("Archive/Update Forms: confirmed - source=" + selection.source + " target=" + selection.target);

            // Close selection popup (but don't cancel)
            if (ARCHIVE_UPDATE_FORMS_POPUP_REF) {
                // Remove the onClose handler to prevent it from running
                var popupEl = ARCHIVE_UPDATE_FORMS_POPUP_REF.element;
                if (popupEl) popupEl.remove();
                ARCHIVE_UPDATE_FORMS_POPUP_REF = null;
            }

            // Execute the replacement
            await executeArchiveUpdateForms(selection.source, selection.target, guiContent.sourceFormsMap, guiContent.targetFormsArray);
        });

        log("Archive/Update Forms: selection GUI displayed");
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
            log("Eligibility already locked; proceeding to update study status");
            var editLink = document.querySelector('a[href^="/secure/administration/manage/studies/update/"][href$="/basics"]');
            if (!editLink) {
                log("Edit basics link not found; navigating to Study Show");
                location.href = "/secure/administration/studies/show";
                return;
            }
            editLink.click();
            log("Routing to edit basics to update study status");
            return;
        }
        var href = target.getAttribute("href") + "";
        if (href.length === 0) {
            if (mode === "study") {
                clearRunMode();
                log("Eligibility href missing; ending Study Update");
                return;
            }
            log("Eligibility href missing; continuing ALL to Study Show");
            location.href = "/secure/administration/studies/show";
            return;
        }
        var url = location.origin + href;
        if (href.indexOf("?") === -1) {
            url = url + "?autolock=1";
        } else {
            url = url + "&autolock=1";
        }
        try {
            localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
        } catch (e7) {}
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

    // Store pending autostate ids list.
    function setPendingIds(ids) {
        var payload = JSON.stringify(ids);
        try {
            localStorage.setItem(STORAGE_PENDING, payload);
            log("Pending IDs=" + String(ids.length));
        } catch (e) {}
    }

    // If autostate param present, open and save edit state modal and close tab.
    // NOTE: This function is now stubbed out because Lock Activity Plans uses background HTTP requests.
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
            w = PANEL_DEFAULT_WIDTH;
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
            panel.style.width = PANEL_DEFAULT_WIDTH;
            panel.style.height = String(PANEL_HEADER_HEIGHT_PX + 12) + "px"; // Add padding
            panel.style.overflow = "hidden";

            if (bodyContainer) {
                bodyContainer.style.display = "none";
            }
            if (resizeHandle) {
                resizeHandle.style.display = "none";
            }
            if (collapseBtn) {
                collapseBtn.textContent = "Expand";
            }
        } else {
            // Expand: restore last stored size
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
                collapseBtn.textContent = "Collapse";
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
        // Clear SA Builder storage and set cancelled flag
        SA_BUILDER_CANCELLED = true;
        clearSABuilderStorage();
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
        var maxTop = vh - ph;
        var minRight = 0;
        var maxRight = vw - pw;
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

    // Create a draggable, closeable popup styled like the panel
    function createPopup(options) {
        options = options || {};
        var title = options.title || "Popup";
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

        var titleEl = document.createElement("div");
        titleEl.textContent = title;
        titleEl.style.fontWeight = "600";
        titleEl.style.textAlign = "left";
        headerBar.appendChild(titleEl);

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

            var minW = toInt(PANEL_DEFAULT_WIDTH);
            if (newW < minW) {
                newW = minW;
            }
            if (newW > PANEL_MAX_WIDTH_PX) {
                newW = PANEL_MAX_WIDTH_PX; // Enforce max width
            }

            var minH = PANEL_HEADER_HEIGHT_PX + 80;
            if (newH < minH) {
                newH = minH;
            }

            panel.style.width = String(newW) + "px";
            panel.style.height = String(newH) + "px";

            if (bodyContainer) {
                bodyContainer.style.display = "block";
                bodyContainer.style.height = "calc(100% - " + String(PANEL_HEADER_HEIGHT_PX) + "px)";
                bodyContainer.style.maxHeight = "calc(100% - " + String(PANEL_HEADER_HEIGHT_PX) + "px)";
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

    function setPanelHidden(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_HIDDEN, flag ? "1" : "0");
            log("Panel hidden state set to " + String(flag));
        } catch (e) {
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
        panel.style.borderRadius = "8px";
        panel.style.padding = "12px";
        panel.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        panel.style.fontSize = "14px";
        panel.style.minWidth = PANEL_DEFAULT_WIDTH;
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
        headerBar.style.gap = String(PANEL_HEADER_GAP_PX) + "px";
        headerBar.style.height = String(PANEL_HEADER_HEIGHT_PX) + "px";
        headerBar.style.boxSizing = "border-box";
        headerBar.style.cursor = "grab";
        headerBar.style.userSelect = "none";
        var leftSpacer = document.createElement("div");
        leftSpacer.style.width = "32px";
        var title = document.createElement("div");
        title.textContent = "ClinSpark Test Automator";
        title.style.fontWeight = "600";
        title.style.textAlign = "center";
        title.style.justifySelf = "center";
        title.style.transform = "translateX(16px)"
        headerBar.appendChild(title);
        headerBar.appendChild(leftSpacer);
        var rightControls = document.createElement("div");
        rightControls.style.display = "inline-flex";
        rightControls.style.alignItems = "center";
        rightControls.style.gap = String(PANEL_HEADER_GAP_PX) + "px";
        var collapseBtn = document.createElement("button");
        collapseBtn.textContent = getPanelCollapsed() ? "Expand" : "Collapse";
        collapseBtn.style.background = "transparent";
        collapseBtn.style.color = "#fff";
        collapseBtn.style.border = "none";
        collapseBtn.style.cursor = "pointer";
        var closeBtn = document.createElement("button");
        closeBtn.textContent = "✕";
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
        bodyContainer.style.height = "calc(100% - " + String(PANEL_HEADER_HEIGHT_PX) + "px)";
        bodyContainer.style.maxHeight = "calc(100% - " + String(PANEL_HEADER_HEIGHT_PX) + "px)";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.boxSizing = "border-box";
        var btnRow = document.createElement("div");
        btnRow.style.display = "grid";
        btnRow.style.gridTemplateColumns = "1fr 1fr";
        btnRow.style.gap = "8px";
        btnRowRef = btnRow;

        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Run Barcode";
        runBarcodeBtn.style.background = "#5b43c7";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.style.fontWeight = "500";
        runBarcodeBtn.style.transition = "background 0.2s";
        runBarcodeBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        runBarcodeBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };

        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#6c757d";
        pauseBtn.style.color = "#fff";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = "6px";
        pauseBtn.style.padding = "8px";
        pauseBtn.style.cursor = "pointer";
        pauseBtn.style.fontWeight = "500";
        pauseBtn.style.transition = "background 0.2s";
        pauseBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        pauseBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var clearLogsBtn = document.createElement("button");
        clearLogsBtn.textContent = "Clear Logs";
        clearLogsBtn.style.background = "#6c757d";
        clearLogsBtn.style.color = "#fff";
        clearLogsBtn.style.border = "none";
        clearLogsBtn.style.borderRadius = "6px";
        clearLogsBtn.style.padding = "8px";
        clearLogsBtn.style.cursor = "pointer";
        clearLogsBtn.style.fontWeight = "500";
        clearLogsBtn.style.transition = "background 0.2s";
        clearLogsBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        clearLogsBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#6c757d";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = "6px";
        toggleLogsBtn.style.padding = "8px";
        toggleLogsBtn.style.cursor = "pointer";
        toggleLogsBtn.style.fontWeight = "500";
        toggleLogsBtn.style.transition = "background 0.2s";
        toggleLogsBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var saBuilderBtn = document.createElement("button");
        saBuilderBtn.textContent = "SA Builder";
        saBuilderBtn.style.background = "#17a2b8";
        saBuilderBtn.style.color = "#fff";
        saBuilderBtn.style.border = "none";
        saBuilderBtn.style.borderRadius = "6px";
        saBuilderBtn.style.padding = "8px";
        saBuilderBtn.style.cursor = "pointer";
        saBuilderBtn.style.fontWeight = "500";
        saBuilderBtn.style.transition = "background 0.2s";
        saBuilderBtn.onmouseenter = function() { this.style.background = "#138496"; };
        saBuilderBtn.onmouseleave = function() { this.style.background = "#17a2b8"; };

        var archiveUpdateFormsBtn = document.createElement("button");
        archiveUpdateFormsBtn.textContent = "Archive/Update Forms";
        archiveUpdateFormsBtn.style.background = "#e83e8c";
        archiveUpdateFormsBtn.style.color = "#fff";
        archiveUpdateFormsBtn.style.border = "none";
        archiveUpdateFormsBtn.style.borderRadius = "6px";
        archiveUpdateFormsBtn.style.padding = "8px";
        archiveUpdateFormsBtn.style.cursor = "pointer";
        archiveUpdateFormsBtn.style.fontWeight = "500";
        archiveUpdateFormsBtn.style.transition = "background 0.2s";
        archiveUpdateFormsBtn.onmouseenter = function() { this.style.background = "#d63384"; };
        archiveUpdateFormsBtn.onmouseleave = function() { this.style.background = "#e83e8c"; };

        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);
        btnRow.appendChild(saBuilderBtn);
        btnRow.appendChild(archiveUpdateFormsBtn);

        bodyContainer.appendChild(btnRow);
        var status = document.createElement("div");
        status.style.marginTop = "10px";
        status.style.background = "#1a1a1a";
        status.style.border = "1px solid #333";
        status.style.borderRadius = "6px";
        status.style.padding = "6px";
        status.style.fontSize = "13px";
        status.style.whiteSpace = "pre-wrap";
        status.textContent = "Ready";
        bodyContainer.appendChild(status);
        var logBox = document.createElement("div");
        logBox.id = LOG_ID;
        logBox.style.marginTop = "8px";
        logBox.style.height = "220px";
        logBox.style.overflowY = "auto";
        logBox.style.background = "#141414";
        logBox.style.border = "1px solid #333";
        logBox.style.borderRadius = "6px";
        logBox.style.padding = "6px";
        logBox.style.fontSize = "12px";
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
                COLLECT_ALL_CANCELLED = true;
                SA_BUILDER_PAUSE = true;
                SA_BUILDER_CANCELLED = true;
                clearCollectAllData();
                clearEligibilityWorkingState();
                // Close SA Builder popup if open
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
        archiveUpdateFormsBtn.addEventListener("click", async function () {
            ARCHIVE_UPDATE_FORMS_CANCELLED = false;
            log("Archive/Update Forms: button clicked");
            await runArchiveUpdateForms();
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

        var runModeRaw = null;
        try {
            runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            runModeRaw = null;
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();