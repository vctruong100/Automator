
// ==UserScript==
// @name ClinSpark Start Automator
// @namespace vinh.activity.plan.state
// @version 1.1.0
// @description
// @match https://cenexeltest.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Start%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Start%20Automator.js
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// @ts-check
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

    // Add Existing Subject feature storage keys
    var STORAGE_ADD_EXISTING_SUBJECT_POPUP = "activityPlanState.addExistingSubject.popup";
    var STORAGE_ADD_EXISTING_SUBJECT_MODE = "activityPlanState.addExistingSubject.mode";
    var STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH = "activityPlanState.addExistingSubject.selectedEpoch";
    var STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_URL = "activityPlanState.addExistingSubject.selectedEpochUrl";
    var STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA = "activityPlanState.addExistingSubject.collectedData";
    var STORAGE_ADD_EXISTING_SUBJECT_SELECTED_SUBJECT = "activityPlanState.addExistingSubject.selectedSubject";
    var STORAGE_ADD_EXISTING_SUBJECT_EDIT_DONE = "activityPlanState.addExistingSubject.editDone";
    var STORAGE_ADD_EXISTING_SUBJECT_ADD_DONE = "activityPlanState.addExistingSubject.addDone";
    var ADD_EXISTING_SUBJECT_POPUP_REF = null;
    var ADD_EXISTING_SUBJECT_INTENTIONAL_CLOSE = false;
    // Temp page coordination keys
    var STORAGE_ADD_EXISTING_SUBJECT_SCAN_QUEUE = "activityPlanState.addExistingSubject.scanQueue";
    var STORAGE_ADD_EXISTING_SUBJECT_SCAN_COMPLETE = "activityPlanState.addExistingSubject.scanComplete";
    var STORAGE_ADD_EXISTING_SUBJECT_IS_TEMP_PAGE = "activityPlanState.addExistingSubject.isTempPage";
    var STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH = "activityPlanState.addExistingSubject.currentEpoch";
    var STORAGE_ADD_EXISTING_SUBJECT_CURRENT_COHORT = "activityPlanState.addExistingSubject.currentCohort";
    var STORAGE_ADD_EXISTING_SUBJECT_ALL_EPOCHS = "activityPlanState.addExistingSubject.allEpochs";
    var STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_SUBJECTS = "activityPlanState.addExistingSubject.selectedEpochSubjects";

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

    var STORAGE_FIND_FORM_PENDING = "activityPlanState.findForm.pending";
    var STORAGE_FIND_FORM_KEYWORD = "activityPlanState.findForm.keyword";
    var STORAGE_FIND_FORM_SUBJECT = "activityPlanState.findForm.subject";
    var STORAGE_FIND_FORM_STATUS_VALUES = "activityPlanState.findForm.statusValues";


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
                log("processLockSamplePathsPage: continuing to Study Show for ALL mode");
                updateRunAllPopupStatus("Running Update Study Status");
                location.href = STUDY_SHOW_URL + "?autoupdate=1";
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
            log("processLockSamplePathsPage: continuing to Study Show for ALL mode");
            updateRunAllPopupStatus("Running Update Study Status");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
        }
    }

    async function processLockSamplePathDetailPage() {
        log("processLockSamplePathDetailPage: not needed with background approach");
    }

    async function processLockSamplePathUpdatePage() {
        log("processLockSamplePathUpdatePage: not needed with background approach");
    }

    
    //==========================
    // CLEAR SUBJECT ELIGIBILITY FEATURE
    //==========================
    // This section contains all functions related to clearing subject eligibility.
    // This feature automates clearing all existing eligibility mapping in the table.
    //==========================

    function ClearEligibilityFunctions() {}

    function startClearMapping() {
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

        log("ClearMapping: clicking delete link");
        deleteLink.click();
        await sleep(600);

        var okBtn = null;
        var waited = 0;
        var step = 150;
        var maxWait = 4000;

        while (waited < maxWait) {
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
            title: "Select Epoch for Import",
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
    // ADD EXISTING SUBJECT FEATURE
    //==========================
    // This section contains all functions related to adding an existing subject from another epoch.
    // This feature collects subject data from all epochs (except the selected one), displays
    // a selection popup, and automates adding the selected subject to the target epoch's first cohort.
    //==========================

    function AddExistingSubjectFunctions() {}

    function clearAddExistingSubjectState() {
        try {
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_POPUP);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_URL);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_SUBJECT);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_EDIT_DONE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_ADD_DONE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_QUEUE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_COMPLETE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_IS_TEMP_PAGE);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_COHORT);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_ALL_EPOCHS);
            localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_SUBJECTS);
            localStorage.removeItem("addExisting.isSelectedEpochCohort");
        } catch (e) {}
        ADD_EXISTING_SUBJECT_POPUP_REF = null;
        ADD_EXISTING_SUBJECT_INTENTIONAL_CLOSE = false;
        log("AddExistingSubject: state cleared");
    }

    function isTempPageForAddExisting() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_IS_TEMP_PAGE);
        } catch (e) {}
        return raw === "1";
    }

    function getScanQueue() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_QUEUE);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        try {
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function setScanQueue(queue) {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_QUEUE, JSON.stringify(queue));
        } catch (e) {}
    }

    function appendCollectedSubjects(subjects) {
        var existing = getAddExistingSubjectCollectedData();
        var combined = existing.concat(subjects);
        setAddExistingSubjectCollectedData(combined);
        log("AddExistingSubject: Appended " + String(subjects.length) + " subjects, total=" + String(combined.length));
    }

    function markScanComplete() {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_COMPLETE, "1");
        } catch (e) {}
    }

    function isScanComplete() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_COMPLETE);
        } catch (e) {}
        return raw === "1";
    }

    function getSelectedEpochSubjects() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_SUBJECTS);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        try {
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function setSelectedEpochSubjects(subjects) {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_SUBJECTS, JSON.stringify(subjects));
        } catch (e) {}
    }

    function getAddExistingSubjectCollectedData() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA);
        } catch (e) {}
        if (!raw) {
            return [];
        }
        try {
            return JSON.parse(raw);
        } catch (e) {
            return [];
        }
    }

    function setAddExistingSubjectCollectedData(data) {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA, JSON.stringify(data));
        } catch (e) {}
    }

    function getAddExistingSubjectSelectedSubject() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_SUBJECT);
        } catch (e) {}
        if (!raw) {
            return null;
        }
        try {
            return JSON.parse(raw);
        } catch (e) {
            return null;
        }
    }

    function setAddExistingSubjectSelectedSubject(subject) {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_SUBJECT, JSON.stringify(subject));
        } catch (e) {}
    }

    function isAddExistingSubjectEditDone() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_EDIT_DONE);
        } catch (e) {}
        return raw === "1";
    }

    function setAddExistingSubjectEditDone() {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_EDIT_DONE, "1");
        } catch (e) {}
    }

    function isAddExistingSubjectAddDone() {
        var raw = null;
        try {
            raw = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_ADD_DONE);
        } catch (e) {}
        return raw === "1";
    }

    function setAddExistingSubjectAddDone() {
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_ADD_DONE, "1");
        } catch (e) {}
    }

    // Collect subjects from the current page's DOM (study subjects list or cohort pages)
    function collectSubjectsFromCurrentPage() {
        log("AddExistingSubject: Collecting subjects from current page DOM");
        var subjects = [];
        
        // Look for the cohort assignment table on the current page
        var tbody = document.querySelector('tbody#cohortAssignmentListBody');
        if (tbody) {
            log("AddExistingSubject: Found cohortAssignmentListBody on current page");
            var rows = tbody.querySelectorAll('tr');
            log("AddExistingSubject: Found " + String(rows.length) + " rows in cohortAssignmentListBody");
            
            var rowIdx = 0;
            while (rowIdx < rows.length) {
                var row = rows[rowIdx];
                
                // Extract subject number
                var subjectSpan = row.querySelector('span.tooltips[data-original-title="Subject Number"]');
                var subjectNumber = subjectSpan ? (subjectSpan.textContent + "").trim() : "";
                
                // Extract volunteer info and ID
                var volunteerLink = row.querySelector('a[href^="/secure/volunteers/manage/show/"]');
                var volunteerInfo = "";
                var volunteerId = "";
                if (volunteerLink) {
                    volunteerInfo = (volunteerLink.textContent + "").trim().replace(/\s+/g, " ");
                    var volunteerHref = volunteerLink.getAttribute("href");
                    if (volunteerHref) {
                        var volIdMatch = volunteerHref.match(/\/show\/(\d+)/);
                        if (volIdMatch) {
                            volunteerId = volIdMatch[1];
                        }
                    }
                }
                
                if (subjectNumber && subjectNumber.length > 0) {
                    subjects.push({
                        subjectNumber: subjectNumber,
                        volunteerInfo: volunteerInfo,
                        volunteerId: volunteerId,
                        epochName: "",
                        cohortName: ""
                    });
                    log("AddExistingSubject: Collected subject " + subjectNumber + " (volunteerId: " + volunteerId + ", volunteer: " + volunteerInfo + ")");
                }
                
                rowIdx = rowIdx + 1;
            }
        } else {
            log("AddExistingSubject: cohortAssignmentListBody not found on current page");
        }
        
        return subjects;
    }

    function showSubjectSelectionPopup(subjects, selectedEpochUrl) {
        if (subjects.length === 0) {
            var noSubjectsContent = document.createElement("div");
            noSubjectsContent.style.textAlign = "center";
            noSubjectsContent.style.padding = "20px";
            noSubjectsContent.style.color = "#fff";
            noSubjectsContent.textContent = "No subjects found in other epochs.";

            if (ADD_EXISTING_SUBJECT_POPUP_REF && document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
                var popupContent = ADD_EXISTING_SUBJECT_POPUP_REF.element.querySelector('.popup-content');
                if (popupContent) {
                    popupContent.innerHTML = "";
                    popupContent.appendChild(noSubjectsContent);
                }
            } else {
                ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                    title: "Add Existing Subject",
                    content: noSubjectsContent,
                    width: "400px",
                    height: "auto",
                    onClose: function() {
                        clearAddExistingSubjectState();
                    }
                });
            }
            return;
        }

        var contentDiv = document.createElement("div");
        contentDiv.style.display = "flex";
        contentDiv.style.flexDirection = "column";
        contentDiv.style.gap = "8px";
        contentDiv.style.maxHeight = "400px";
        contentDiv.style.overflowY = "auto";

        var i = 0;
        while (i < subjects.length) {
            var subject = subjects[i];

            var itemDiv = document.createElement("div");
            itemDiv.style.display = "flex";
            itemDiv.style.flexDirection = "column";
            itemDiv.style.gap = "4px";
            itemDiv.style.padding = "10px";
            itemDiv.style.background = "#222";
            itemDiv.style.borderRadius = "6px";
            itemDiv.style.border = "1px solid #444";

            var topRow = document.createElement("div");
            topRow.style.display = "flex";
            topRow.style.justifyContent = "space-between";
            topRow.style.alignItems = "center";

            var subjectInfo = document.createElement("div");
            subjectInfo.style.display = "flex";
            subjectInfo.style.flexDirection = "column";
            subjectInfo.style.gap = "2px";

            var subjectNumSpan = document.createElement("span");
            subjectNumSpan.style.fontWeight = "600";
            subjectNumSpan.style.fontSize = "14px";
            subjectNumSpan.style.color = "#fff";
            subjectNumSpan.textContent = subject.subjectNumber;
            subjectInfo.appendChild(subjectNumSpan);

            if (subject.volunteerInfo && subject.volunteerInfo.length > 0) {
                var volunteerSpan = document.createElement("span");
                volunteerSpan.style.fontSize = "12px";
                volunteerSpan.style.color = "#aaa";
                volunteerSpan.textContent = subject.volunteerInfo;
                subjectInfo.appendChild(volunteerSpan);
            }

            var locationSpan = document.createElement("span");
            locationSpan.style.fontSize = "11px";
            locationSpan.style.color = "#888";
            locationSpan.textContent = subject.epochName + " → " + subject.cohortName;
            subjectInfo.appendChild(locationSpan);

            topRow.appendChild(subjectInfo);

            var selectBtn = document.createElement("button");
            selectBtn.textContent = "Select";
            selectBtn.style.padding = "6px 12px";
            selectBtn.style.cursor = "pointer";
            selectBtn.style.background = "#4a90e2";
            selectBtn.style.color = "#fff";
            selectBtn.style.border = "none";
            selectBtn.style.borderRadius = "4px";
            selectBtn.style.fontSize = "12px";
            selectBtn.style.fontWeight = "500";
            selectBtn.style.transition = "background 0.2s";

            selectBtn.addEventListener("mouseenter", function() {
                this.style.background = "#357abd";
            });
            selectBtn.addEventListener("mouseleave", function() {
                this.style.background = "#4a90e2";
            });

(function(subj, popupContentDiv) {
                selectBtn.addEventListener("click", function() {
                    log("AddExistingSubject: Selected subject " + subj.subjectNumber);
                    setAddExistingSubjectSelectedSubject(subj);
                    setAddExistingSubjectCollectedData([]);

                    try {
                        localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_MODE, "addSubject");
                        localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_POPUP, "1");
                    } catch (e) {}

                    popupContentDiv.innerHTML = "";
                    var runningDiv = document.createElement("div");
                    runningDiv.style.textAlign = "center";
                    runningDiv.style.padding = "40px 20px";
                    runningDiv.style.fontSize = "16px";
                    runningDiv.style.color = "#fff";
                    runningDiv.id = "addExistingRunningStatus";
                    runningDiv.textContent = "Running.";
                    popupContentDiv.appendChild(runningDiv);

                    var dots = 1;
                    var runningInterval = setInterval(function() {
                        var statusEl = document.getElementById("addExistingRunningStatus");
                        if (!statusEl) {
                            clearInterval(runningInterval);
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
                        statusEl.textContent = text;
                    }, 500);

                    var studyId = null;
                    var studyShowMatch = location.pathname.match(/\/secure\/administration\/studies\/show\/(\d+)/);
                    if (studyShowMatch) {
                        studyId = studyShowMatch[1];
                    }

                    if (studyId) {
                        log("AddExistingSubject: Navigating to study show page for study " + studyId);
                        setTimeout(function() {
                            location.href = location.origin + "/secure/administration/studies/show/" + studyId + "?autoaddexisting=1";
                        }, 500);
                    } else {
                        log("AddExistingSubject: Could not determine study ID from current URL");
                    }
                });
            })(subject, contentDiv);

            topRow.appendChild(selectBtn);
            itemDiv.appendChild(topRow);
            contentDiv.appendChild(itemDiv);

            i = i + 1;
        }

if (ADD_EXISTING_SUBJECT_POPUP_REF && document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
            var popupContent = ADD_EXISTING_SUBJECT_POPUP_REF.element.querySelector('.popup-content');
            if (popupContent) {
                popupContent.innerHTML = "";
                popupContent.appendChild(contentDiv);
            }
            var popupTitle = ADD_EXISTING_SUBJECT_POPUP_REF.element.querySelector('.popup-title');
            if (popupTitle) {
                popupTitle.textContent = "Select Subject to Add (" + String(subjects.length) + " found)";
            }
        } else {
            ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                title: "Select Subject to Add (" + String(subjects.length) + " found)",
                content: contentDiv,
                width: "450px",
                height: "auto",
                maxHeight: "500px",
                onClose: function() {
                    log("AddExistingSubject: User closed popup, stopping and clearing state");
                    clearAddExistingSubjectState();
                }
            });
        }
    }

    async function processStudyShowPageForAddExistingSubject() {
        var autoAddExisting = getQueryParam("autoaddexisting");
        var mode = null;
        try {
            mode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
        } catch (e) {}

        // Handle mode "addSubject" - navigate to selected epoch
        if (mode === "addSubject" && autoAddExisting === "1") {
            var selectedEpochUrl = null;
            try {
                selectedEpochUrl = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_URL);
            } catch (e) {}

            if (!selectedEpochUrl) {
                log("AddExistingSubject: No selected epoch URL found");
                clearAddExistingSubjectState();
                return;
            }

            log("AddExistingSubject: Navigating to selected epoch: " + selectedEpochUrl);
            location.href = location.origin + selectedEpochUrl + "?autoaddexisting=1";
            return;
        }

        // Handle mode "selectEpoch" - show epoch selection popup
        if (!(mode === "selectEpoch" && getQueryParam("autoaddexistingsubject") === "1")) {
            return;
        }

        var tbody = await waitForSelector('tbody#epochTableBody', 5000);
        if (!tbody) {
            log("AddExistingSubject: Epoch table not found");
            return;
        }

        var anchors = tbody.querySelectorAll('a[href^="/secure/administration/studies/epoch/show/"]');
        if (anchors.length === 0) {
            log("AddExistingSubject: No epochs found");
            return;
        }

        var allEpochs = [];
        var i = 0;
        while (i < anchors.length) {
            var a = anchors[i];
            var name = (a.textContent + "").trim();
            var href = a.getAttribute("href");
            allEpochs.push({
                name: name,
                url: location.origin + href,
                href: href,
                isScreening: isScreeningLabel(name)
            });
            i = i + 1;
        }

        // Store all epochs for later use
        try {
            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_ALL_EPOCHS, JSON.stringify(allEpochs));
        } catch (e) {}

        var contentDiv = document.createElement("div");
        contentDiv.style.display = "flex";
        contentDiv.style.flexDirection = "column";
        contentDiv.style.gap = "8px";

        var instructionDiv = document.createElement("div");
        instructionDiv.style.color = "#aaa";
        instructionDiv.style.fontSize = "12px";
        instructionDiv.style.marginBottom = "8px";
        instructionDiv.textContent = "Select the epoch where you want to add the existing subject. Screening epochs are excluded.";
        contentDiv.appendChild(instructionDiv);

        var popup = null;

        var epochIndex = 0;
        while (epochIndex < allEpochs.length) {
            var epochInfo = allEpochs[epochIndex];

            if (epochInfo.isScreening) {
                epochIndex = epochIndex + 1;
                continue;
            }

            var btn = document.createElement("button");
            btn.textContent = epochInfo.name;
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
                this.style.background = "#357abd";
            });
            btn.addEventListener("mouseleave", function() {
                this.style.background = "#4a90e2";
            });

            (function(selectedEpoch, allEpochsList) {
                btn.addEventListener("click", async function() {
                    log("AddExistingSubject: Selected target epoch: " + selectedEpoch.name);

                    try {
                        localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH, selectedEpoch.name);
                        localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_URL, selectedEpoch.href);
                        // Clear previous collected data
                        localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA);
                        localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH_SUBJECTS);
                        localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SCAN_COMPLETE);
                    } catch (e) {}

                    ADD_EXISTING_SUBJECT_INTENTIONAL_CLOSE = true;
                    if (popup) {
                        popup.close();
                    }

                    // Create progress popup
                    var progressContainer = document.createElement("div");
                    progressContainer.style.display = "flex";
                    progressContainer.style.flexDirection = "column";
                    progressContainer.style.gap = "16px";
                    progressContainer.style.padding = "8px";

                    var statusDiv = document.createElement("div");
                    statusDiv.id = "addExistingScanStatus";
                    statusDiv.style.textAlign = "center";
                    statusDiv.style.fontSize = "14px";
                    statusDiv.style.color = "#fff";
                    statusDiv.textContent = "Opening temporary pages to scan epochs...";

                    var progressDiv = document.createElement("div");
                    progressDiv.id = "addExistingScanProgress";
                    progressDiv.style.textAlign = "center";
                    progressDiv.style.fontSize = "12px";
                    progressDiv.style.color = "#9df";
                    progressDiv.textContent = "Preparing scan queue...";

                    var loadingDiv = document.createElement("div");
                    loadingDiv.style.textAlign = "center";
                    loadingDiv.style.fontSize = "14px";
                    loadingDiv.style.color = "#9df";
                    loadingDiv.textContent = "Running.";

                    progressContainer.appendChild(statusDiv);
                    progressContainer.appendChild(progressDiv);
                    progressContainer.appendChild(loadingDiv);

                    ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                        title: "Scanning Epochs",
                        content: progressContainer,
                        width: "400px",
                        height: "auto",
                        onClose: function() {
                            clearAddExistingSubjectState();
                        }
                    });

                    var dots = 1;
                    var loadingInterval = setInterval(function() {
                        if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
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
                        loadingDiv.textContent = text;
                    }, 500);

                    // Build scan queue - one entry per epoch (we'll discover cohorts when we open the epoch page)
                    var scanQueue = [];
                    var scanIdx = 0;
                    while (scanIdx < allEpochsList.length) {
                        var ep = allEpochsList[scanIdx];
                        scanQueue.push({
                            type: "epoch",
                            epochName: ep.name,
                            epochUrl: ep.url,
                            epochHref: ep.href,
                            isSelectedEpoch: ep.name === selectedEpoch.name
                        });
                        scanIdx = scanIdx + 1;
                    }

                    log("AddExistingSubject: Created scan queue with " + String(scanQueue.length) + " epochs");
                    setScanQueue(scanQueue);

                    // Start scanning by opening the first epoch in a new tab
                    if (scanQueue.length > 0) {
                        var firstItem = scanQueue[0];
                        // Remove first item from queue
                        scanQueue.shift();
                        setScanQueue(scanQueue);

                        try {
                            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_IS_TEMP_PAGE, "1");
                            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH, firstItem.epochName);
                            localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_MODE, "scanning");
                        } catch (e) {}

                        log("AddExistingSubject: Opening first epoch for scanning: " + firstItem.epochName);
                        var scanUrl = firstItem.epochUrl + "?addexistingscan=1";
                        
                        // Open in a new tab using GM.openInTab
                        if (typeof GM !== "undefined" && typeof GM.openInTab === "function") {
                            GM.openInTab(scanUrl, { active: false, insert: true, setParent: true });
                        } else if (typeof GM_openInTab === "function") {
                            GM_openInTab(scanUrl, { active: false, insert: true, setParent: true });
                        } else {
                            window.open(scanUrl, "_blank");
                        }

                        // Monitor for scan completion
                        var checkInterval = setInterval(function() {
                            if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
                                clearInterval(checkInterval);
                                clearInterval(loadingInterval);
                                return;
                            }

                            var queue = getScanQueue();
                            var complete = isScanComplete();
                            var collected = getAddExistingSubjectCollectedData();
                            var selectedEpochSubjs = getSelectedEpochSubjects();

                            // Update progress
                            var pDiv = document.getElementById("addExistingScanProgress");
                            if (pDiv) {
                                pDiv.textContent = "Collected " + String(collected.length) + " subjects, " + String(queue.length) + " items remaining";
                            }

                            if (complete) {
                                clearInterval(checkInterval);
                                clearInterval(loadingInterval);
                                log("AddExistingSubject: Scan complete, collected " + String(collected.length) + " subjects");

                                // Filter out subjects that exist in the selected epoch
                                var filteredSubjects = [];
                                var volunteerIdsInSelectedEpoch = {};
                                var selIdx = 0;
                                while (selIdx < selectedEpochSubjs.length) {
                                    var selSubj = selectedEpochSubjs[selIdx];
                                    if (selSubj.volunteerId) {
                                        volunteerIdsInSelectedEpoch[selSubj.volunteerId] = true;
                                    }
                                    selIdx = selIdx + 1;
                                }

                                var colIdx = 0;
                                while (colIdx < collected.length) {
                                    var subj = collected[colIdx];
                                    // Only include if not in selected epoch
                                    if (!volunteerIdsInSelectedEpoch[subj.volunteerId]) {
                                        filteredSubjects.push(subj);
                                    }
                                    colIdx = colIdx + 1;
                                }

                                log("AddExistingSubject: Filtered to " + String(filteredSubjects.length) + " subjects (excluded " + String(collected.length - filteredSubjects.length) + " from selected epoch)");

                                // Clear temp page flag
                                try {
                                    localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_IS_TEMP_PAGE);
                                } catch (e) {}

                                // Don't close popup, just update its content
                                showSubjectSelectionPopup(filteredSubjects, selectedEpoch.href);
                            }
                        }, 1000);
                    } else {
                        clearInterval(loadingInterval);
                        log("AddExistingSubject: No epochs to scan");
                        if (ADD_EXISTING_SUBJECT_POPUP_REF) {
                            ADD_EXISTING_SUBJECT_POPUP_REF.close();
                        }
                        clearAddExistingSubjectState();
                    }
                });
            })(epochInfo, allEpochs);

            contentDiv.appendChild(btn);
            epochIndex = epochIndex + 1;
        }

        popup = createPopup({
            title: "Select Target Epoch for Add Existing Subject",
            content: contentDiv,
            width: "400px",
            height: "auto",
            maxHeight: "80%",
            onClose: function() {
                if (ADD_EXISTING_SUBJECT_INTENTIONAL_CLOSE) {
                    ADD_EXISTING_SUBJECT_INTENTIONAL_CLOSE = false;
                    return;
                }
                clearAddExistingSubjectState();
            }
        });
    }

    // Process epoch page when opened as a temp scanning page
    async function processEpochShowPageForAddExistingScan() {
        var scanParam = getQueryParam("addexistingscan");
        var mode = null;
        try {
            mode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
        } catch (e) {}

        if (!(mode === "scanning" && scanParam === "1")) {
            return;
        }

        var currentEpoch = null;
        try {
            currentEpoch = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH);
        } catch (e) {}

        log("AddExistingSubject TempPage: Processing epoch " + String(currentEpoch));

        // Wait for cohort list to load
        var cohortAnchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        var waitedCohorts = 0;
        while (cohortAnchors.length === 0 && waitedCohorts < 5000) {
            await sleep(300);
            waitedCohorts = waitedCohorts + 300;
            cohortAnchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        }

        if (cohortAnchors.length === 0) {
            log("AddExistingSubject TempPage: No cohorts found in epoch " + String(currentEpoch));
            // Continue to next item in queue
            continueToNextScanItem();
            return;
        }

        // Add cohorts to scan queue
        var queue = getScanQueue();
        var selectedEpochName = null;
        try {
            selectedEpochName = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_EPOCH);
        } catch (e) {}
        var isSelectedEpoch = currentEpoch === selectedEpochName;

        var cohortIdx = 0;
        while (cohortIdx < cohortAnchors.length) {
            var cohortAnchor = cohortAnchors[cohortIdx];
            var cohortName = (cohortAnchor.textContent + "").trim();
            var cohortHref = cohortAnchor.getAttribute("href");
            queue.unshift({
                type: "cohort",
                epochName: currentEpoch,
                cohortName: cohortName,
                cohortUrl: location.origin + cohortHref,
                cohortHref: cohortHref,
                isSelectedEpoch: isSelectedEpoch
            });
            cohortIdx = cohortIdx + 1;
        }

        log("AddExistingSubject TempPage: Added " + String(cohortAnchors.length) + " cohorts to queue");
        setScanQueue(queue);

        // Continue to next item (which will be the first cohort we just added)
        continueToNextScanItem();
    }

    // Process cohort page when opened as a temp scanning page
    async function processCohortShowPageForAddExistingScan() {
        var scanParam = getQueryParam("addexistingscan");
        var mode = null;
        try {
            mode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
        } catch (e) {}

        if (!(mode === "scanning" && scanParam === "1")) {
            return;
        }

        var currentCohort = null;
        var currentEpoch = null;
        try {
            currentCohort = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_COHORT);
            currentEpoch = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH);
        } catch (e) {}

        log("AddExistingSubject TempPage: Processing cohort " + String(currentCohort) + " in epoch " + String(currentEpoch));

        // Wait for cohortAssignmentListBody to load
        var tbody = await waitForSelector('tbody#cohortAssignmentListBody', 8000);
        if (!tbody) {
            log("AddExistingSubject TempPage: cohortAssignmentListBody not found");
            continueToNextScanItem();
            return;
        }

        // Wait for rows to appear (DataTables might load async)
        await sleep(1500);

        var rows = tbody.querySelectorAll('tr.cohortAssignmentRow');
        log("AddExistingSubject TempPage: Found " + String(rows.length) + " assignment rows");

        var subjects = [];
        var rowIdx = 0;
        while (rowIdx < rows.length) {
            var row = rows[rowIdx];

            // Extract subject number from 2nd td
            var tds = row.querySelectorAll('td');
            var subjectNumber = "";
            var volunteerId = "";
            var volunteerInfo = "";

            if (tds.length >= 2) {
                // 2nd td contains subject number
                var subjectSpan = tds[1].querySelector('span.tooltips[data-original-title="Subject Number"]');
                if (subjectSpan) {
                    subjectNumber = (subjectSpan.textContent + "").trim();
                }
            }

            if (tds.length >= 5) {
                // 5th td contains volunteer info
                var volunteerLink = tds[4].querySelector('a[href^="/secure/volunteers/manage/show/"]');
                if (volunteerLink) {
                    volunteerInfo = (volunteerLink.textContent + "").trim().replace(/\s+/g, " ");
                    var volHref = volunteerLink.getAttribute("href");
                    if (volHref) {
                        var volMatch = volHref.match(/\/show\/(\d+)/);
                        if (volMatch) {
                            volunteerId = volMatch[1];
                        }
                    }
                }
            }

            if (subjectNumber && subjectNumber.length > 0) {
                subjects.push({
                    subjectNumber: subjectNumber,
                    volunteerInfo: volunteerInfo,
                    volunteerId: volunteerId,
                    epochName: currentEpoch,
                    cohortName: currentCohort
                });
                log("AddExistingSubject TempPage: Collected " + subjectNumber + " (volunteerId: " + volunteerId + ")");
            }

            rowIdx = rowIdx + 1;
        }

        // Check if this is from the selected epoch using flag
        var isSelectedEpochCohort = null;
        try {
            isSelectedEpochCohort = localStorage.getItem("addExisting.isSelectedEpochCohort");
        } catch (e) {}

        if (isSelectedEpochCohort === "1") {
            // Store as selected epoch subjects (for filtering later)
            var existingSelectedSubjs = getSelectedEpochSubjects();
            var combined = existingSelectedSubjs.concat(subjects);
            setSelectedEpochSubjects(combined);
            log("AddExistingSubject TempPage: Added " + String(subjects.length) + " to selected epoch subjects (will be excluded from selection)");
        } else {
            // Append to collected data
            appendCollectedSubjects(subjects);
        }

        // Continue to next item
        continueToNextScanItem();
    }

    function continueToNextScanItem() {
        var queue = getScanQueue();

        if (queue.length === 0) {
            log("AddExistingSubject TempPage: Scan queue empty, marking complete");
            markScanComplete();
            // Close this temp tab
            window.close();
            return;
        }

        var nextItem = queue.shift();
        setScanQueue(queue);

        log("AddExistingSubject TempPage: Processing next item: " + nextItem.type + " - " + (nextItem.cohortName || nextItem.epochName) + " (isSelectedEpoch=" + String(nextItem.isSelectedEpoch) + ")");

        try {
            if (nextItem.type === "epoch") {
                localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH, nextItem.epochName);
                localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_COHORT);
                // Track if this is the selected epoch
                if (nextItem.isSelectedEpoch) {
                    localStorage.setItem("addExisting.isSelectedEpochCohort", "1");
                } else {
                    localStorage.removeItem("addExisting.isSelectedEpochCohort");
                }
            } else if (nextItem.type === "cohort") {
                localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_EPOCH, nextItem.epochName);
                localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_CURRENT_COHORT, nextItem.cohortName);
                // Track if this cohort belongs to the selected epoch
                if (nextItem.isSelectedEpoch) {
                    localStorage.setItem("addExisting.isSelectedEpochCohort", "1");
                } else {
                    localStorage.removeItem("addExisting.isSelectedEpochCohort");
                }
            }
        } catch (e) {}

        // Navigate to the next URL
        var nextUrl = "";
        if (nextItem.type === "epoch") {
            nextUrl = nextItem.epochUrl + "?addexistingscan=1";
        } else if (nextItem.type === "cohort") {
            nextUrl = nextItem.cohortUrl + "?addexistingscan=1";
        }

        location.href = nextUrl;
    }

async function processEpochShowPageForAddExistingSubject() {
        var autoAddExisting = getQueryParam("autoaddexisting");
        var mode = null;
        try {
            mode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
        } catch (e) {}

        if (!(mode === "addSubject" && autoAddExisting === "1")) {
            return;
        }

        var selectedSubject = getAddExistingSubjectSelectedSubject();
        if (!selectedSubject) {
            log("AddExistingSubject: No selected subject found");
            clearAddExistingSubjectState();
            return;
        }

        log("AddExistingSubject: Processing epoch for subject " + selectedSubject.subjectNumber);

        if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
            var runningContent = document.createElement("div");
            runningContent.style.textAlign = "center";
            runningContent.style.padding = "40px 20px";
            runningContent.style.fontSize = "16px";
            runningContent.style.color = "#fff";
            runningContent.id = "addExistingRunningStatus";
            runningContent.textContent = "Running.";

            ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                title: "Add Existing Subject",
                content: runningContent,
                width: "450px",
                height: "auto",
                onClose: function() {
                    log("AddExistingSubject: User closed popup, stopping and clearing state");
                    clearAddExistingSubjectState();
                }
            });

            var dots = 1;
            var runningInterval = setInterval(function() {
                var statusEl = document.getElementById("addExistingRunningStatus");
                if (!statusEl) {
                    clearInterval(runningInterval);
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
                statusEl.textContent = text;
            }, 500);
        }

        var cohortAnchors = document.querySelectorAll('a[href^="/secure/administration/studies/cohort/show/"]');
        if (cohortAnchors.length === 0) {
            log("AddExistingSubject: No cohorts found in epoch");
            if (ADD_EXISTING_SUBJECT_POPUP_REF) {
                ADD_EXISTING_SUBJECT_POPUP_REF.close();
            }
            ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                title: "Error",
                content: "<div style='text-align:center;padding:20px;color:#f99;'>No Cohort Found in this epoch.</div>",
                width: "300px",
                height: "auto",
                onClose: function() {
                    clearAddExistingSubjectState();
                }
            });
            return;
        }

        var firstCohortHref = cohortAnchors[0].getAttribute("href");
        log("AddExistingSubject: Navigating to first cohort: " + firstCohortHref);
        location.href = location.origin + firstCohortHref + "?autoaddexistingcohort=1";
    }

async function processCohortShowPageForAddExistingSubject() {
        var autoAddExisting = getQueryParam("autoaddexistingcohort");
        var mode = null;
        try {
            mode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
        } catch (e) {}

        if (!(mode === "addSubject" && autoAddExisting === "1")) {
            return;
        }

        var selectedSubject = getAddExistingSubjectSelectedSubject();
        if (!selectedSubject) {
            log("AddExistingSubject: No selected subject found");
            clearAddExistingSubjectState();
            return;
        }

        log("AddExistingSubject: Processing cohort for subject " + selectedSubject.subjectNumber);

        if (!ADD_EXISTING_SUBJECT_POPUP_REF || !document.body.contains(ADD_EXISTING_SUBJECT_POPUP_REF.element)) {
            var runningContent = document.createElement("div");
            runningContent.style.textAlign = "center";
            runningContent.style.padding = "40px 20px";
            runningContent.style.fontSize = "16px";
            runningContent.style.color = "#fff";
            runningContent.id = "addExistingRunningStatus";
            runningContent.textContent = "Running.";

            ADD_EXISTING_SUBJECT_POPUP_REF = createPopup({
                title: "Add Existing Subject",
                content: runningContent,
                width: "450px",
                height: "auto",
                onClose: function() {
                    log("AddExistingSubject: User closed popup, stopping and clearing state");
                    clearAddExistingSubjectState();
                }
            });

            var dots = 1;
            var runningInterval = setInterval(function() {
                var statusEl = document.getElementById("addExistingRunningStatus");
                if (!statusEl) {
                    clearInterval(runningInterval);
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
                statusEl.textContent = text;
            }, 500);
        }

        // Check if add step is done (set before page reload)
        var addDone = isAddExistingSubjectAddDone();
        var successAlert = hasSuccessAlert();
        if (addDone || successAlert) {
            log("AddExistingSubject: Add step complete (addDone=" + String(addDone) + ", successAlert=" + String(successAlert) + "), proceeding to activation");
            await performAddExistingSubjectActivation(selectedSubject);
            return;
        }

        var editDone = isAddExistingSubjectEditDone();
        if (!editDone) {
            log("AddExistingSubject: Edit not done, opening Edit modal");

            var opened = await clickActionsDropdownIfNeeded();
            if (!opened) {
                log("AddExistingSubject: Actions dropdown not found");
                return;
            }

            var editLink = document.querySelector('a[href^="/secure/administration/studies/cohort/update/"][data-toggle="modal"]');
            if (!editLink) {
                log("AddExistingSubject: Edit link not found");
                return;
            }

            editLink.click();
            log("AddExistingSubject: Edit link clicked");

            var modal = await waitForSelector("#ajaxModal, .modal", 6000);
            if (!modal) {
                log("AddExistingSubject: Modal did not open");
                return;
            }

            await sleep(1500);

            log("AddExistingSubject: Applying checkbox settings for existing subject");
            var ok6 = true;
            ok6 = ok6 && setCheckboxStateById("subjectInitiation", false);
            ok6 = ok6 && setCheckboxStateById("sourceVolunteerDatabase", false);
            ok6 = ok6 && setCheckboxStateById("sourceAppointments", false);
            ok6 = ok6 && setCheckboxStateById("sourceAppointmentsCohort", false);
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
            if (reason) {
                reason.value = "test";
                var evtR = new Event("input", { bubbles: true });
                reason.dispatchEvent(evtR);
            }

            setAddExistingSubjectEditDone();

            var saveBtn = await waitForSelector("#actionButton", 5000);
            if (saveBtn) {
                saveBtn.click();
                log("AddExistingSubject: Edit modal Save clicked");
                await sleep(2000);
            }
            return;
        }

        log("AddExistingSubject: Edit done, proceeding to Add");
        await performAddExistingSubjectAdd(selectedSubject);
    }

    async function performAddExistingSubjectAdd(selectedSubject) {
        var opened = await clickActionsDropdownIfNeeded();
        if (!opened) {
            log("AddExistingSubject: Actions dropdown not found for Add");
            return;
        }

        var addLink = await waitForSelector('a#addCohortAssignmentButton[data-toggle="modal"]', 3000);
        if (!addLink) {
            addLink = document.querySelector('a[href^="/secure/study/cohortassign/manage/save/"][data-toggle="modal"]');
        }
        if (!addLink) {
            log("AddExistingSubject: Add link not found");
            return;
        }

        addLink.click();
        log("AddExistingSubject: Add link clicked");

        var modal = await waitForSelector("#ajaxModal, .modal", 5000);
        if (!modal) {
            log("AddExistingSubject: Add modal did not open");
            return;
        }

        var planSel = await waitForSelector('select#activityPlan', 5000);
        if (!planSel) {
            log("AddExistingSubject: ActivityPlan select not found");
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
            log("AddExistingSubject: ActivityPlan chosen value=" + String(chosen.value));
        }

        await sleep(800);

        // Handle Initial Reference Time datepicker (same as Add Cohort Subject)
        log("AddExistingSubject: Checking for Initial Reference Time datepicker");
        var initialSegmentRefDiv = modal.querySelector('div#initialSegmentReferenceDiv');
        var hasInitialRef = !!initialSegmentRefDiv;
        if (hasInitialRef) {
            var style = window.getComputedStyle(initialSegmentRefDiv);
            var isVisible = style.display !== "none" && style.visibility !== "hidden";
            if (!isVisible) {
                hasInitialRef = false;
                log("AddExistingSubject: initialSegmentReferenceDiv exists but is hidden; skipping datepicker");
            } else {
                log("AddExistingSubject: initialSegmentReferenceDiv present and visible");
            }
        } else {
            log("AddExistingSubject: initialSegmentReferenceDiv not present; skipping datepicker");
        }

        if (hasInitialRef) {
            var picker = modal.querySelector('span#initialSegmentReferencePicker');
            var addon = null;
            if (picker) {
                addon = picker.querySelector('span.input-group-addon');
            }
            var openedCal = false;
            if (addon) {
                addon.click();
                log("AddExistingSubject: calendar addon clicked");
                await sleep(400);
                openedCal = true;
            } else {
                log("AddExistingSubject: calendar addon not found");
            }

            // Auto-select today
            if (openedCal) {
                var waitedD = 0;
                var maxWaitD = 8000;
                var todayCell = null;
                while (waitedD < maxWaitD) {
                    if (isPaused()) {
                        log("AddExistingSubject: Paused; exiting during datepicker wait");
                        return;
                    }
                    var daysPanel = document.querySelector('div.datepicker-days');
                    if (daysPanel) {
                        todayCell = daysPanel.querySelector('td.day.active.today');
                        if (!todayCell) {
                            todayCell = daysPanel.querySelector('td.day.today.active');
                        }
                        if (!todayCell) {
                            todayCell = daysPanel.querySelector('td.day.today');
                        }
                        if (todayCell) {
                            break;
                        }
                    }
                    await sleep(300);
                    waitedD = waitedD + 300;
                }
                if (todayCell) {
                    todayCell.click();
                    log("AddExistingSubject: today clicked");
                    await sleep(300);
                } else {
                    log("AddExistingSubject: today cell not found in datepicker");
                }
            }

            var dateInputFinal = modal.querySelector('input#initialSegmentReference');
            var dateVal = dateInputFinal ? (dateInputFinal.value + "") : "";
            log("AddExistingSubject: date input value='" + String(dateVal) + "'");
        }

        var s2container = modal.querySelector('#s2id_volunteer');
        if (!s2container) {
            s2container = modal.querySelector('.select2-container.form-control.select2');
        }
        if (!s2container) {
            log("AddExistingSubject: Select2 container not found");
            return;
        }

        var s2choice = s2container.querySelector('a.select2-choice');
        if (s2choice) {
            s2choice.click();
            await sleep(200);
        }

        var s2drop = document.querySelector('#select2-drop.select2-drop-active');
        if (!s2drop) {
            s2drop = await waitForSelector('#select2-drop.select2-drop-active', 3000);
        }

        var s2input = null;
        if (s2drop) {
            s2input = s2drop.querySelector('input.select2-input');
        }
        if (!s2input) {
            s2input = await waitForSelector('#select2-drop.select2-drop-active input.select2-input', 3000);
        }
        if (!s2input) {
            log("AddExistingSubject: Select2 input not found");
            return;
        }

        log("AddExistingSubject: Typing subject number: " + selectedSubject.subjectNumber);
        s2input.value = selectedSubject.subjectNumber;
        var inpEvt = new Event("input", { bubbles: true });
        s2input.dispatchEvent(inpEvt);
        var keyEvt = new KeyboardEvent("keyup", { bubbles: true });
        s2input.dispatchEvent(keyEvt);

        await sleep(1500);

        var enterDown = new KeyboardEvent("keydown", { bubbles: true, key: "Enter", keyCode: 13 });
        s2input.dispatchEvent(enterDown);
        var enterUp = new KeyboardEvent("keyup", { bubbles: true, key: "Enter", keyCode: 13 });
        s2input.dispatchEvent(enterUp);

        await sleep(1000);

        var saveBtn = await waitForSelector('button#actionButton.btn.green[type="button"]', 5000);
        if (!saveBtn) {
            saveBtn = document.querySelector('button#actionButton');
        }
        if (!saveBtn) {
            log("AddExistingSubject: Save button not found");
            return;
        }

        saveBtn.click();
        log("AddExistingSubject: Save clicked");

        await sleep(1000);

        var errorAlert = document.querySelector('div.alert.alert-danger.alert-dismissable');
        if (errorAlert) {
            var errorText = (errorAlert.textContent + "").trim().toLowerCase();
            if (errorText.indexOf("please correct the validation errors") !== -1) {
                log("AddExistingSubject: Validation error detected");
                
                if (ADD_EXISTING_SUBJECT_POPUP_REF) {
                    ADD_EXISTING_SUBJECT_POPUP_REF.close();
                }
                
                clearAddExistingSubjectState();
                createPopup({
                    title: "Error",
                    content: "<div style='text-align:center;padding:20px;color:#f99;'>Error encountered. Stopping program.</div>",
                    width: "300px",
                    height: "auto"
                });
                return;
            }
        }

        var success = hasSuccessAlert();
        if (success) {
            log("AddExistingSubject: Save successful, setting add done flag and reloading");
            setAddExistingSubjectAddDone();
            await sleep(500);
            location.reload();
        }
    }

    async function performAddExistingSubjectActivation(selectedSubject) {
        log("AddExistingSubject: Starting activation for " + selectedSubject.subjectNumber);

        var listReady = await waitForListTable(15000);
        if (!listReady) {
            await sleep(800);
        }

        var targetRow = null;
        var tbody = document.querySelector('tbody#cohortAssignmentListBody');
        if (tbody) {
            var rows = tbody.querySelectorAll('tr');
            var rowIdx = 0;
            while (rowIdx < rows.length) {
                var row = rows[rowIdx];
                var subjectSpan = row.querySelector('span.tooltips[data-original-title="Subject Number"]');
                if (subjectSpan) {
                    var subjectNum = (subjectSpan.textContent + "").trim();
                    if (subjectNum === selectedSubject.subjectNumber) {
                        targetRow = row;
                        break;
                    }
                }
                rowIdx = rowIdx + 1;
            }
        }

        if (!targetRow) {
            log("AddExistingSubject: Target row not found for subject " + selectedSubject.subjectNumber);
            return;
        }

        log("AddExistingSubject: Found target row, activating plan");

        var actionBtn = getRowActionButton(targetRow);
        if (!actionBtn) {
            log("AddExistingSubject: Row action button not found");
            return;
        }

        actionBtn.click();
        await sleep(300);

        var planLink = getMenuLinkActivatePlan(targetRow);
        if (!planLink) {
            log("AddExistingSubject: Activate Plan link not found");
            return;
        }

        planLink.click();
        var ok1 = await clickBootboxOk(5000);
        if (!ok1) {
            await sleep(500);
        }

        await sleep(2000);

        actionBtn = getRowActionButton(targetRow);
        if (actionBtn) {
            actionBtn.click();
            await sleep(300);
        }

        var volunteerLink = getMenuLinkActivateVolunteer(targetRow);
        if (volunteerLink) {
            volunteerLink.click();
            var ok2 = await clickBootboxOk(5000);
            if (!ok2) {
                await sleep(500);
            }
        }

        log("AddExistingSubject: Activation complete");
        
        if (ADD_EXISTING_SUBJECT_POPUP_REF) {
            ADD_EXISTING_SUBJECT_POPUP_REF.close();
        }
        
        clearAddExistingSubjectState();

        createPopup({
            title: "Completed",
            content: "<div style='text-align:center;padding:20px;color:#9f9;'>Successfully added and activated existing subject: " + selectedSubject.subjectNumber + "</div>",
            width: "350px",
            height: "auto"
        });
    }

    function getMenuLinkActivateVolunteer(row) {
        var menu = row.querySelector("ul.dropdown-menu");
        if (!menu) {
            var parent = row.closest("tr");
            if (parent) {
                menu = parent.querySelector("ul.dropdown-menu");
            }
        }
        if (!menu) {
            menu = document.querySelector("ul.dropdown-menu");
        }
        if (!menu) {
            return null;
        }
        var links = menu.querySelectorAll("a");
        var i = 0;
        while (i < links.length) {
            var txt = (links[i].textContent + "").trim().toLowerCase();
            if (txt.indexOf("activate volunteer") !== -1 || txt.indexOf("activate subject") !== -1) {
                return links[i];
            }
            i = i + 1;
        }
        return null;
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
                log("Study edit flag set - will lock eligibility first");
            } catch (e) {}
            log("Routing to Study Metadata for eligibility lock");
            location.href = STUDY_METADATA_URL + "?autoeliglock=1";
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

    // Clear all Collect All related data
    function clearCollectAllData() {
        COLLECT_ALL_CANCELLED = false;
        COLLECT_ALL_POPUP_REF = null;
        log("CollectAll: data cleared");
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
            } else if (runMode === "consent" || runMode === "allconsent") {
                localStorage.removeItem(STORAGE_ICF_BARCODE_POPUP);
                if (ICF_BARCODE_POPUP_REF && ICF_BARCODE_POPUP_REF.close) {
                    try {
                        ICF_BARCODE_POPUP_REF.close();
                    } catch (e7) {}
                }
                ICF_BARCODE_POPUP_REF = null;
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
        var runPlansBtn = document.createElement("button");
        runPlansBtn.textContent = "Lock Activity Plans";
        runPlansBtn.style.background = "#4a90e2";
        runPlansBtn.style.color = "#fff";
        runPlansBtn.style.border = "none";
        runPlansBtn.style.borderRadius = "6px";
        runPlansBtn.style.padding = "8px";
        runPlansBtn.style.cursor = "pointer";
        runPlansBtn.style.fontWeight = "500";
        runPlansBtn.style.transition = "background 0.2s";
        runPlansBtn.onmouseenter = function() { this.style.background = "#357abd"; };
        runPlansBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var runStudyBtn = document.createElement("button");
        runStudyBtn.textContent = "Update Study Status";
        runStudyBtn.style.background = "#4a90e2";
        runStudyBtn.style.color = "#fff";
        runStudyBtn.style.border = "none";
        runStudyBtn.style.borderRadius = "6px";
        runStudyBtn.style.padding = "8px";
        runStudyBtn.style.cursor = "pointer";
        runStudyBtn.style.fontWeight = "500";
        runStudyBtn.style.transition = "background 0.2s";
        runStudyBtn.onmouseenter = function() { this.style.background = "#357abd"; };
        runStudyBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var runAddCohortBtn = document.createElement("button");
        runAddCohortBtn.textContent = "Add Cohort Subjects";
        runAddCohortBtn.style.background = "#4a90e2";
        runAddCohortBtn.style.color = "#fff";
        runAddCohortBtn.style.border = "none";
        runAddCohortBtn.style.borderRadius = "6px";
        runAddCohortBtn.style.padding = "8px";
        runAddCohortBtn.style.cursor = "pointer";
        runAddCohortBtn.style.fontWeight = "500";
        runAddCohortBtn.style.transition = "background 0.2s";
        runAddCohortBtn.onmouseenter = function() { this.style.background = "#357abd"; };
        runAddCohortBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var runAddExistingSubjectBtn = document.createElement("button");
        runAddExistingSubjectBtn.textContent = "Add Existing Subject";
        runAddExistingSubjectBtn.style.background = "#e67e22";
        runAddExistingSubjectBtn.style.color = "#fff";
        runAddExistingSubjectBtn.style.border = "none";
        runAddExistingSubjectBtn.style.borderRadius = "6px";
        runAddExistingSubjectBtn.style.padding = "8px";
        runAddExistingSubjectBtn.style.cursor = "pointer";
        runAddExistingSubjectBtn.style.fontWeight = "500";
        runAddExistingSubjectBtn.style.transition = "background 0.2s";
        runAddExistingSubjectBtn.onmouseenter = function() { this.style.background = "#d35400"; };
        runAddExistingSubjectBtn.onmouseleave = function() { this.style.background = "#e67e22"; };
        var runConsentBtn = document.createElement("button");
        runConsentBtn.textContent = "Run ICF Barcode";
        runConsentBtn.style.background = "#4a90e2";
        runConsentBtn.style.color = "#fff";
        runConsentBtn.style.border = "none";
        runConsentBtn.style.borderRadius = "6px";
        runConsentBtn.style.padding = "8px";
        runConsentBtn.style.cursor = "pointer";
        runConsentBtn.style.fontWeight = "500";
        runConsentBtn.style.transition = "background 0.2s";
        runConsentBtn.onmouseenter = function() { this.style.background = "#357abd"; };
        runConsentBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var runAllBtn = document.createElement("button");
        runAllBtn.textContent = "Run Button (1-5)";
        runAllBtn.style.background = "#5cb85c";
        runAllBtn.style.color = "#fff";
        runAllBtn.style.border = "none";
        runAllBtn.style.borderRadius = "6px";
        runAllBtn.style.padding = "8px";
        runAllBtn.style.cursor = "pointer";
        runAllBtn.style.fontWeight = "600";
        runAllBtn.style.transition = "background 0.2s";
        runAllBtn.onmouseenter = function() { this.style.background = "#449d44"; };
        runAllBtn.onmouseleave = function() { this.style.background = "#5cb85c"; };
        var runNonScrnBtn = document.createElement("button");
        runNonScrnBtn.textContent = "Import Cohort Subject";
        runNonScrnBtn.style.background = "#5b43c7";
        runNonScrnBtn.style.color = "#fff";
        runNonScrnBtn.style.border = "none";
        runNonScrnBtn.style.borderRadius = "6px";
        runNonScrnBtn.style.padding = "8px";
        runNonScrnBtn.style.cursor = "pointer";
        runNonScrnBtn.style.fontWeight = "500";
        runNonScrnBtn.style.transition = "background 0.2s";
        runNonScrnBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        runNonScrnBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
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
        var runFormOORBtn = document.createElement("button");
        runFormOORBtn.textContent = "Run Form (OOR) Below Range";
        runFormOORBtn.style.background = "#f0ad4e";
        runFormOORBtn.style.color = "#fff";
        runFormOORBtn.style.border = "none";
        runFormOORBtn.style.borderRadius = "6px";
        runFormOORBtn.style.padding = "8px";
        runFormOORBtn.style.cursor = "pointer";
        runFormOORBtn.style.fontWeight = "500";
        runFormOORBtn.style.transition = "background 0.2s";
        runFormOORBtn.onmouseenter = function() { this.style.background = "#ec971f"; };
        runFormOORBtn.onmouseleave = function() { this.style.background = "#f0ad4e"; };
        var runFormOORABtn = document.createElement("button");
        runFormOORABtn.textContent = "Run Form (OOR) Above Range";
        runFormOORABtn.style.background = "#f0ad4e";
        runFormOORABtn.style.color = "#fff";
        runFormOORABtn.style.border = "none";
        runFormOORABtn.style.borderRadius = "6px";
        runFormOORABtn.style.padding = "8px";
        runFormOORABtn.style.cursor = "pointer";
        runFormOORABtn.style.fontWeight = "500";
        runFormOORABtn.style.transition = "background 0.2s";
        runFormOORABtn.onmouseenter = function() { this.style.background = "#ec971f"; };
        runFormOORABtn.onmouseleave = function() { this.style.background = "#f0ad4e"; };
        var runFormIRBtn = document.createElement("button");
        runFormIRBtn.textContent = "Run Form (In Range)";
        runFormIRBtn.style.background = "#f0ad4e";
        runFormIRBtn.style.color = "#fff";
        runFormIRBtn.style.border = "none";
        runFormIRBtn.style.borderRadius = "6px";
        runFormIRBtn.style.padding = "8px";
        runFormIRBtn.style.cursor = "pointer";
        runFormIRBtn.style.fontWeight = "500";
        runFormIRBtn.style.transition = "background 0.2s";
        runFormIRBtn.onmouseenter = function() { this.style.background = "#ec971f"; };
        runFormIRBtn.onmouseleave = function() { this.style.background = "#f0ad4e"; };
        var parseMethodBtn = document.createElement("button");
        parseMethodBtn.textContent = "Parse Method";
        parseMethodBtn.style.background = "#4a90e2";
        parseMethodBtn.style.color = "#fff";
        parseMethodBtn.style.border = "none";
        parseMethodBtn.style.borderRadius = "6px";
        parseMethodBtn.style.padding = "8px";
        parseMethodBtn.style.cursor = "pointer";
        parseMethodBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        parseMethodBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
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

        var runLockSamplePathsBtn = document.createElement("button");
        runLockSamplePathsBtn.textContent = "Lock Sample Paths";
        runLockSamplePathsBtn.style.background = "#4a90e2";
        runLockSamplePathsBtn.style.color = "#fff";
        runLockSamplePathsBtn.style.border = "none";
        runLockSamplePathsBtn.style.borderRadius = "6px";
        runLockSamplePathsBtn.style.padding = "8px";
        runLockSamplePathsBtn.style.cursor = "pointer";
        runLockSamplePathsBtn.style.fontWeight = "500";
        runLockSamplePathsBtn.style.transition = "background 0.2s";
        runLockSamplePathsBtn.onmouseenter = function() { this.style.background = "#357abd"; };
        runLockSamplePathsBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };

        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import I/E";
        importEligBtn.style.background = "#38dae6";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = "6px";
        importEligBtn.style.padding = "8px";
        importEligBtn.style.cursor = "pointer";
        importEligBtn.style.fontWeight = "500";
        importEligBtn.style.transition = "background 0.2s";
        importEligBtn.onmouseenter = function() { this.style.background = "#2bb9c4"; };
        importEligBtn.onmouseleave = function() { this.style.background = "#38dae6"; };

        var clearMappingBtn = document.createElement("button");
        clearMappingBtn.textContent = "Clear Mapping";
        clearMappingBtn.style.background = "#38dae6";
        clearMappingBtn.style.color = "#fff";
        clearMappingBtn.style.border = "none";
        clearMappingBtn.style.borderRadius = "6px";
        clearMappingBtn.style.padding = "8px";
        clearMappingBtn.style.cursor = "pointer";
        clearMappingBtn.style.fontWeight = "500";
        clearMappingBtn.style.transition = "background 0.2s";
        clearMappingBtn.onmouseenter = function() { this.style.background = "#2bb9c4"; };
        clearMappingBtn.onmouseleave = function() { this.style.background = "#38dae6"; };

        var collectAllBtn = document.createElement("button");
        collectAllBtn.textContent = "Collect All";
        collectAllBtn.style.background = "#f0ad4e";
        collectAllBtn.style.color = "#fff";
        collectAllBtn.style.border = "none";
        collectAllBtn.style.borderRadius = "6px";
        collectAllBtn.style.padding = "8px";
        collectAllBtn.style.cursor = "pointer";
        collectAllBtn.style.fontWeight = "500";
        collectAllBtn.style.transition = "background 0.2s";
        collectAllBtn.onmouseenter = function() { this.style.background = "#ec971f"; };
        collectAllBtn.onmouseleave = function() { this.style.background = "#f0ad4e"; };

        btnRow.appendChild(runPlansBtn);
        btnRow.appendChild(runLockSamplePathsBtn);
        btnRow.appendChild(runStudyBtn);
        btnRow.appendChild(runAddCohortBtn);
        btnRow.appendChild(runAddExistingSubjectBtn);
        btnRow.appendChild(runConsentBtn);
        btnRow.appendChild(runAllBtn);
        btnRow.appendChild(runNonScrnBtn);
        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);

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
        runAddExistingSubjectBtn.addEventListener("click", function () {
            status.textContent = "Preparing Add Existing Subject...";
            log("Add Existing Subject clicked");
            try {
                localStorage.setItem(STORAGE_ADD_EXISTING_SUBJECT_MODE, "selectEpoch");
                localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_COLLECTED_DATA);
                localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_SELECTED_SUBJECT);
                localStorage.removeItem(STORAGE_ADD_EXISTING_SUBJECT_EDIT_DONE);
            } catch (e) {}
            location.href = STUDY_SHOW_URL + "?autoaddexistingsubject=1";
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
                log("Resumed");
            } else {
                setPaused(true);
                pauseBtn.textContent = "Resume";
                status.textContent = "Paused";
                log("Paused");
                clearAllRunState();
                COLLECT_ALL_CANCELLED = true;
                clearCollectAllData();
                clearEligibilityWorkingState();
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

        runBarcodeBtn.addEventListener("click", function () {
            BARCODE_START_TS = Date.now();
            clearBarcodeResult();
            var info = getSubjectFromBreadcrumbOrTooltip();
            var hasText = info.subjectText && info.subjectText.length > 0;
            var hasId = info.subjectId && info.subjectId.length > 0;
            if (!hasText && !hasId) {
                log("Run Barcode: subject breadcrumb or tooltip not found");
                return;
            }
            if (hasText) {
                setBarcodeSubjectText(info.subjectText);
            } else {
                clearBarcodeSubjectText();
            }
            if (hasId) {
                setBarcodeSubjectId(info.subjectId);
            } else {
                clearBarcodeSubjectId();
            }
            log("Opening Barcode Printing Subjects in background…");
            openInTab(location.origin + "/secure/barcodeprinting/subjects", false);

            // Create loading popup
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

            // Animate loading text
            var dots = 1;
            var loadingInterval = setInterval(function() {
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

            var waited = 0;
            var maxWait = 10000;
            var intervalMs = 500;
            var timer = setInterval(async function () {
                var r = getBarcodeResult();
                if (r && r.length > 0) {
                    clearInterval(loadingInterval);
                    if (popup && popup.close) {
                        popup.close();
                    }
                    log(r);
                    var inputBox = document.querySelector("input.bootbox-input.bootbox-input-text.form-control");
                    if (!inputBox) {
                        inputBox = await openBarcodeDataEntryModalIfNeeded(6000);
                    }
                    if (inputBox) {
                        inputBox.value = r;
                        var evt = new Event("input", { bubbles: true });
                        inputBox.dispatchEvent(evt);
                        log("Run Barcode: bootbox input autopopulated in parent page");
                        var okBtn = document.querySelector("button[data-bb-handler=\"confirm\"].btn.btn-primary");
                        if (!okBtn) {
                            okBtn = document.querySelector("button.btn.btn-primary[data-bb-handler=\"confirm\"]");
                        }
                        if (okBtn) {
                            okBtn.click();
                            log("Run Barcode: bootbox OK clicked in parent page");
                        } else {
                            log("Run Barcode: bootbox OK button not found in parent page");
                        }
                    } else {
                        log("Run Barcode: Unable to open barcode modal or locate input field");
                    }
                    var secs1 = (Date.now() - BARCODE_START_TS) / 1000;
                    var s1 = secs1.toFixed(2);
                    log("Run Barcode: elapsed " + String(s1) + " s");
                    BARCODE_START_TS = 0;
                    clearBarcodeResult();
                    clearInterval(timer);
                } else {
                    waited = waited + intervalMs;
                    if (waited >= maxWait) {
                        clearInterval(loadingInterval);
                        if (popup && popup.close) {
                            popup.close();
                        }
                        clearInterval(timer);
                        log("Run Barcode: timeout with no result");
                        var secs2 = (Date.now() - BARCODE_START_TS) / 1000;
                        var s2 = secs2.toFixed(2);
                        log("Run Barcode: elapsed " + String(s2) + " s");
                        BARCODE_START_TS = 0;
                    }
                }
            }, intervalMs);
        });
        clearMappingBtn.addEventListener("click", function () {
            log("ClearMapping: button clicked");

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
            processStudyShowPageForAddExistingSubject();
            return;
        }
        var onEditBasics = isStudyEditBasicsPage();
        if (onEditBasics) {
            processStudyEditBasicsPageIfFlag();
            return;
        }

        var onEpoch = isEpochShowPage();
        if (onEpoch) {
            // Check for Add Existing Subject scanning mode first
            var addExistingScanParam = getQueryParam("addexistingscan");
            var addExistingScanMode = null;
            try {
                addExistingScanMode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
            } catch (e) {}
            if (addExistingScanMode === "scanning" && addExistingScanParam === "1") {
                processEpochShowPageForAddExistingScan();
                return;
            }

            var imode = getRunMode();
            if (imode === "epochImport") {
                processEpochShowPageForImport();
                return;
            }
            if (imode === "epochAddCohort") {
                processEpochShowPageForAddCohort();
                return;
            }
            var addExistingMode = null;
            try {
                addExistingMode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
            } catch (e) {}
            if (addExistingMode === "addSubject") {
                processEpochShowPageForAddExistingSubject();
                return;
            }
            processEpochShowPage();
            return;
        }

        var onCohort = isCohortShowPage();
        if (onCohort) {
            // Check for Add Existing Subject scanning mode first
            var cohortScanParam = getQueryParam("addexistingscan");
            var cohortScanMode = null;
            try {
                cohortScanMode = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
            } catch (e) {}
            if (cohortScanMode === "scanning" && cohortScanParam === "1") {
                processCohortShowPageForAddExistingScan();
                return;
            }

            var amode = getRunMode();
            var autoImport = getQueryParam("autocohortimport");
            var autoAdd = getQueryParam("autocohortadd");
            var autoAddExisting = getQueryParam("autoaddexistingcohort");
            var addExistingMode2 = null;
            try {
                addExistingMode2 = localStorage.getItem(STORAGE_ADD_EXISTING_SUBJECT_MODE);
            } catch (e) {}
            if (amode === "epochImport" || autoImport === "1") {
                processCohortShowPageImportNonScrn();
            } else if (amode === "epochAddCohort" || autoAdd === "1") {
                processCohortShowPageAddCohort();
            } else if (addExistingMode2 === "addSubject" && autoAddExisting === "1") {
                processCohortShowPageForAddExistingSubject();
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