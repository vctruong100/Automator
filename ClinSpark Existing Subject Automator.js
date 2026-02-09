
// ==UserScript==
// @name ClinSpark EA Automator
// @namespace vinh.activity.plan.state
// @version 1.1.0
// @description
// @match https://cenexeltest.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Basic%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Basic%20Automator.js
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

    // Run Find Study Event

    var STORAGE_FIND_FORM_PENDING = "activityPlanState.findForm.pending";
    var STORAGE_FIND_FORM_KEYWORD = "activityPlanState.findForm.keyword";
    var STORAGE_FIND_FORM_SUBJECT = "activityPlanState.findForm.subject";
    var STORAGE_FIND_FORM_STATUS_VALUES = "activityPlanState.findForm.statusValues";

    var STORAGE_FIND_STUDY_EVENT_PENDING = "activityPlanState.findStudyEvent.pending";
    var STORAGE_FIND_STUDY_EVENT_KEYWORD = "activityPlanState.findStudyEvent.keyword";
    var STORAGE_FIND_STUDY_EVENT_SUBJECT = "activityPlanState.findStudyEvent.subject";
    var STORAGE_FIND_STUDY_EVENT_STATUS_VALUES = "activityPlanState.findStudyEvent.statusValues";

    // Add Existing Subject Feature
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

        var addExistingSubjectBtn = document.createElement("button");
        addExistingSubjectBtn.textContent = "Add Existing Subject";
        addExistingSubjectBtn.style.background = "#2e7d32";
        addExistingSubjectBtn.style.color = "#fff";
        addExistingSubjectBtn.style.border = "none";
        addExistingSubjectBtn.style.borderRadius = "6px";
        addExistingSubjectBtn.style.padding = "8px";
        addExistingSubjectBtn.style.cursor = "pointer";
        addExistingSubjectBtn.style.fontWeight = "500";
        addExistingSubjectBtn.style.transition = "background 0.2s";
        addExistingSubjectBtn.onmouseenter = function() { this.style.background = "#1b5e20"; };
        addExistingSubjectBtn.onmouseleave = function() { this.style.background = "#2e7d32"; };

        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(addExistingSubjectBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);

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
        runBarcodeBtn.addEventListener("click", async function () {
            log("Run Barcode: button clicked");
            await APS_RunBarcode();
        });
        addExistingSubjectBtn.addEventListener("click", async function () {
            log("Add Existing Subject: button clicked");
            await startAddExistingSubject();
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


    //==========================
    // ADD EXISTING SUBJECT FEATURE
    //==========================
    // This section contains all functions for adding an existing subject from other epochs.
    //==========================

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

        // Click the select2 choice link to open the dropdown
        var select2Container = document.getElementById("s2id_volunteer");
        if (!select2Container) {
            showAESError("Could not find volunteer search container.");
            return;
        }

        var select2Choice = select2Container.querySelector("a.select2-choice");
        if (select2Choice) {
            select2Choice.click();
            log("AES: clicked select2 choice to open dropdown");
        } else {
            select2Container.click();
            log("AES: clicked select2 container");
        }

        await sleep(800);

        // Find the search input in the active dropdown
        var searchInput = document.querySelector(".select2-drop-active input.select2-input");
        if (!searchInput) {
            // Try alternative selector for the search input
            searchInput = document.querySelector("input.select2-input.select2-focused");
        }
        if (!searchInput) {
            // Try by partial ID match
            var allInputs = document.querySelectorAll("input.select2-input");
            for (var inp = 0; inp < allInputs.length; inp++) {
                var inputId = allInputs[inp].id || "";
                if (inputId.indexOf("s2id_autogen") !== -1 && inputId.indexOf("_search") !== -1) {
                    searchInput = allInputs[inp];
                    break;
                }
            }
        }

        if (!searchInput) {
            log("AES: could not find select2 search input, trying direct input");
            showAESError("Could not find volunteer search input.");
            return;
        }

        log("AES: found search input: " + (searchInput.id || "no-id"));

        // Focus the input
        searchInput.focus();
        await sleep(200);

        // Set the value directly
        searchInput.value = subjectNumber;
        log("AES: set search input value to: " + subjectNumber);

        // Try jQuery trigger first (Select2 is a jQuery plugin)
        if (typeof jQuery !== "undefined" && jQuery(searchInput).length) {
            log("AES: using jQuery to trigger events");
            jQuery(searchInput).trigger("input");
            jQuery(searchInput).trigger("keyup");
            // Also try triggering change
            jQuery(searchInput).trigger("change");
        } else if (typeof $ !== "undefined" && $(searchInput).length) {
            log("AES: using $ to trigger events");
            $(searchInput).trigger("input");
            $(searchInput).trigger("keyup");
            $(searchInput).trigger("change");
        } else {
            // Fallback to native events
            log("AES: jQuery not found, using native events");
            var inputEvt = new Event("input", { bubbles: true });
            searchInput.dispatchEvent(inputEvt);
            var keyupEvt = new KeyboardEvent("keyup", { bubbles: true, keyCode: 13 });
            searchInput.dispatchEvent(keyupEvt);
        }

        log("AES: triggered search events for: " + subjectNumber);

        // Wait for AJAX search results (increased wait time)
        await sleep(4000);

        // Look for results in the active dropdown
        var resultsContainer = document.querySelector(".select2-drop-active .select2-results");
        if (!resultsContainer) {
            resultsContainer = document.querySelector(".select2-results");
        }

        var resultsList = resultsContainer ? resultsContainer.querySelectorAll("li.select2-result") : [];
        log("AES: found " + resultsList.length + " results");

        var found = false;
        for (var r = 0; r < resultsList.length; r++) {
            var resultEl = resultsList[r];
            var resultText = (resultEl.textContent || "").trim();
            log("AES: result " + r + ": " + resultText.substring(0, 50));
            
            // Check if result contains subject number and is selectable
            if (resultText.indexOf(subjectNumber) !== -1 && !resultEl.classList.contains("select2-disabled")) {
                resultEl.click();
                found = true;
                log("AES: selected subject from dropdown");
                break;
            }
        }

        if (!found) {
            // Try clicking using mousedown/mouseup events (some Select2 versions need this)
            for (var r2 = 0; r2 < resultsList.length; r2++) {
                var resultEl2 = resultsList[r2];
                var resultText2 = (resultEl2.textContent || "").trim();
                if (resultText2.indexOf(subjectNumber) !== -1 && !resultEl2.classList.contains("select2-disabled")) {
                    var mousedownEvt = new MouseEvent("mousedown", { bubbles: true });
                    resultEl2.dispatchEvent(mousedownEvt);
                    var mouseupEvt = new MouseEvent("mouseup", { bubbles: true });
                    resultEl2.dispatchEvent(mouseupEvt);
                    found = true;
                    log("AES: selected subject via mousedown/mouseup");
                    break;
                }
            }
        }

        if (!found) {
            showAESError("Could not find subject " + subjectNumber + " in volunteer search. Found " + resultsList.length + " results.");
            return;
        }

        await sleep(500);
        updateAESProgressStatus("Saving assignment...");
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

        showAESProgressPopup("Locating new assignment...");

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
                targetRow = row;
                var rowId = row.id || "";
                var match = rowId.match(/ca_(\d+)/);
                if (match) assignmentId = match[1];
                break;
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
        
        showAESProgressPopup("Activating plan...");

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
            
            await sleep(3000);
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
            
            await sleep(2000);
            showAESComplete();
        } else {
            log("AES: Activate Volunteer link not found after waiting");
            showAESComplete();
        }
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

        // Process Add Existing Subject on page load
        if (runModeRaw === RUNMODE_ADD_EXISTING_SUBJECT) {
            processAESOnPageLoad();
            return;
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();