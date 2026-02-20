
// ==UserScript==
// @name ClinSpark Admin Automator
// @namespace vinh.activity.plan.state
// @version 1.1.3
// @description
// @match https://cenexel.clinspark.com/*
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

    // Parse Deviation Feature
    var STORAGE_PARSE_DEVIATION_RUNNING = "activityPlanState.parseDeviation.running";
    var STORAGE_PARSE_DEVIATION_KEYWORDS = "activityPlanState.parseDeviation.keywords";
    var STORAGE_PARSE_DEVIATION_TIMESTAMP = "activityPlanState.parseDeviation.timestamp";
    var STORAGE_PARSE_DEVIATION_RESULTS = "activityPlanState.parseDeviation.results";
    var PARSE_DEVIATION_DATA_LIST_URL = "https://cenexel.clinspark.com/secure/study/data/list";
    var PARSE_DEVIATION_POPUP_REF = null;
    var PARSE_DEVIATION_CANCELED = false;

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
            if (e.key === "`" || e.key === "~") {
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
            var warningContent = document.createElement("div");
            warningContent.style.textAlign = "center";
            warningContent.style.padding = "20px";
            warningContent.innerHTML = "<div style='font-size:48px;margin-bottom:16px;'>⚠️</div>" +
                "<div style='font-size:16px;font-weight:600;margin-bottom:12px;'>Invalid Page</div>" +
                "<div style='font-size:14px;color:#aaa;line-height:1.6;'>You must be on one of these pages:<br><br>" +
                "<code style='background:#2a2a2a;padding:4px 8px;border-radius:4px;font-size:12px;'>cenexel.clinspark.com/.../studyevent</code><br><br>" +
                "<code style='background:#2a2a2a;padding:4px 8px;border-radius:4px;font-size:12px;'>cenexeltest.clinspark.com/.../studyevent</code></div>";

            createPopup({
                title: "Parse Study Event",
                content: warningContent,
                width: "450px",
                height: "auto"
            });
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

    //==========================
    // PARSE DEVIATION FEATURE
    //==========================
    // This section contains all functions related to the Parse Deviation feature.
    // This feature automates extracting deviation form data from the study data list
    // and formatting it for Excel output.
    //==========================

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
            localStorage.removeItem(STORAGE_PARSE_DEVIATION_TIMESTAMP);
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

    //==========================
    // MAKE PANEL FUNCTIONS
    //==========================
    // This section contains functions used to create and manage the panel UI.
    // These functions are used to create the panel UI and manage its state.
    //==========================
    function makePanel() {
        // ... (rest of the code remains the same)
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
        title.textContent = "Admin Automator";
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

        var parseStudyEventBtn = document.createElement("button");
        parseStudyEventBtn.textContent = "Parse Study Event";
        parseStudyEventBtn.style.background = "#17a2b8";
        parseStudyEventBtn.style.color = "#fff";
        parseStudyEventBtn.style.border = "none";
        parseStudyEventBtn.style.borderRadius = "6px";
        parseStudyEventBtn.style.padding = "8px";
        parseStudyEventBtn.style.cursor = "pointer";
        parseStudyEventBtn.style.fontWeight = "500";
        parseStudyEventBtn.style.transition = "background 0.2s";
        parseStudyEventBtn.onmouseenter = function() { this.style.background = "#138496"; };
        parseStudyEventBtn.onmouseleave = function() { this.style.background = "#17a2b8"; };
        parseStudyEventBtn.addEventListener("click", function() {
            log("[ParseStudyEvent] Button clicked");
            APS_ParseStudyEvent();
        });

        var searchMethodsBtn = document.createElement("button");
        searchMethodsBtn.textContent = "Search Methods";
        searchMethodsBtn.style.background = "#28a745";
        searchMethodsBtn.style.color = "#fff";
        searchMethodsBtn.style.border = "none";
        searchMethodsBtn.style.borderRadius = "6px";
        searchMethodsBtn.style.padding = "8px";
        searchMethodsBtn.style.cursor = "pointer";
        searchMethodsBtn.style.fontWeight = "500";
        searchMethodsBtn.style.transition = "background 0.2s";
        searchMethodsBtn.onmouseenter = function() { this.style.background = "#218838"; };
        searchMethodsBtn.onmouseleave = function() { this.style.background = "#28a745"; };
        searchMethodsBtn.addEventListener("click", function() {
            log("[SearchMethods] Button clicked");
            openMethodsLibraryModal();
        });

        var parseDeviationBtn = document.createElement("button");
        parseDeviationBtn.textContent = "Parse Deviation";
        parseDeviationBtn.style.background = "#4a90e2";
        parseDeviationBtn.style.color = "#fff";
        parseDeviationBtn.style.border = "none";
        parseDeviationBtn.style.borderRadius = "6px";
        parseDeviationBtn.style.padding = "8px";
        parseDeviationBtn.style.cursor = "pointer";
        parseDeviationBtn.style.fontWeight = "500";
        parseDeviationBtn.style.transition = "background 0.2s";
        parseDeviationBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        parseDeviationBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        parseDeviationBtn.addEventListener("click", function() {
            log("[ParseDeviation] Button clicked");
            APS_ParseDeviation();
        });

        // btnRow.appendChild(runBarcodeBtn);
        // btnRow.appendChild(pauseBtn);
        // btnRow.appendChild(clearLogsBtn);
        // btnRow.appendChild(toggleLogsBtn);
        // btnRow.appendChild(parseStudyEventBtn);
        // btnRow.appendChild(searchMethodsBtn);
        btnRow.appendChild(parseDeviationBtn);

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
        //  bodyContainer.appendChild(status);
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
        // bodyContainer.appendChild(logBox);
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

        // Check for Parse Deviation continuation after page navigation
        parseDeviationCheckOnPageLoad();

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