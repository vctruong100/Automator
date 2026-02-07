
// ==UserScript==
// @name ClinSpark SA Builder Automator
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
    var SA_BUILDER_POPUP_REF = null;
    var SA_BUILDER_PROGRESS_POPUP_REF = null;
    var SA_BUILDER_CANCELLED = false;
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
        } catch (e) {}
        SA_BUILDER_CANCELLED = false;
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

    // Select a value in a Select2 dropdown
    async function selectSelect2Value(selectId, value) {
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
        var segmentColumn = document.createElement("div");
        segmentColumn.style.cssText = "overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;";
        segmentColumn.id = "saBuilderSegments";

        // Study Events column (draggable items)
        var eventColumn = document.createElement("div");
        eventColumn.style.cssText = "overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;";
        eventColumn.id = "saBuilderEvents";

        // Forms column (checkboxes)
        var formColumn = document.createElement("div");
        formColumn.style.cssText = "overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;";
        formColumn.id = "saBuilderForms";

        // Time inputs column
        var timeColumn = document.createElement("div");
        timeColumn.style.cssText = "display:flex;flex-direction:column;gap:12px;padding:8px;border:1px solid #333;border-radius:4px;background:#1a1a1a;";
        timeColumn.id = "saBuilderTime";

        // Track segment-event assignments
        var segmentEventMap = {};

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

                            var evLabel = document.createElement("span");
                            evLabel.textContent = ev.text;
                            evLabel.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";

                            var removeBtn = document.createElement("button");
                            removeBtn.textContent = "×";
                            removeBtn.style.cssText = "background:transparent;border:none;color:#f66;cursor:pointer;font-size:16px;padding:0 4px;";
                            removeBtn.addEventListener("click", (function(segV, evV) {
                                return function(e) {
                                    e.stopPropagation();
                                    var arr = segmentEventMap[segV] || [];
                                    segmentEventMap[segV] = arr.filter(function(x) { return x.value !== evV; });
                                    renderSegments(segmentSearch.value);
                                };
                            })(segVal, ev.value));

                            evItem.appendChild(evLabel);
                            evItem.appendChild(removeBtn);

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
    } catch (e) {}
    
    await sleep(300);
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

                eventColumn.appendChild(evItem);
            }
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
                checkbox.style.cssText = "width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-top:2px;";

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

        timeColumn.appendChild(createTimeInput("Days", "saBuilderDays", 200));
        timeColumn.appendChild(createTimeInput("Hours", "saBuilderHours", 23));
        timeColumn.appendChild(createTimeInput("Minutes", "saBuilderMinutes", 59));
        timeColumn.appendChild(createTimeInput("Seconds", "saBuilderSeconds", 59));

        // Initial render
        renderSegments("");
        renderStudyEvents("");
        renderForms("");

        // Search handlers
        segmentSearch.addEventListener("input", function() { renderSegments(this.value); });
        eventSearch.addEventListener("input", function() { renderStudyEvents(this.value); });
        formSearch.addEventListener("input", function() { renderForms(this.value); });

        contentRow.appendChild(segmentColumn);
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

            // Store selection
            var userSelection = {
                segments: selectedSegments,
                forms: selectedForms,
                timeOffset: timeOffset
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
        stopBtn.textContent = "Stop and Close";
        stopBtn.style.cssText = "padding:10px 24px;border-radius:6px;border:none;background:#dc3545;color:#fff;font-weight:600;cursor:pointer;";
        stopBtn.addEventListener("click", function() {
            SA_BUILDER_CANCELLED = true;
            clearSABuilderStorage();
            if (SA_BUILDER_PROGRESS_POPUP_REF) {
                SA_BUILDER_PROGRESS_POPUP_REF.close();
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
            log("SA Builder: stopped by user");
        });

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
                clearSABuilderStorage();
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
        });

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

                await sleep(1000);

                // Select segment
                log("SA Builder: selecting segment " + item.segmentText + " (value: " + item.segmentValue + ")");
                await selectSelect2Value("segment", item.segmentValue);
                await sleep(500);

                // Select study event
                log("SA Builder: selecting study event " + item.eventText + " (value: " + item.eventValue + ")");
                await selectSelect2Value("studyEvent", item.eventValue);
                await sleep(500);

                // Select form
                log("SA Builder: selecting form " + item.formText + " (value: " + item.formValue + ")");
                await selectSelect2Value("form", item.formValue);
                await sleep(500);

                // Clear pre/post window fields
                var preWindow = document.getElementById("preWindow");
                var postWindow = document.getElementById("postWindow");
                if (preWindow) {
                    preWindow.value = "";
                    preWindow.dispatchEvent(new Event("input", { bubbles: true }));
                }
                if (postWindow) {
                    postWindow.value = "";
                    postWindow.dispatchEvent(new Event("input", { bubbles: true }));
                }

                // Set time offset values
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

                await sleep(300);

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

                // Update status
                mappedItems[idx].status = "Complete";
                progressContent.updateItem(idx, "Complete");
                log("SA Builder: item completed - " + item.key);

                // Wait for table refresh
                await sleep(2000);

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

        // Clear running state
        try {
            localStorage.removeItem(STORAGE_SA_BUILDER_RUNNING);
            localStorage.removeItem(STORAGE_SA_BUILDER_CURRENT_INDEX);
        } catch (e) {}
    }

    // Main entry point for SA Builder
    async function runSABuilder() {
        log("SA Builder: starting");
        SA_BUILDER_CANCELLED = false;

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
            onClose: function() {
                SA_BUILDER_CANCELLED = true;
                clearSABuilderStorage();
            }
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
                clearSABuilderStorage();
                SA_BUILDER_POPUP_REF = null;
            }
        });

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

        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);
        btnRow.appendChild(saBuilderBtn);

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
        saBuilderBtn.addEventListener("click", async function () {
            log("SA Builder: button clicked");
            await runSABuilder();
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