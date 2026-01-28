
// ==UserScript==
// @name ClinSpark Delete Automator
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

    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const PANEL_TOGGLE_KEY = "F2";
    const ELIGIBILITY_LIST_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";

    // Delete Templates Feature
    var DELETE_TEMPLATES_CANCELLED = false;
    var DELETE_TEMPLATES_POPUP_REF = null;
    var DELETE_TEMPLATES_RUNNING = false;
    var DELETE_TEMPLATES_CURRENT_ITEM = "";
    var DELETE_TEMPLATES_TYPE = "";
    const TEMPLATE_URLS = {
        form: "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/form",
        itemgroup: "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/itemgroup",
        item: "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/item",
        codelist: "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/codelist",
        method: "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/list/method"
    };

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

    var FORM_LIST_URL = "https://cenexeltest.clinspark.com/secure/study/data/list";
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
    const STORAGE_RUN_ALL_POPUP = "activityPlanState.runAllPopup";
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

    function clearCollectAllData() {
        COLLECT_ALL_CANCELLED = false;
        COLLECT_ALL_POPUP_REF = null;
        log("CollectAll: data cleared");
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
    function clearBarcodeResult() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_RESULT);
        } catch (e) {}
    }
    // Clear all run state from storage.
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
            popup.remove();
        });

        headerBar.appendChild(closeBtn);
        popup.appendChild(headerBar);

        document.addEventListener("keydown", function(e) {
            if (e.key === "Escape") {
                try { clearAllRunState(); } catch(e2){}
                if (popup && popup.remove) popup.remove();
            }
        });

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
    // DELETE TEMPLATES FEATURE
    //==========================
    // This section contains functions for deleting template items (Forms, Item Groups, Items, Codelists, Methods)
    // from the ClinSpark study library in the background.
    //==========================

    function resetDeleteTemplatesState() {
        DELETE_TEMPLATES_CANCELLED = false;
        DELETE_TEMPLATES_RUNNING = false;
        DELETE_TEMPLATES_CURRENT_ITEM = "";
        DELETE_TEMPLATES_TYPE = "";
    }

    function cancelDeleteTemplates() {
        DELETE_TEMPLATES_CANCELLED = true;
        DELETE_TEMPLATES_RUNNING = false;
        log("DeleteTemplates: cancelled by user");
    }

    function updateDeleteTemplatesStatus(statusEl, text, isRunning) {
        if (!statusEl) return;
        var animationHtml = isRunning ? '<span class="delete-templates-spinner"></span> ' : '';
        statusEl.innerHTML = animationHtml + text;
    }

    function addDeletedItemToStatus(statusEl, itemName) {
        if (!statusEl) return;
        var itemDiv = document.createElement("div");
        itemDiv.style.padding = "4px 0";
        itemDiv.style.borderBottom = "1px solid #2a2a2a";
        itemDiv.textContent = "✓ " + itemName;
        statusEl.appendChild(itemDiv);
        statusEl.scrollTop = statusEl.scrollHeight;
    }

    async function deleteTemplatesForType(templateType, statusEl, countEl) {
        if (DELETE_TEMPLATES_CANCELLED) {
            log("DeleteTemplates: already cancelled before start");
            return;
        }

        DELETE_TEMPLATES_RUNNING = true;
        DELETE_TEMPLATES_TYPE = templateType;
        var listUrl = TEMPLATE_URLS[templateType];
        var deletedCount = 0;
        var requiresUnlock = (templateType === "form");
        var skippedItemNames = [];

        log("DeleteTemplates: starting for type=" + templateType + " url=" + listUrl);
        updateDeleteTemplatesStatus(statusEl, "Navigating to " + templateType + " list...", true);

        try {
            while (!DELETE_TEMPLATES_CANCELLED) {
                var listHtml = await fetchPage(listUrl);
                var listDoc = parseHtml(listHtml);
                var tbody = listDoc.querySelector("#listTable tbody");

                if (!tbody) {
                    log("DeleteTemplates: listTable tbody not found");
                    updateDeleteTemplatesStatus(statusEl, "Error: Could not find table", false);
                    break;
                }

                var rows = tbody.querySelectorAll("tr");
                log("DeleteTemplates: found " + rows.length + " rows");

                if (rows.length < 1) {
                    log("DeleteTemplates: no more items to delete (table empty)");
                    updateDeleteTemplatesStatus(statusEl, "Completed! Deleted " + deletedCount + " items.", false);
                    break;
                }

                // Check for pause frequently
                if (DELETE_TEMPLATES_CANCELLED) break;

                var targetRow = null;
                var itemLink = null;
                var itemName = null;
                var itemHref = null;
                
                // Find first non-skipped item
                for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
                    if (DELETE_TEMPLATES_CANCELLED) break;
                    
                    var currentRow = rows[rowIndex];
                    var firstTd = currentRow.querySelector("td");
                    if (!firstTd) continue;
                    
                    var currentLink = firstTd.querySelector("a");
                    if (!currentLink) continue;
                    
                    var currentName = (currentLink.textContent || "").trim();
                    if (skippedItemNames.indexOf(currentName) === -1) {
                        targetRow = currentRow;
                        itemLink = currentLink;
                        itemName = currentName;
                        itemHref = itemLink.getAttribute("href");
                        break;
                    }
                }
                
                if (!itemLink || !itemName) {
                    log("DeleteTemplates: no non-skipped items found in table");
                    updateDeleteTemplatesStatus(statusEl, "Completed! Deleted " + deletedCount + " items.", false);
                    break;
                }
                
                DELETE_TEMPLATES_CURRENT_ITEM = itemName;

                log("DeleteTemplates: processing item '" + itemName + "' href=" + itemHref);
                updateDeleteTemplatesStatus(statusEl, "Deleting: " + itemName, true);

                // Check for pause before making requests
                if (DELETE_TEMPLATES_CANCELLED) break;
                
                var itemUrl = location.origin + itemHref;
                var itemHtml = await fetchPage(itemUrl);
                var itemDoc = parseHtml(itemHtml);

                if (DELETE_TEMPLATES_CANCELLED) break;

                if (requiresUnlock) {
                    var unlocked = await handleFormUnlockIfNeeded(itemDoc, itemUrl, statusEl);
                    if (DELETE_TEMPLATES_CANCELLED) break;
                    if (!unlocked) {
                        log("DeleteTemplates: failed to unlock form, skipping");
                        updateDeleteTemplatesStatus(statusEl, "Error unlocking: " + itemName, false);
                        await sleep(1000);
                        continue;
                    }
                    itemHtml = await fetchPage(itemUrl);
                    itemDoc = parseHtml(itemHtml);
                }

                if (DELETE_TEMPLATES_CANCELLED) break;

                var deleteResult = await performDelete(itemDoc, templateType, itemUrl);

                if (deleteResult.success) {
                    deletedCount++;
                    if (countEl) {
                        countEl.textContent = "Deleted: " + deletedCount;
                    }
                    addDeletedItemToStatus(statusEl, itemName);
                    log("DeleteTemplates: successfully deleted '" + itemName + "' (total: " + deletedCount + ")");
                } else if (deleteResult.hasAlert) {
                    log("DeleteTemplates: alert detected for '" + itemName + "', adding to skip list");
                    log("DeleteTemplates: alert message: " + deleteResult.alertMessage);
                    skippedItemNames.push(itemName);
                    await sleep(500);
                } else {
                    log("DeleteTemplates: failed to delete '" + itemName + "', adding to skip list");
                    skippedItemNames.push(itemName);
                    await sleep(500);
                }

                await sleep(300);
            }

        } catch (err) {
            log("DeleteTemplates: error - " + String(err));
            updateDeleteTemplatesStatus(statusEl, "Error: " + String(err), false);
        }

        DELETE_TEMPLATES_RUNNING = false;
        DELETE_TEMPLATES_CURRENT_ITEM = "";
        log("DeleteTemplates: finished. Total deleted: " + deletedCount);
    }

    async function handleFormUnlockIfNeeded(itemDoc, itemUrl, statusEl) {
        var allLinks = itemDoc.querySelectorAll("a");
        var unlockUrl = null;

        for (var i = 0; i < allLinks.length; i++) {
            var link = allLinks[i];
            var href = link.getAttribute("href") || "";
            var onclick = link.getAttribute("onclick") || "";
            
            if (href.indexOf("/locking/") !== -1 || onclick.indexOf("/locking/") !== -1) {
                var icon = link.querySelector("i.fa-unlock");
                if (icon) {
                    if (href && href !== "#" && href.indexOf("/locking/") !== -1) {
                        unlockUrl = href;
                    } else {
                        var match = onclick.match(/['"]([^'"]*\/locking\/[^'"]*)['"]/);
                        if (match) {
                            unlockUrl = match[1];
                        }
                    }
                    if (unlockUrl) {
                        log("DeleteTemplates: found unlock link: " + unlockUrl);
                        break;
                    }
                }
            }
        }

        if (!unlockUrl) {
            log("DeleteTemplates: no unlock needed (form is not locked)");
            return true;
        }

        log("DeleteTemplates: form is locked, need to unlock");
        updateDeleteTemplatesStatus(statusEl, "Unlocking form...", true);

        var fullUnlockUrl = location.origin + unlockUrl;
        log("DeleteTemplates: fetching unlock modal from " + fullUnlockUrl);

        var unlockHtml = await fetchPage(fullUnlockUrl);
        var unlockDoc = parseHtml(unlockHtml);

        var textarea = unlockDoc.querySelector("textarea#reasonForChange");
        if (!textarea) {
            log("DeleteTemplates: reasonForChange textarea not found in unlock modal");
            return false;
        }

        var form = unlockDoc.querySelector("form");
        if (!form) {
            log("DeleteTemplates: unlock form not found");
            return false;
        }

        var formAction = form.getAttribute("action");
        if (!formAction) {
            log("DeleteTemplates: form action not found");
            return false;
        }

        var formData = new URLSearchParams();
        var inputs = form.querySelectorAll("input, textarea, select");
        for (var j = 0; j < inputs.length; j++) {
            var input = inputs[j];
            var name = input.getAttribute("name");
            if (!name) continue;

            var value = "";
            var type = (input.getAttribute("type") || "").toLowerCase();

            if (name === "reasonForChange") {
                value = "Test";
            } else if (type === "checkbox" || type === "radio") {
                if (input.checked) {
                    value = input.value || "on";
                } else {
                    continue;
                }
            } else {
                value = input.value || "";
            }

            formData.append(name, value);
        }

        var submitUrl = location.origin + formAction;
        log("DeleteTemplates: submitting unlock form to " + submitUrl);

        await submitForm(submitUrl, formData.toString());
        log("DeleteTemplates: unlock form submitted, waiting for page refresh");
        await sleep(1500);

        return true;
    }

    async function performDelete(itemDoc, templateType, itemUrl) {
        var allLinks = itemDoc.querySelectorAll("a");
        var deleteUrl = null;

        for (var i = 0; i < allLinks.length; i++) {
            var link = allLinks[i];
            var onclick = link.getAttribute("onclick") || "";
            var match = onclick.match(/maybeRedirect\(['"]([^'"]*delete[^'"]*)['"]\)/);
            if (match) {
                deleteUrl = match[1];
                log("DeleteTemplates: found delete URL: " + deleteUrl);
                break;
            }
        }

        if (!deleteUrl) {
            log("DeleteTemplates: no delete link found on page");
            return { success: false, hasAlert: false };
        }

        var fullDeleteUrl = location.origin + deleteUrl;
        log("DeleteTemplates: performing delete at " + fullDeleteUrl);

        try {
            var responseHtml = await fetchPage(fullDeleteUrl);
            var responseDoc = parseHtml(responseHtml);
            
            var alertDanger = responseDoc.querySelector(".alert-danger");
            if (alertDanger) {
                var alertMessage = (alertDanger.textContent || "").trim();
                log("DeleteTemplates: alert detected after delete attempt");
                return { success: false, hasAlert: true, alertMessage: alertMessage };
            }
            
            log("DeleteTemplates: delete request completed successfully");
            return { success: true, hasAlert: false };
        } catch (e) {
            log("DeleteTemplates: delete request failed - " + String(e));
            return { success: false, hasAlert: false };
        }
    }

    function showDeleteTemplatesPopup() {
        if (DELETE_TEMPLATES_POPUP_REF) {
            try {
                DELETE_TEMPLATES_POPUP_REF.close();
            } catch (e) {}
        }

        resetDeleteTemplatesState();

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";

        var spinnerStyle = document.createElement("style");
        spinnerStyle.textContent = `
            @keyframes deleteTemplatesSpinner {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            .delete-templates-spinner {
                display: inline-block;
                width: 14px;
                height: 14px;
                border: 2px solid #666;
                border-top-color: #fff;
                border-radius: 50%;
                animation: deleteTemplatesSpinner 0.8s linear infinite;
                vertical-align: middle;
                margin-right: 6px;
            }
        `;
        container.appendChild(spinnerStyle);

        var instruction = document.createElement("div");
        instruction.style.marginBottom = "8px";
        instruction.style.fontSize = "14px";
        instruction.textContent = "Select a template type to delete all items:";
        container.appendChild(instruction);

        var btnContainer = document.createElement("div");
        btnContainer.style.display = "grid";
        btnContainer.style.gridTemplateColumns = "1fr 1fr";
        btnContainer.style.gap = "8px";

        var templateTypes = [
            { key: "form", label: "Forms" },
            { key: "itemgroup", label: "Item Groups" },
            { key: "item", label: "Items" },
            { key: "codelist", label: "Codelists" },
            { key: "method", label: "Methods" }
        ];

        var statusEl = document.createElement("div");
        statusEl.style.marginTop = "12px";
        statusEl.style.padding = "10px";
        statusEl.style.background = "#1a1a1a";
        statusEl.style.border = "1px solid #333";
        statusEl.style.borderRadius = "6px";
        statusEl.style.fontSize = "13px";
        statusEl.style.minHeight = "100px";
        statusEl.style.maxHeight = "300px";
        statusEl.style.overflowY = "auto";
        statusEl.textContent = "Ready - Select a template type to begin";

        var countEl = document.createElement("div");
        countEl.style.marginTop = "8px";
        countEl.style.fontSize = "14px";
        countEl.style.fontWeight = "600";
        countEl.textContent = "Deleted: 0";

        templateTypes.forEach(function(type) {
            var btn = document.createElement("button");
            btn.textContent = type.label;
            btn.style.background = "#5b43c7";
            btn.style.color = "#fff";
            btn.style.border = "none";
            btn.style.borderRadius = "6px";
            btn.style.padding = "12px 16px";
            btn.style.cursor = "pointer";
            btn.style.fontWeight = "500";
            btn.style.fontSize = "14px";
            btn.style.transition = "background 0.2s";

            btn.onmouseenter = function() { 
                if (!DELETE_TEMPLATES_RUNNING) {
                    btn.style.background = "#4a37a0"; 
                }
            };
            btn.onmouseleave = function() { 
                if (!DELETE_TEMPLATES_RUNNING) {
                    btn.style.background = "#5b43c7"; 
                }
            };

            btn.addEventListener("click", async function() {
                if (DELETE_TEMPLATES_RUNNING) {
                    log("DeleteTemplates: already running, ignoring click");
                    return;
                }

                var allBtns = btnContainer.querySelectorAll("button");
                allBtns.forEach(function(b) {
                    b.style.background = "#444";
                    b.style.cursor = "not-allowed";
                });
                btn.style.background = "#2d7d46";
                btn.style.cursor = "default";

                countEl.textContent = "Deleted: 0";

                await deleteTemplatesForType(type.key, statusEl, countEl);

                allBtns.forEach(function(b) {
                    b.style.background = "#5b43c7";
                    b.style.cursor = "pointer";
                });
            });

            btnContainer.appendChild(btn);
        });

        container.appendChild(btnContainer);
        container.appendChild(statusEl);
        container.appendChild(countEl);

        DELETE_TEMPLATES_POPUP_REF = createPopup({
            title: "Delete Templates",
            content: container,
            width: "380px",
            height: "auto",
            onClose: function() {
                cancelDeleteTemplates();
                DELETE_TEMPLATES_POPUP_REF = null;
            }
        });

        log("DeleteTemplates: popup opened");
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

        var deleteTemplatesBtn = document.createElement("button");
        deleteTemplatesBtn.textContent = "Delete Templates";
        deleteTemplatesBtn.style.background = "#dc3545";
        deleteTemplatesBtn.style.color = "#fff";
        deleteTemplatesBtn.style.border = "none";
        deleteTemplatesBtn.style.borderRadius = "6px";
        deleteTemplatesBtn.style.padding = "8px";
        deleteTemplatesBtn.style.cursor = "pointer";
        deleteTemplatesBtn.style.fontWeight = "500";
        deleteTemplatesBtn.style.transition = "background 0.2s";
        deleteTemplatesBtn.onmouseenter = function() { this.style.background = "#c82333"; };
        deleteTemplatesBtn.onmouseleave = function() { this.style.background = "#dc3545"; };

        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);
        btnRow.appendChild(deleteTemplatesBtn);
        
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
        deleteTemplatesBtn.addEventListener("click", function () {
            log("Delete Templates: button clicked");
            showDeleteTemplatesPopup();
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
        if (isPaused()) {
            log("Paused; automation halted");
            return;
        }
        var onBarcodeSubjects = isBarcodeSubjectsPage();
        if (onBarcodeSubjects) {
            processBarcodeSubjectsPage();
            return;
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();
