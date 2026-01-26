
// ==UserScript==
// @name ClinSpark Form Automator
// @namespace vinh.activity.plan.state
// @version 1.1.0
// @description
// @match https://cenexeltest.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Form%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/ClinSpark%20Form%20Automator.js
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
    var STORAGE_PENDING = "activityPlanState.pendingIds";
    var STORAGE_AFTER_REFRESH = "activityPlanState.afterRefresh";
    var STORAGE_EDIT_STUDY = "activityPlanState.editStudy";
    var STORAGE_RUN_MODE = "activityPlanState.runMode";
    var STORAGE_CONTINUE_EPOCH = "activityPlanState.continueEpoch";
    var PANEL_ID = "activityPlanStatePanel";
    var LOG_ID = "activityPlanStateLog";
    var STORAGE_SELECTED_IDS = "activityPlanState.selectedVolunteerIds";
    var STORAGE_CONSENT_SCAN_INDEX = "activityPlanState.consent.scanIndex";
    var STORAGE_PAUSED = "activityPlanState.paused";
    var STORAGE_CHECK_ELIG_LOCK = "activityPlanState.checkEligLock";
    var btnRowRef = null;
    var PENDING_BUTTONS = [];
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
    var STORAGE_LOG_VISIBLE = "activityPlanState.log.visible";
    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const PANEL_TOGGLE_KEY = "F2";
    //==========================
    // COLLECT ALL FEATURES
    //==========================
    //==========================

    const DATA_COLLECTION_SUBJECT_URL = "https://cenexeltest.clinspark.com/secure/datacollection/subject";
    var COLLECT_ALL_CANCELLED = false;
    var COLLECT_ALL_POPUP_REF = null;
    var RUN_ALL_POPUP_REF = null;
    var CLEAR_MAPPING_POPUP_REF = null;
    var IMPORT_ELIG_POPUP_REF = null;
    const STORAGE_RUN_ALL_POPUP = "activityPlanState.runAllPopup";
    const STORAGE_RUN_ALL_STATUS = "activityPlanState.runAllStatus";
    const STORAGE_CLEAR_MAPPING_POPUP = "activityPlanState.clearMappingPopup";
    const STORAGE_IMPORT_ELIG_POPUP = "activityPlanState.importEligPopup";

    // Clear all Collect All related data
    function clearCollectAllData() {
        COLLECT_ALL_CANCELLED = false;
        COLLECT_ALL_POPUP_REF = null;
        log("CollectAll: data cleared");
    }

    // Recreate popups on page load if they should be active
    function recreatePopupsIfNeeded() {
        try {
            var runMode = getRunMode();

            // Recreate Run All popup
            if (runMode === "all") {
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
        } catch (e) {
            log("Error recreating popups: " + e);
        }
    }

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


    // Run Barcode feature once from the current page context.
    // Mirrors the behavior of clicking the "8. Run Barcode" button.
    async function APS_RunBarcode() {
        BARCODE_START_TS = Date.now();
        clearBarcodeResult();

        var info = getSubjectFromBreadcrumbOrTooltip();
        var hasText = !!(info.subjectText && info.subjectText.length > 0);
        var hasId = !!(info.subjectId && info.subjectId.length > 0);

        if (!hasText && !hasId) {
            log("Run Barcode: subject breadcrumb or tooltip not found (APS_RunBarcode)");
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

        log("APS_RunBarcode: Opening Barcode Printing Subjects in background…");
        openInTab(location.origin + "/secure/barcodeprinting/subjects", false);

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

        var waited = 0;
        var maxWait = 10000;
        var step = 500;

        while (waited <= maxWait) {
            var r = getBarcodeResult();
            if (r && r.length > 0) {
                try {
                    clearInterval(loadingInterval);
                } catch (e1) {}
                try {
                    if (popup && popup.close) {
                        popup.close();
                    }
                } catch (e2) {}

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

                try {
                    var secs1 = (Date.now() - BARCODE_START_TS) / 1000;
                    var s1 = secs1.toFixed(2);
                    log("APS_RunBarcode: elapsed " + String(s1) + " s");
                } catch (e3) {}
                BARCODE_START_TS = 0;
                clearBarcodeResult();
                return;
            }

            await sleep(step);
            waited = waited + step;
        }

        try {
            clearInterval(loadingInterval);
        } catch (e4) {}
        try {
            if (popup && popup.close) {
                popup.close();
            }
        } catch (e5) {}

        log("APS_RunBarcode: timeout with no result");

        try {
            var secs2 = (Date.now() - BARCODE_START_TS) / 1000;
            var s2 = secs2.toFixed(2);
            log("APS_RunBarcode: elapsed " + String(s2) + " s");
        } catch (e6) {}
        BARCODE_START_TS = 0;
        clearBarcodeResult();
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
                var okClosed = await waitForAnyModalToClose(6000);
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
                // For "< X", generate random number between 0 and X-1
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
                // For "≤ X", generate random number between 0 and X
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
                // For "> X", generate random number between X+1 and X+100 (reasonable upper bound)
                var minVal = t3 + 1;
                var maxVal = t3 + 100; // Reasonable upper bound to avoid infinite ranges
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
                // For "≥ X", generate random number between X and X+100 (reasonable upper bound)
                var minVal = t4;
                var maxVal = t4 + 100; // Reasonable upper bound to avoid infinite ranges
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
            var ok = await radioSelectBestEffort(controlTd, rowId);
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
    async function radioSelectBestEffort(controlTd, rowId) {
        var radios = controlTd.querySelectorAll("div.radio-list input[type=\"radio\"]");
        if (!radios || radios.length === 0) {
            return false;
        }
        var r = radios[0];
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
        // var runPlansBtn = document.createElement("button");
        // runPlansBtn.textContent = "1. Run Activity Plans";
        // runPlansBtn.style.background = "#2d7";
        // runPlansBtn.style.color = "#000";
        // runPlansBtn.style.border = "none";
        // runPlansBtn.style.borderRadius = "6px";
        // runPlansBtn.style.padding = "8px";
        // runPlansBtn.style.cursor = "pointer";
        // var runStudyBtn = document.createElement("button");
        // runStudyBtn.textContent = "3. Run Study Update";
        // runStudyBtn.style.background = "#4af";
        // runStudyBtn.style.color = "#000";
        // runStudyBtn.style.border = "none";
        // runStudyBtn.style.borderRadius = "6px";
        // runStudyBtn.style.padding = "8px";
        // runStudyBtn.style.cursor = "pointer";
        // var runEpochBtn = document.createElement("button");
        // runEpochBtn.textContent = "3. Run Add Cohort Subject";
        // runEpochBtn.style.background = "#fd4";
        // runEpochBtn.style.color = "#000";
        // runEpochBtn.style.border = "none";
        // runEpochBtn.style.borderRadius = "6px";
        // runEpochBtn.style.padding = "8px";
        // runEpochBtn.style.cursor = "pointer";

        // var runAddCohortBtn = document.createElement("button");
        // runAddCohortBtn.textContent = "4. Run Add Cohort Subjects";
        // runAddCohortBtn.style.background = "#fd4";
        // runAddCohortBtn.style.color = "#000";
        // runAddCohortBtn.style.border = "none";
        // runAddCohortBtn.style.borderRadius = "6px";
        // runAddCohortBtn.style.padding = "8px";
        // runAddCohortBtn.style.cursor = "pointer";
        // var runConsentBtn = document.createElement("button");
        // runConsentBtn.textContent = "5. Run ICF Barcode";
        // runConsentBtn.style.background = "#a8f";
        // runConsentBtn.style.color = "#000";
        // runConsentBtn.style.border = "none";
        // runConsentBtn.style.borderRadius = "6px";
        // runConsentBtn.style.padding = "8px";
        // runConsentBtn.style.cursor = "pointer";
        // var runAllBtn = document.createElement("button");
        // runAllBtn.textContent = "6. Run Button (1-5)";
        // runAllBtn.style.background = "#fb6";
        // runAllBtn.style.color = "#000";
        // runAllBtn.style.border = "none";
        // runAllBtn.style.borderRadius = "6px";
        // runAllBtn.style.padding = "8px";
        // runAllBtn.style.cursor = "pointer";
        // var runNonScrnBtn = document.createElement("button");
        // runNonScrnBtn.textContent = "7. Run Import Cohort Subject";
        // runNonScrnBtn.style.background = "#ff7";
        // runNonScrnBtn.style.color = "#000";
        // runNonScrnBtn.style.border = "none";
        // runNonScrnBtn.style.borderRadius = "6px";
        // runNonScrnBtn.style.padding = "8px";
        // runNonScrnBtn.style.cursor = "pointer";
        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "8. Run Barcode";
        runBarcodeBtn.style.background = "#9df";
        runBarcodeBtn.style.color = "#000";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
        var runFormOORBtn = document.createElement("button");
        runFormOORBtn.textContent = "9. Run Form (OOR) B";
        runFormOORBtn.style.background = "#f99";
        runFormOORBtn.style.color = "#000";
        runFormOORBtn.style.border = "none";
        runFormOORBtn.style.borderRadius = "6px";
        runFormOORBtn.style.padding = "8px";
        runFormOORBtn.style.cursor = "pointer";
        var runFormOORABtn = document.createElement("button");
        runFormOORABtn.textContent = "10. Run Form (OOR) A";
        runFormOORABtn.style.background = "#f99";
        runFormOORABtn.style.color = "#000";
        runFormOORABtn.style.border = "none";
        runFormOORABtn.style.borderRadius = "6px";
        runFormOORABtn.style.padding = "8px";
        runFormOORABtn.style.cursor = "pointer";
        var runFormIRBtn = document.createElement("button");
        runFormIRBtn.textContent = "11. Run Form (IR)";
        runFormIRBtn.style.background = "#9f9";
        runFormIRBtn.style.color = "#000";
        runFormIRBtn.style.border = "none";
        runFormIRBtn.style.borderRadius = "6px";
        runFormIRBtn.style.padding = "8px";
        runFormIRBtn.style.cursor = "pointer";

        var collectAllBtn = document.createElement("button");
        collectAllBtn.textContent = "Collect All";
        collectAllBtn.style.background = "#6cf";
        collectAllBtn.style.color = "#000";
        collectAllBtn.style.border = "none";
        collectAllBtn.style.borderRadius = "6px";
        collectAllBtn.style.padding = "8px";
        collectAllBtn.style.cursor = "pointer";

        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#ccc";
        pauseBtn.style.color = "#000";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = "6px";
        pauseBtn.style.padding = "8px";
        pauseBtn.style.cursor = "pointer";

        var clearLogsBtn = document.createElement("button");
        clearLogsBtn.textContent = "Clear Logs";
        clearLogsBtn.style.background = "#555";
        clearLogsBtn.style.color = "#fff";
        clearLogsBtn.style.border = "none";
        clearLogsBtn.style.borderRadius = "6px";
        clearLogsBtn.style.padding = "8px";
        clearLogsBtn.style.cursor = "pointer";

        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#555";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = "6px";
        toggleLogsBtn.style.padding = "8px";
        toggleLogsBtn.style.cursor = "pointer";

        // var runLockSamplePathsBtn = document.createElement("button");
        // runLockSamplePathsBtn.textContent = "2. Run Sample Paths";
        // runLockSamplePathsBtn.style.background = "#f77";
        // runLockSamplePathsBtn.style.color = "#000";
        // runLockSamplePathsBtn.style.border = "none";
        // runLockSamplePathsBtn.style.borderRadius = "6px";
        // runLockSamplePathsBtn.style.padding = "8px";
        // runLockSamplePathsBtn.style.cursor = "pointer";

        // btnRow.appendChild(runPlansBtn);
        // btnRow.appendChild(runLockSamplePathsBtn);
        // btnRow.appendChild(runStudyBtn);
        // btnRow.appendChild(runAddCohortBtn);
        // btnRow.appendChild(runEpochBtn);
        // btnRow.appendChild(runConsentBtn);
        // btnRow.appendChild(runAllBtn);
        // btnRow.appendChild(runNonScrnBtn);
        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(runFormOORBtn);
        btnRow.appendChild(runFormOORABtn);
        btnRow.appendChild(runFormIRBtn);
        btnRow.appendChild(collectAllBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);



        // runLockSamplePathsBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1");
        //     } catch (e) { }

        //     status.textContent = "Navigating to Sample Paths…";
        //     log("Run Lock Sample Paths clicked");
        //     location.href = "https://cenexeltest.clinspark.com/secure/samples/configure/paths";
        // });


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
        // runNonScrnBtn.addEventListener("click", function () {
        //     status.textContent = "Preparing non-SCRN subject import...";
        //     log("Run Add non-SCRN Subject clicked");
        //     localStorage.setItem(STORAGE_RUN_MODE, "nonscrn");
        //     location.href = STUDY_SHOW_URL + "?autononscrn=1";
        // });
        // runAddCohortBtn.addEventListener("click", function () {
        //     status.textContent = "Preparing Add Cohort Subjects...";
        //     log("Run Add Cohort Subjects clicked");
        //     try {
        //         localStorage.setItem(STORAGE_RUN_MODE, "epochAddCohort");
        //         localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
        //         log("Cleared cohortAdd.editDoneMap for new run");
        //     } catch (e) {}
        //     location.href = STUDY_SHOW_URL + "?autoaddcohort=1";
        // });
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
        // runPlansBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_KEY, "1");
        //         localStorage.setItem(STORAGE_RUN_MODE, "activity");
        //         localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
        //         localStorage.removeItem(STORAGE_AFTER_REFRESH);
        //     } catch (e) {}
        //     status.textContent = "Opening Activity Plans…";
        //     log("Run Activity Plans clicked");
        //     location.href = LIST_URL;
        // });
        // runStudyBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_RUN_MODE, "study");
        //         localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
        //     } catch (e) {}
        //     status.textContent = "Navigating to Study Show…";
        //     log("Run Study Update clicked");
        //     location.href = STUDY_SHOW_URL + "?autoupdate=1";
        // });
        // runEpochBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_RUN_MODE, "epoch");
        //     } catch (e) {}
        //     try {
        //         localStorage.setItem(STORAGE_CONTINUE_EPOCH, "1");
        //     } catch (e2) {}
        //     status.textContent = "Navigating to Study Show for Cohort Add…";
        //     log("Run Cohort Add clicked");
        //     location.href = STUDY_SHOW_URL;
        //     clearCohortGuard();
        // });
        // runConsentBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_RUN_MODE, "consent");
        //     } catch (e) {}
        //     status.textContent = "Navigating to Study Show for Consent…";
        //     log("Run Informed Consent clicked");
        //     location.href = STUDY_SHOW_URL + "?autoconsent=1";
        // });
        // runAllBtn.addEventListener("click", function () {
        //     try {
        //         localStorage.setItem(STORAGE_KEY, "1");
        //         localStorage.setItem(STORAGE_RUN_MODE, "all");
        //         localStorage.setItem(STORAGE_CONTINUE_EPOCH, "1");
        //         localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
        //     } catch (e) {}
        //     status.textContent = "Starting ALL: Activity Plans…";
        //     log("Run ALL clicked");
        //     location.href = LIST_URL;
        //     clearCohortGuard();
        // });
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
        // var isSamplePathsPage = location.pathname === "/secure/samples/configure/paths";
        // if (isSamplePathsPage) {
        //     processLockSamplePathsPage();
        //     return;
        // }

        // var isSamplePathDetail = location.pathname.indexOf("/secure/samples/configure/paths/show/") !== -1;
        // if (isSamplePathDetail) {
        //     processLockSamplePathDetailPage();
        //     return;
        // }

        // var isSamplePathUpdate = location.pathname.indexOf("/secure/samples/configure/paths/update/") !== -1;
        // if (isSamplePathUpdate) {
        //     processLockSamplePathUpdatePage();
        //     return;
        // }

        // var onList = isListPage();
        // if (onList) {
        //     processListPage();
        //     return;
        // }
        // var onShow = isShowPage();
        // if (onShow) {
        //     processShowPageIfAuto();
        //     return;
        // }
        // var onStudyShow = isStudyShowPage();
        // if (onStudyShow) {
        //     processStudyShowPage();
        //     processStudyShowPageForNonScrn();
        //     processStudyShowPageForAddCohort();
        //     return;
        // }
        // var onEditBasics = isStudyEditBasicsPage();
        // if (onEditBasics) {
        //     processStudyEditBasicsPageIfFlag();
        //     return;
        // }

        // var onEpoch = isEpochShowPage();
        // if (onEpoch) {
        //     var imode = getRunMode();
        //     if (imode === "epochImport") {
        //         processEpochShowPageForImport();
        //         return;
        //     }
        //     if (imode === "epochAddCohort") {
        //         processEpochShowPageForAddCohort();
        //         return;
        //     }
        //     processEpochShowPage();
        //     return;
        // }

        // var onCohort = isCohortShowPage();
        // if (onCohort) {
        //     var amode = getRunMode();
        //     var autoImport = getQueryParam("autocohortimport");
        //     var autoAdd = getQueryParam("autocohortadd");
        //     if (amode === "epochImport" || autoImport === "1") {
        //         processCohortShowPageImportNonScrn();
        //     } else if (amode === "epochAddCohort" || autoAdd === "1") {
        //         processCohortShowPageAddCohort();
        //     } else {
        //         processCohortShowPage();
        //     }
        //     return;
        // }
        // var onMetadata = isStudyMetadataPage();
        // if (onMetadata) {
        //     processStudyMetadataPageForEligibilityLock();
        //     return;
        // }
        // var onSubjectsList = isSubjectsListPage();
        // if (onSubjectsList) {
        //     processSubjectsListPageForConsent();
        //     return;
        // }
        // var onSubjectShow = isSubjectShowPage();
        // if (onSubjectShow) {
        //     processSubjectShowPageForConsent();
        //     return;
        // }
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