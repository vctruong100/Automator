
// ==UserScript==
// @name        ClinSpark Automator
// @namespace   vinh.activity.plan.state
// @version     1.3.6
// @description Retain only Barcode feature; production environment only
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
    var PANEL_DEFAULT_WIDTH = "340px";
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 48;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 600;
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


    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const PANEL_TOGGLE_KEY = "F2";

    const ELIGIBILITY_LIST_URL = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";

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
    function clearBarcodeResult() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_RESULT);
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
        return s;
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


    function continueFindForm(keyword, subjectInput) {
        var url = FORM_LIST_URL;
        log("Find Form: opening list url='" + String(url) + "'");
        var w = window;
        w.location.href = url;
        var t = 0;
        var r = setInterval(function () {
            t = t + 1;
            if (!w || w.closed) {
                clearInterval(r);
                log("Find Form: child window closed");
                return;
            }
            if (t > 120) {
                clearInterval(r);
                log("Find Form: timeout waiting for child page");
                return;
            }
            var d = w.document;
            if (!d) {
                log("Find Form: child document not ready at t=" + String(t));
                return;
            }
            var subjNorm = aeNormalize(subjectInput || "");
            var y = "";
            if (subjNorm && subjNorm.length > 0) {
                var os = d.querySelectorAll('#subjectIds option,select[name="subjectIds"] option');
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
                        var matchById = val === subjectInput;
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
                            log("Find Form: subject matched optionIndex=" + String(j) + " text='" + String(txt) + "' value='" + String(val) + "'");
                            break;
                        }
                        j = j + 1;
                    }
                } else {
                    log("Find Form: subject options not found");
                }
            } else {
                log("Find Form: no subject input provided; skipping subject selection");
            }

            var formsSel = d.querySelector('select#formIds, select[name="formIds"]');
            var selectedCount = 0;
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
            } else {
                log("Find Form: forms select not found");
            }

            if (selectedCount === 0) {
                clearInterval(r);
                log("Find Form: no forms selected; closing child and notifying user");
                showFormNoMatchPopup();
                try {
                    w.close();
                } catch (e) {}
                return;
            }

            if (y && y.length > 0) {
                var subjSel = d.querySelector('select#subjectIds, select[name="subjectIds"]');
                if (subjSel) {
                    subjSel.value = y;
                    var evtSubj = new Event("change", { bubbles: true });
                    subjSel.dispatchEvent(evtSubj);
                    log("Find Form: subject applied value='" + String(y) + "'");
                } else {
                    log("Find Form: subject select not found");
                }
            }

            var searchBtn = d.getElementById("dataSearchButton");
            if (searchBtn) {
                searchBtn.click();
                log("Find Form: Search button clicked");
            } else {
                log("Find Form: Search button not found");
            }

            clearInterval(r);
            log("Find Form: inputs populated and search initiated; leaving child page open");
        }, 200);
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


    function aeNormalize(x) {
        var s = x || "";
        s = s.replace(/\s+/g, "");
        s = s.replace(/›|»|▶|►|→/g, "");
        s = s.trim();
        s = s.toUpperCase();
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


    // ================
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
            w = PANEL_DEFAULT_WIDTH;
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
                newW = PANEL_MAX_WIDTH_PX;
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

    function updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar) {
        var collapsed = getPanelCollapsed();
        if (collapsed) {
            panel.style.width = PANEL_DEFAULT_WIDTH;
            panel.style.height = String(PANEL_HEADER_HEIGHT_PX + 12) + "px";
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
        panel.style.borderRadius = "8px";
        panel.style.padding = "0";
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
        headerBar.style.padding = "0 12px";
        var leftSpacer = document.createElement("div");
        leftSpacer.style.width = "32px";
        var title = document.createElement("div");
        title.textContent = "ClinSpark Automator";
        title.style.fontWeight = "600";
        title.style.textAlign = "center";
        title.style.justifySelf = "center";
        title.style.transform = "translateX(16px)";
        headerBar.appendChild(leftSpacer);
        headerBar.appendChild(title);
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
        bodyContainer.style.padding = "12px";
        var btnRow = document.createElement("div");
        btnRow.style.display = "grid";
        btnRow.style.gridTemplateColumns = "1fr 1fr";
        btnRow.style.gap = "8px";
        btnRowRef = btnRow;
        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#6c757d";
        pauseBtn.style.color = "#fff";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = "6px";
        pauseBtn.style.padding = "8px";
        pauseBtn.style.cursor = "pointer";
        pauseBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        pauseBtn.onmouseleave = function() { this.style.background = "#6c757d"; };
        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Run Barcode";
        runBarcodeBtn.style.background = "#4a90e2";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        runBarcodeBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var findAeBtn = document.createElement("button");
        findAeBtn.textContent = "Find Adverse Event";
        findAeBtn.style.background = "#4a90e2";
        findAeBtn.style.color = "#fff";
        findAeBtn.style.border = "none";
        findAeBtn.style.borderRadius = "6px";
        findAeBtn.style.padding = "8px";
        findAeBtn.style.cursor = "pointer";
        findAeBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        findAeBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var findFormBtn = document.createElement("button");
        findFormBtn.textContent = "Find Form";
        findFormBtn.style.background = "#4a90e2";
        findFormBtn.style.color = "#fff";
        findFormBtn.style.border = "none";
        findFormBtn.style.borderRadius = "6px";
        findFormBtn.style.padding = "8px";
        findFormBtn.style.cursor = "pointer";
        findFormBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        findFormBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#6c757d";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = "6px";
        toggleLogsBtn.style.padding = "8px";
        toggleLogsBtn.style.cursor = "pointer";
        toggleLogsBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = function() { this.style.background = "#6c757d"; }; 
        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import I/E";
        importEligBtn.style.background = "#4a90e2";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = "6px";
        importEligBtn.style.padding = "8px";
        importEligBtn.style.cursor = "pointer";
        importEligBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        importEligBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(findAeBtn);
        btnRow.appendChild(findFormBtn);
        btnRow.appendChild(importEligBtn);
        btnRow.appendChild(pauseBtn);
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
        }        importEligBtn.addEventListener("click", function () {
            log("ImportElig: button clicked");
            startImportEligibilityMapping();
        });
        importEligBtn.addEventListener("click", function () {
            log("ImportElig: button clicked");
            startImportEligibilityMapping();
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
            BARCODE_BG_TAB = openInTab(location.origin + "/secure/barcodeprinting/subjects", false);
            log("Run Barcode: background tab opened");
            var loadingText = document.createElement("div");
            loadingText.style.textAlign = "center";
            loadingText.style.fontSize = "16px";
            loadingText.style.color = "#fff";
            loadingText.style.padding = "20px";
            loadingText.textContent = "Locating barcode.";
            var canceled = false;
            var popup = createPopup({
                title: "Locating Barcode",
                content: loadingText,
                width: "300px",
                height: "auto",
                onClose: function () {
                    canceled = true;
                    try {
                        if (BARCODE_BG_TAB && typeof BARCODE_BG_TAB.close === "function") {
                            BARCODE_BG_TAB.close();
                            log("Run Barcode: background tab closed due to popup close");
                        }
                    } catch (e) {
                        log("Run Barcode: error closing background tab " + String(e));
                    }
                    BARCODE_BG_TAB = null;
                    clearBarcodeResult();
                    BARCODE_START_TS = 0;
                    log("Run Barcode: canceled by user; stopping");
                }
            });
            var dots = 1;
            var loadingInterval = setInterval(function () {
                dots = dots + 1;
                if (dots > 3) {
                    dots = 1;
                }
                var text = "Locating barcode";
                var i2 = 0;
                while (i2 < dots) {
                    text = text + ".";
                    i2 = i2 + 1;
                }
                loadingText.textContent = text;
            }, 500);
            var waited = 0;
            var maxWait = 15000;
            var intervalMs = 500;
            var timer = setInterval(async function () {
                if (canceled) {
                    clearInterval(loadingInterval);
                    clearInterval(timer);
                    return;
                }
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
                    try {
                        if (BARCODE_BG_TAB && typeof BARCODE_BG_TAB.close === "function") {
                            BARCODE_BG_TAB.close();
                            log("Run Barcode: background tab closed after success");
                        }
                    } catch (e) {
                        log("Run Barcode: error closing background tab " + String(e));
                    }
                    BARCODE_BG_TAB = null;
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
                        try {
                            if (BARCODE_BG_TAB && typeof BARCODE_BG_TAB.close === "function") {
                                BARCODE_BG_TAB.close();
                                log("Run Barcode: background tab closed after timeout");
                            }
                        } catch (e2) {
                            log("Run Barcode: error closing background tab " + String(e2));
                        }
                        BARCODE_BG_TAB = null;
                        var secs2 = (Date.now() - BARCODE_START_TS) / 1000;
                        var s2 = secs2.toFixed(2);
                        log("Run Barcode: elapsed " + String(s2) + " s");
                        BARCODE_START_TS = 0;
                    }
                }
            }, intervalMs);
        });
        findAeBtn.addEventListener("click", function () {
            openAndLocateAdverseEvent();
        });
        findFormBtn.addEventListener("click", function () {
            openFindForm();
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
            processFindFormOnList();
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

