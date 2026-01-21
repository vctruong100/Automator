
// ==UserScript==
// @name ClinSpark Test Automator
// @namespace vinh.activity.plan.state
// @version 2.5.3
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

    //==========================
    // CLEAR SUBJECT ELIGIBILITY FEATURE
    //==========================
    // This section contains all functions related to clearing subject eligibility.
    // This feature automates clearing all existing eligibility mapping in the table.
    //==========================

    function ClearEligibilityFunctions() {}
    const RUNMODE_CLEAR_MAPPING = "clearMapping";

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
        planPriDefaultBtn.textContent = "Use Default";
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
        ignoreDefaultBtn.textContent = "Use Default";
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
            title: "Import Eligibility Mapping",
            content: container,
            width: "450px",
            height: "auto",
            onClose: function () {
                // X pressed → STOP everything (but don't pause)
                log("ImportElig: popup X pressed → stopping automation");
                clearAllRunState();
                clearEligibilityWorkingState();
            }
        });

        confirmBtn.addEventListener("click", function () {
            log("ImportElig: Confirm clicked");

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



    async function startImportEligibilityMapping() {

        if (isPaused()) {
            return;
        }

        try {
            localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ELIG_IMPORT);
        } catch (e) {
        }

        if (!isEligibilityListPage()) {
            location.href = ELIGIBILITY_LIST_URL;
            return;
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
            var opened = await openAddEligibilityModal();
            if (!opened) {
                log("ImportElig: cannot open modal; stopping");
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
                var v1 = t1 - 1;
                if (v1 >= 0) {
                    return v1;
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
                return t2;
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
                return t3 + 1;
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
                return t4;
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

        var i = 0;
        while (i < rows.length) {
            var row = rows[i];

            var tds = row.querySelectorAll("td");
            var tdCount = tds.length;

            if (tdCount >= 4) {
                var lockCell = tds[3];
                var lockText = (lockCell.textContent + "").trim();

                log("Row " + String(i) + " lockText=" + lockText);

                if (lockText.toLowerCase() === "no") {
                    var link = tds[0].querySelector("a");

                    if (link) {
                        var href = link.getAttribute("href") + "";
                        if (href.length > 0) {
                            var fullUrl = location.origin + href;
                            log("Opening unlocked Sample Path: " + fullUrl);
                            openInTab(fullUrl + "?autolocksamplepath=1", false);
                        }
                    }
                }
            }

            i = i + 1;
        }

        log("processLockSamplePathsPage: all unlocked paths opened");

        try {
            localStorage.removeItem(STORAGE_RUN_LOCK_SAMPLE_PATHS);
            log("processLockSamplePathsPage: flag cleared");
        } catch (e) { }

        var mode = getRunMode();
        if (mode === "all") {
            await sleep(1000);
            log("processLockSamplePathsPage: continuing to Study Show for ALL mode");
            location.href = STUDY_SHOW_URL + "?autoupdate=1";
        }
    }

    async function processLockSamplePathDetailPage() {
        log("processLockSamplePathDetailPage: start");

        var successAlert = document.querySelector('div.alert.alert-success.alert-dismissable');
        if (successAlert) {
            var alertText = (successAlert.textContent + "").trim();
            if (alertText.indexOf("The sample path has been updated") !== -1) {
                log("processLockSamplePathDetailPage: success alert detected, closing tab");
                await sleep(300);
                try {
                    window.close();
                } catch (e) {
                    log("Error closing window: " + String(e));
                }
                setTimeout(function() {
                    try { window.close(); } catch (e2) {}
                }, 500);
                setTimeout(function() {
                    try { window.close(); } catch (e3) {}
                }, 2000);
                return;
            }
        }

        var auto = getQueryParam("autolocksamplepath");
        var isAuto = auto === "1";

        if (!isAuto) {
            log("processLockSamplePathDetailPage: not auto mode");

            var checkSuccess = setInterval(function() {
                var alert = document.querySelector('div.alert.alert-success.alert-dismissable');
                if (alert) {
                    var text = (alert.textContent + "").trim();
                    if (text.indexOf("The sample path has been updated") !== -1) {
                        clearInterval(checkSuccess);
                        log("processLockSamplePathDetailPage: success alert detected, closing tab");
                        setTimeout(function() {
                            try { window.close(); } catch (e) {}
                        }, 300);
                        setTimeout(function() {
                            try { window.close(); } catch (e) {}
                        }, 800);
                        setTimeout(function() {
                            try { window.close(); } catch (e) {}
                        }, 2300);
                    }
                }
            }, 500);

            setTimeout(function() {
                clearInterval(checkSuccess);
            }, 5000);

            return;
        }

        var checkSuccessInterval = setInterval(function() {
            var alert = document.querySelector('div.alert.alert-success.alert-dismissable');
            if (alert) {
                var text = (alert.textContent + "").trim();
                if (text.indexOf("The sample path has been updated") !== -1) {
                    clearInterval(checkSuccessInterval);
                    log("processLockSamplePathDetailPage: success alert detected during processing, closing tab");
                    setTimeout(function() {
                        try { window.close(); } catch (e) {}
                    }, 300);
                    setTimeout(function() {
                        try { window.close(); } catch (e) {}
                    }, 800);
                    setTimeout(function() {
                        try { window.close(); } catch (e) {}
                    }, 2300);
                }
            }
        }, 500);

        var closeTimeout = setTimeout(function() {
            clearInterval(checkSuccessInterval);
            log("processLockSamplePathDetailPage: timeout, closing tab");
            try { window.close(); } catch (e) {}
        }, 5000);

        var actionsOpened = await clickActionsDropdownIfNeeded();
        await sleep(500);
        log("Actions dropdown opened=" + String(actionsOpened));

        if (!actionsOpened) {
            log("Actions dropdown not found");
            clearTimeout(closeTimeout);
            clearInterval(checkSuccessInterval);
            await sleep(1000);
            try { window.close(); } catch (e) {}
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }

        var editLink = await waitForSelector('a[href*="/secure/samples/configure/paths/update/"]', 5000);
        if (!editLink) {
            editLink = document.querySelector('a[href*="/secure/samples/configure/paths/update/"]');
        }
        var hasEditLink = !!editLink;

        log("Edit Path link exists=" + String(hasEditLink));

        if (!hasEditLink) {
            log("Edit Path link missing");
            clearTimeout(closeTimeout);
            clearInterval(checkSuccessInterval);
            await sleep(1000);
            try { window.close(); } catch (e) {}
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }

        clearTimeout(closeTimeout);

        var href = editLink.getAttribute("href") + "";
        var updateUrl = location.origin + href;
        if (href.indexOf("?") === -1) {
            updateUrl = updateUrl + "?autolocksamplepath=1";
        } else {
            updateUrl = updateUrl + "&autolocksamplepath=1";
        }
        log("Edit Path clicked, navigating to: " + updateUrl);

        clearInterval(checkSuccessInterval);
        location.href = updateUrl;
    }

    async function processLockSamplePathUpdatePage() {
        log("processLockSamplePathUpdatePage: start");

        var auto = getQueryParam("autolocksamplepath");
        var isAuto = auto === "1";

        if (!isAuto) {
            log("processLockSamplePathUpdatePage: not auto mode");
            setTimeout(function() {
                try { window.close(); } catch (e) {}
            }, 5000);
            return;
        }

        var closeTimeout = setTimeout(function() {
            log("processLockSamplePathUpdatePage: timeout, closing tab");
            try { window.close(); } catch (e) {}
        }, 5000);

        var lockBox = await waitForSelector('input#locked', 10000);
        var lockBoxExists = !!lockBox;
        log("Lock checkbox exists=" + String(lockBoxExists));

        if (lockBox) {
            if (!lockBox.checked) {
                lockBox.checked = true;
                var evtL = new Event("change", { bubbles: true });
                lockBox.dispatchEvent(evtL);

                var wrapper = lockBox.closest("div.checker");
                if (wrapper) {
                    var span = wrapper.querySelector("span");
                    if (span) {
                        span.classList.add("checked");
                    }
                }

                log("Lock checkbox set to true");
            } else {
                log("Lock checkbox already checked");
            }
        } else {
            log("Lock checkbox not found");
            clearTimeout(closeTimeout);
            await sleep(1000);
            try { window.close(); } catch (e) {}
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }

        var reason = await waitForSelector("textarea#reasonForChange", 10000);
        var hasReason = !!reason;
        log("Reason textarea exists=" + String(hasReason));

        if (reason) {
            reason.value = "Locking...";
            var evtR = new Event("input", { bubbles: true });
            reason.dispatchEvent(evtR);
            log("Reason for Change set");
        } else {
            log("Reason textarea not found");
            clearTimeout(closeTimeout);
            await sleep(1000);
            try { window.close(); } catch (e) {}
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }

        var saveBtn = await waitForSelector("button.btn.green[type='submit']", 10000);
        var hasSave = !!saveBtn;
        log("Save button exists=" + String(hasSave));

        if (saveBtn) {
            clearTimeout(closeTimeout);
            saveBtn.click();
            log("Save clicked");

            var successDetected = false;
            var checkStart = Date.now();
            var checkMax = 15000;
            var checkInterval = 300;

            while (Date.now() - checkStart < checkMax && !successDetected) {
                await sleep(checkInterval);

                var successAlert = document.querySelector('div.alert.alert-success.alert-dismissable');
                if (successAlert) {
                    var alertText = (successAlert.textContent + "").trim();
                    if (alertText.indexOf("The sample path has been updated") !== -1) {
                        log("processLockSamplePathUpdatePage: success alert detected, closing tab");
                        successDetected = true;
                        await sleep(300);
                        try {
                            window.close();
                        } catch (e) {
                            log("Error closing window: " + String(e));
                        }
                        setTimeout(function() {
                            try { window.close(); } catch (e2) {}
                        }, 500);
                        setTimeout(function() {
                            try { window.close(); } catch (e3) {}
                        }, 2000);
                        return;
                    }
                }

                var currentPath = location.pathname;
                if (currentPath.indexOf("/secure/samples/configure/paths/show/") !== -1) {
                    var showPageAlert = document.querySelector('div.alert.alert-success.alert-dismissable');
                    if (showPageAlert) {
                        var showAlertText = (showPageAlert.textContent + "").trim();
                        if (showAlertText.indexOf("The sample path has been updated") !== -1) {
                            log("processLockSamplePathUpdatePage: success alert detected on show page, closing tab");
                            successDetected = true;
                            await sleep(300);
                            try {
                                window.close();
                            } catch (e) {
                                log("Error closing window: " + String(e));
                            }
                            setTimeout(function() {
                                try { window.close(); } catch (e2) {}
                            }, 500);
                            setTimeout(function() {
                                try { window.close(); } catch (e3) {}
                            }, 2000);
                            return;
                        }
                    }
                }
            }

            if (!successDetected) {
                log("processLockSamplePathUpdatePage: success alert not detected within timeout, closing tab anyway");
                await sleep(300);
                try {
                    window.close();
                } catch (e) {
                    log("Error closing window: " + String(e));
                }
                setTimeout(function() {
                    try { window.close(); } catch (e2) {}
                }, 500);
                setTimeout(function() {
                    try { window.close(); } catch (e3) {}
                }, 2000);
            }
        } else {
            log("Save button not found");
            clearTimeout(closeTimeout);
            await sleep(1000);
            try {
                window.close();
            } catch (e) {
                log("Error closing window: " + String(e));
            }
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }
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

    // Pick a random unused letter from a-z.
    function pickRandomUnusedLetter(allLetters, usedLetters) {
        var unused = [];
        var i = 0;
        while (i < allLetters.length) {
            var letter = allLetters[i];
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
            btn.style.background = "#2d7";
            btn.style.color = "#000";
            btn.style.border = "none";
            btn.style.borderRadius = "6px";
            btn.style.fontSize = "14px";
            btn.style.fontWeight = "500";
            btn.style.transition = "background 0.2s";

            btn.addEventListener("mouseenter", function() {
                btn.style.background = "#3e8";
            });
            btn.addEventListener("mouseleave", function() {
                btn.style.background = "#2d7";
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
        clearRunMode();
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
            btn.style.background = "#2d7";
            btn.style.color = "#000";
            btn.style.border = "none";
            btn.style.borderRadius = "6px";
            btn.style.fontSize = "14px";
            btn.style.fontWeight = "500";
            btn.style.transition = "background 0.2s";

            btn.addEventListener("mouseenter", function() {
                btn.style.background = "#3e8";
            });
            btn.addEventListener("mouseleave", function() {
                btn.style.background = "#2d7";
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
                    try {
                        localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
                    } catch (e3) {}
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
            try {
                localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1");
            } catch (e5) {}
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
        var links = getPlanLinks();
        log("Found plan links=" + String(links.length));
        var ids = [];
        var i = 0;
        while (i < links.length) {
            var url = links[i];
            openInTab(url, false);
            var id = extractAutostateIdFromUrl(url);
            if (id) {
                ids.push(id);
            }
            await sleep(200);
            i = i + 1;
        }
        if (ids.length > 0) {
            setPendingIds(ids);
            monitorCompletionThenAdvance();
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
            setContinueEpoch();
            log("Eligibility locked; continuing ALL to Study Show");
            location.href = "/secure/administration/studies/show";
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
    async function processShowPageIfAuto() {
        var autoVal = getQueryParam("autostate");
        var doAuto = !!autoVal;
        if (!doAuto) {
            setTimeout(function() {
                try { window.close(); } catch (e) {}
            }, 5000);
            return;
        }
        var mode = getRunMode();
        var id = getCurrentPlanId();
        var ids = getPendingIds();
        var isPending = false;
        var k = 0;
        while (k < ids.length) {
            var v = String(ids[k]);
            if (id && String(id) === v) {
                isPending = true;
                break;
            }
            k = k + 1;
        }
        if (!(mode === "activity" || isPending)) {
            log("processShowPageIfAuto ignored; run not active");
            setTimeout(function() {
                try { window.close(); } catch (e) {}
            }, 5000);
            return;
        }

        var closeTimeout = setTimeout(function() {
            log("processShowPageIfAuto: timeout, closing tab");
            try { window.close(); } catch (e) {}
        }, 5000);

        log("processShowPageIfAuto autostate=" + String(autoVal));
        var opened = await findAndOpenEditStateModal();
        if (!opened) {
            clearTimeout(closeTimeout);
            await sleep(1000);
            try { window.close(); } catch (e) {}
            setTimeout(function() {
                try { window.close(); } catch (e2) {}
            }, 500);
            return;
        }
        var saved = await clickSaveInModal();
        var id2 = getCurrentPlanId();
        if (id2) {
            removePendingId(id2);
        }
        clearTimeout(closeTimeout);
        await sleep(300);
        try {
            window.close();
        } catch (e) {
            log("Error closing window: " + String(e));
        }
        setTimeout(function() {
            try { window.close(); } catch (e2) {}
        }, 500);
        setTimeout(function() {
            try { window.close(); } catch (e3) {}
        }, 2000);
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

    // Detect if current path is a Study Library form-show page.
    function isStudyLibraryFormShowPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/crfdesign/studylibrary/show/form/") !== -1;
        if (ok) {
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
        runPlansBtn.textContent = "1. Lock Activity Plans";
        runPlansBtn.style.background = "#2d7";
        runPlansBtn.style.color = "#000";
        runPlansBtn.style.border = "none";
        runPlansBtn.style.borderRadius = "6px";
        runPlansBtn.style.padding = "8px";
        runPlansBtn.style.cursor = "pointer";
        var runStudyBtn = document.createElement("button");
        runStudyBtn.textContent = "3. Update Study Status";
        runStudyBtn.style.background = "#4af";
        runStudyBtn.style.color = "#000";
        runStudyBtn.style.border = "none";
        runStudyBtn.style.borderRadius = "6px";
        runStudyBtn.style.padding = "8px";
        runStudyBtn.style.cursor = "pointer";
        var runAddCohortBtn = document.createElement("button");
        runAddCohortBtn.textContent = "4. Add Cohort Subjects";
        runAddCohortBtn.style.background = "#fd4";
        runAddCohortBtn.style.color = "#000";
        runAddCohortBtn.style.border = "none";
        runAddCohortBtn.style.borderRadius = "6px";
        runAddCohortBtn.style.padding = "8px";
        runAddCohortBtn.style.cursor = "pointer";
        var runConsentBtn = document.createElement("button");
        runConsentBtn.textContent = "5. Run ICF Barcode";
        runConsentBtn.style.background = "#a8f";
        runConsentBtn.style.color = "#000";
        runConsentBtn.style.border = "none";
        runConsentBtn.style.borderRadius = "6px";
        runConsentBtn.style.padding = "8px";
        runConsentBtn.style.cursor = "pointer";
        var runAllBtn = document.createElement("button");
        runAllBtn.textContent = "6. Run Button (1-5)";
        runAllBtn.style.background = "#fb6";
        runAllBtn.style.color = "#000";
        runAllBtn.style.border = "none";
        runAllBtn.style.borderRadius = "6px";
        runAllBtn.style.padding = "8px";
        runAllBtn.style.cursor = "pointer";
        var runNonScrnBtn = document.createElement("button");
        runNonScrnBtn.textContent = "7. Import Cohort Subject";
        runNonScrnBtn.style.background = "#ff7";
        runNonScrnBtn.style.color = "#000";
        runNonScrnBtn.style.border = "none";
        runNonScrnBtn.style.borderRadius = "6px";
        runNonScrnBtn.style.padding = "8px";
        runNonScrnBtn.style.cursor = "pointer";
        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "8. Run Barcode";
        runBarcodeBtn.style.background = "#9df";
        runBarcodeBtn.style.color = "#000";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
        var runFormOORBtn = document.createElement("button");
        runFormOORBtn.textContent = "9. Run Form (OOR) Below Range";
        runFormOORBtn.style.background = "#f99";
        runFormOORBtn.style.color = "#000";
        runFormOORBtn.style.border = "none";
        runFormOORBtn.style.borderRadius = "6px";
        runFormOORBtn.style.padding = "8px";
        runFormOORBtn.style.cursor = "pointer";
        var runFormOORABtn = document.createElement("button");
        runFormOORABtn.textContent = "10. Run Form (OOR) Above Range";
        runFormOORABtn.style.background = "#f99";
        runFormOORABtn.style.color = "#000";
        runFormOORABtn.style.border = "none";
        runFormOORABtn.style.borderRadius = "6px";
        runFormOORABtn.style.padding = "8px";
        runFormOORABtn.style.cursor = "pointer";
        var runFormIRBtn = document.createElement("button");
        runFormIRBtn.textContent = "11. Run Form (In Range)";
        runFormIRBtn.style.background = "#9f9";
        runFormIRBtn.style.color = "#000";
        runFormIRBtn.style.border = "none";
        runFormIRBtn.style.borderRadius = "6px";
        runFormIRBtn.style.padding = "8px";
        runFormIRBtn.style.cursor = "pointer";
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

        var runLockSamplePathsBtn = document.createElement("button");
        runLockSamplePathsBtn.textContent = "2. Lock Sample Paths";
        runLockSamplePathsBtn.style.background = "#f77";
        runLockSamplePathsBtn.style.color = "#000";
        runLockSamplePathsBtn.style.border = "none";
        runLockSamplePathsBtn.style.borderRadius = "6px";
        runLockSamplePathsBtn.style.padding = "8px";
        runLockSamplePathsBtn.style.cursor = "pointer";

        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import Eligibility Mapping";
        importEligBtn.style.background = "#9df";
        importEligBtn.style.color = "#000";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = "6px";
        importEligBtn.style.padding = "8px";
        importEligBtn.style.cursor = "pointer";


        var clearMappingBtn = document.createElement("button");
        clearMappingBtn.textContent = "Clear Mapping";
        clearMappingBtn.style.background = "#f99";
        clearMappingBtn.style.color = "#000";
        clearMappingBtn.style.border = "none";
        clearMappingBtn.style.borderRadius = "6px";
        clearMappingBtn.style.padding = "8px";
        clearMappingBtn.style.cursor = "pointer";

        btnRow.appendChild(runPlansBtn);
        btnRow.appendChild(runLockSamplePathsBtn);
        btnRow.appendChild(runStudyBtn);
        btnRow.appendChild(runAddCohortBtn);
        btnRow.appendChild(runConsentBtn);
        btnRow.appendChild(runAllBtn);
        btnRow.appendChild(runNonScrnBtn);
        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(runFormOORBtn);
        btnRow.appendChild(runFormOORABtn);
        btnRow.appendChild(runFormIRBtn);
        btnRow.appendChild(importEligBtn);
        btnRow.appendChild(clearMappingBtn);
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
        clearMappingBtn.addEventListener("click", function () {
            log("ClearMapping: button clicked");
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
                var runModeRaw = null;
        try {
            runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            runModeRaw = null;
        }
        if (runModeRaw === RUNMODE_ELIG_IMPORT) {
            if (!isEligibilityListPage()) {
                log("ImportElig: run mode set but not on list page; redirecting to " + ELIGIBILITY_LIST_URL);
                location.href = ELIGIBILITY_LIST_URL;
                return;
            }
            log("ImportElig: run mode set on list page; waiting 3s before resuming");
            setTimeout(function () {
                executeEligibilityMappingAutomation();
            }, 3000);
            return;
        }
        log("Init: ready");


        if (runModeRaw === RUNMODE_CLEAR_MAPPING) {
            if (!isEligibilityListPage()) {
                log("ClearMapping: run mode set but not on list page; redirecting");
                location.href = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
                return;
            }
            log("ClearMapping: run mode set on list page; resuming in 4s");
            setTimeout(function () {
                executeClearMappingAutomation();
            }, 4000);
            return;
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();
