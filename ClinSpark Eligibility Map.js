
// ==UserScript==
// @name ClinSpark Eligibility Mapping Automator
// @namespace vinh.activity.plan.state
// @version 1.4.0
// @description
// @match https://cenexeltest.clinspark.com/*
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
    var BARCODE_START_TS = 0;
    var STORAGE_LOG_VISIBLE = "activityPlanState.log.visible";
    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const PANEL_TOGGLE_KEY = "F2";

    const ELIGIBILITY_LIST_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const ELIGIBILITY_LIST_URL_PROD = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const ELIGIBILITY_VALID_HOSTNAMES = ["cenexeltest.clinspark.com", "cenexel.clinspark.com"];
    const ELIGIBILITY_LIST_PATH = "/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";
    const IE_CODE_REGEX = /\b(INC|EXC)\s*(\d+)\b/i;
    const IE_CODE_REGEX_GLOBAL = /\b(INC|EXC)\s*(\d+)\b/gi;
    const IMPORT_IE_HELPER_TIMEOUT = 15000;
    const IMPORT_IE_POLL_INTERVAL = 120;
    const IMPORT_IE_MODAL_TIMEOUT = 12000;
    const IMPORT_IE_SHORT_DELAY_MIN = 150;
    const IMPORT_IE_SHORT_DELAY_MAX = 400;
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
            location.href = getBaseUrl() + ELIGIBILITY_LIST_PATH;
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
            location.href = getBaseUrl() + ELIGIBILITY_LIST_PATH;
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

        IMPORT_ELIG_POPUP_REF = popup;
        try {
            localStorage.setItem(STORAGE_IMPORT_ELIG_POPUP, "1");
        } catch (e) {}

        confirmBtn.addEventListener("click", function () {
            log("ImportElig: Confirm clicked");

            try {
                localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ELIG_IMPORT);
                localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
            } catch (e) {
            }

            var planPriText = (planPriInput.value + "").trim();
            var ignoreText = (ignoreInput.value + "").trim();
            var excText = (excInput.value + "").trim();
            var priText = (priInput.value + "").trim();
            var priOnlyChecked = priOnlyCheckbox.checked;

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
                } else {
                    if (runningBox) {
                        runningBox.textContent = "No more I/E items to add";
                    }
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

        currentKeywords.push(codeStr);

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

        if (formPriorityOnly && formPriority.length > 0) {
            if (planPriority.length > 0) {
                var formPriOnlyPlan = await scanActivitiesForMatch(true, true);
                if (formPriOnlyPlan) {
                    log("ImportElig: TrySelect found match during Form Priority Only scan (priority plans)");
                    return true;
                }
            }
            var formPriOnly = await scanActivitiesForMatch(true, false);
            if (formPriOnly) {
                log("ImportElig: TrySelect found match during Form Priority Only scan");
                return true;
            }
            log("ImportElig: TrySelect no match in Form Priority Only mode");
            return false;
        }

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
                    var verifySAEl = document.querySelector("select#scheduledActivity");
                    if (!verifySAEl || String(verifySAEl.value) !== String(sVal)) {
                        log("ImportElig: SA changed after itemRef reload, expected '" + String(sVal) + "' but got '" + String(verifySAEl ? verifySAEl.value : "null") + "'; skipping");
                        s = s + 1;
                        continue;
                    }
                    await sleep(300);
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



    async function collectMappingsFromModal(existingCodeSet) {
        log("ImportIE: collectMappingsFromModal start");
        var mappings = [];
        var planSel = document.querySelector("select#activityPlan");
        if (!planSel) {
            planSel = await waitForElement("select#activityPlan", 10000);
        }
        if (!planSel) {
            log("ImportIE: collectMappingsFromModal planSel not found");
            return mappings;
        }
        var planOpts = planSel.querySelectorAll("option");
        var pi = 0;
        while (pi < planOpts.length) {
            var pVal = (planOpts[pi].value + "").trim();
            var pTxt = (planOpts[pi].textContent + "").trim().replace(/\s+/g, " ");
            if (pVal.length === 0) {
                pi = pi + 1;
                continue;
            }
            log("ImportIE: collectMappingsFromModal selecting plan='" + String(pTxt) + "' value='" + String(pVal) + "'");
            planSel.value = pVal;
            select2TriggerChange(planSel);
            await sleep(importIERandomDelay() + 300);
            var schedSel = document.querySelector("select#scheduledActivity");
            if (!schedSel) {
                schedSel = await waitForElement("select#scheduledActivity", 8000);
            }
            if (!schedSel) {
                log("ImportIE: collectMappingsFromModal schedSel not found for plan='" + String(pTxt) + "'");
                pi = pi + 1;
                continue;
            }
            var hasSchedOpts = await waitForSelectOptions(schedSel, 1, 5000);
            if (!hasSchedOpts) {
                log("ImportIE: collectMappingsFromModal no SA options for plan='" + String(pTxt) + "'");
                pi = pi + 1;
                continue;
            }
            var schedOpts = schedSel.querySelectorAll("option");
            var si = 0;
            while (si < schedOpts.length) {
                var sVal = (schedOpts[si].value + "").trim();
                var sTxt = (schedOpts[si].textContent + "").trim().replace(/\s+/g, " ");
                if (sVal.length === 0) {
                    si = si + 1;
                    continue;
                }
                log("ImportIE: collectMappingsFromModal selecting SA='" + String(sTxt) + "' value='" + String(sVal) + "'");
                var prevItemSig = getItemRefOptionsSignature();
                schedSel.value = sVal;
                select2TriggerChange(schedSel);
                await sleep(importIERandomDelay() + 200);
                var itemRefSel = document.querySelector("select#itemRef");
                if (!itemRefSel) {
                    itemRefSel = await waitForElement("select#itemRef", 8000);
                }
                if (!itemRefSel) {
                    log("ImportIE: collectMappingsFromModal itemRefSel not found for SA='" + String(sTxt) + "'");
                    si = si + 1;
                    continue;
                }
                var reloaded = await waitForItemRefReload(prevItemSig, 5000);
                if (!reloaded) {
                    log("ImportIE: collectMappingsFromModal itemRef did not reload for SA='" + String(sTxt) + "', checking current options");
                }
                await sleep(300);
                itemRefSel = document.querySelector("select#itemRef");
                if (!itemRefSel) {
                    log("ImportIE: collectMappingsFromModal itemRefSel lost after wait");
                    si = si + 1;
                    continue;
                }
                var itemOpts = itemRefSel.querySelectorAll("option");
                var ii = 0;
                while (ii < itemOpts.length) {
                    var iVal = (itemOpts[ii].value + "").trim();
                    var iTxt = (itemOpts[ii].textContent + "").trim().replace(/\s+/g, " ");
                    if (iVal.length === 0) {
                        ii = ii + 1;
                        continue;
                    }
                    var code = extractIECodeStrict(iTxt);
                    if (code.length === 0) {
                        code = extractIECode(iTxt);
                    }
                    if (code.length > 0) {
                        var record = {
                            activityPlanText: pTxt,
                            scheduledActivityText: sTxt,
                            checkItemText: iTxt,
                            code: code.toUpperCase(),
                            ids: {
                                activityPlanValue: pVal,
                                scheduledActivityValue: sVal,
                                checkItemValue: iVal
                            }
                        };
                        log("ImportIE: collectMappingsFromModal found mapping code=" + String(record.code) + " plan='" + String(pTxt) + "' sa='" + String(sTxt) + "' item='" + String(iTxt) + "'");
                        mappings.push(record);
                    }
                    ii = ii + 1;
                }
                si = si + 1;
            }
            pi = pi + 1;
        }
        log("ImportIE: collectMappingsFromModal done total=" + String(mappings.length));
        return mappings;
    }

    function deduplicateMappings(mappings) {
        log("ImportIE: deduplicateMappings start count=" + String(mappings.length));
        var seen = {};
        var result = [];
        var mi = 0;
        while (mi < mappings.length) {
            var m = mappings[mi];
            var key = m.code + "|" + m.ids.activityPlanValue + "|" + m.ids.scheduledActivityValue + "|" + m.ids.checkItemValue;
            if (!seen.hasOwnProperty(key)) {
                seen[key] = true;
                result.push(m);
            }
            mi = mi + 1;
        }
        log("ImportIE: deduplicateMappings done count=" + String(result.length));
        return result;
    }

    function buildHierarchy(mappings) {
        log("ImportIE: buildHierarchy start");
        var plans = [];
        var planMap = {};
        var mi = 0;
        while (mi < mappings.length) {
            var m = mappings[mi];
            var pKey = m.ids.activityPlanValue;
            if (!planMap.hasOwnProperty(pKey)) {
                planMap[pKey] = {
                    text: m.activityPlanText,
                    value: pKey,
                    activities: [],
                    activityMap: {}
                };
                plans.push(planMap[pKey]);
            }
            var plan = planMap[pKey];
            var sKey = m.ids.scheduledActivityValue;
            if (!plan.activityMap.hasOwnProperty(sKey)) {
                plan.activityMap[sKey] = {
                    text: m.scheduledActivityText,
                    value: sKey,
                    items: []
                };
                plan.activities.push(plan.activityMap[sKey]);
            }
            plan.activityMap[sKey].items.push(m);
            mi = mi + 1;
        }
        log("ImportIE: buildHierarchy done plans=" + String(plans.length));
        return plans;
    }

    function buildImportIEReviewPanel(existingCodeSet, mappings, onConfirm) {
        log("ImportIE: buildImportIEReviewPanel start existingCodes=" + String(existingCodeSet.size) + " mappings=" + String(mappings.length));
        var hierarchy = buildHierarchy(mappings);
        var existingArr = Array.from(existingCodeSet);
        existingArr.sort();

        var overlay = document.createElement("div");
        overlay.id = "importIEReviewOverlay";
        overlay.style.position = "fixed";
        overlay.style.top = "0";
        overlay.style.left = "0";
        overlay.style.width = "100%";
        overlay.style.height = "100%";
        overlay.style.zIndex = "999997";
        overlay.style.background = "rgba(0,0,0,0.6)";
        overlay.style.display = "flex";
        overlay.style.alignItems = "center";
        overlay.style.justifyContent = "center";
        overlay.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        overlay.style.fontSize = "13px";
        overlay.style.color = "#fff";

        var container = document.createElement("div");
        container.style.background = "#111";
        container.style.border = "1px solid #444";
        container.style.borderRadius = "8px";
        container.style.width = "90%";
        container.style.maxWidth = "1100px";
        container.style.height = "85vh";
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.boxShadow = "0 4px 20px rgba(0,0,0,0.5)";

        var headerBar = document.createElement("div");
        headerBar.style.display = "flex";
        headerBar.style.alignItems = "center";
        headerBar.style.justifyContent = "space-between";
        headerBar.style.padding = "12px 16px";
        headerBar.style.borderBottom = "1px solid #444";
        headerBar.style.flexShrink = "0";

        var headerTitle = document.createElement("div");
        headerTitle.textContent = "Import I/E - Review Mappings";
        headerTitle.style.fontWeight = "600";
        headerTitle.style.fontSize = "16px";

        var headerCloseBtn = document.createElement("button");
        headerCloseBtn.textContent = "\u2715";
        headerCloseBtn.style.background = "transparent";
        headerCloseBtn.style.color = "#fff";
        headerCloseBtn.style.border = "none";
        headerCloseBtn.style.cursor = "pointer";
        headerCloseBtn.style.fontSize = "18px";
        headerCloseBtn.style.padding = "4px 8px";
        headerCloseBtn.addEventListener("click", function () {
            log("ImportIE: review panel closed by user");
            overlay.remove();
        });

        headerBar.appendChild(headerTitle);
        headerBar.appendChild(headerCloseBtn);
        container.appendChild(headerBar);

        var panelBody = document.createElement("div");
        panelBody.style.display = "flex";
        panelBody.style.flex = "1";
        panelBody.style.overflow = "hidden";

        var leftPanel = document.createElement("div");
        leftPanel.style.width = "280px";
        leftPanel.style.flexShrink = "0";
        leftPanel.style.borderRight = "1px solid #444";
        leftPanel.style.display = "flex";
        leftPanel.style.flexDirection = "column";
        leftPanel.style.padding = "12px";

        var leftTitle = document.createElement("div");
        leftTitle.textContent = "Existing Table Codes (" + String(existingArr.length) + ")";
        leftTitle.style.fontWeight = "600";
        leftTitle.style.marginBottom = "8px";
        leftTitle.style.fontSize = "13px";
        leftPanel.appendChild(leftTitle);

        var leftSearch = document.createElement("input");
        leftSearch.type = "text";
        leftSearch.placeholder = "Search codes...";
        leftSearch.style.width = "100%";
        leftSearch.style.padding = "6px 8px";
        leftSearch.style.marginBottom = "8px";
        leftSearch.style.background = "#222";
        leftSearch.style.color = "#fff";
        leftSearch.style.border = "1px solid #555";
        leftSearch.style.borderRadius = "4px";
        leftSearch.style.boxSizing = "border-box";
        leftSearch.style.fontSize = "12px";
        leftPanel.appendChild(leftSearch);

        var leftList = document.createElement("div");
        leftList.style.flex = "1";
        leftList.style.overflowY = "auto";

        function renderLeftList(filter) {
            leftList.innerHTML = "";
            var fi = 0;
            while (fi < existingArr.length) {
                var c = existingArr[fi];
                if (filter && filter.length > 0) {
                    if (c.toLowerCase().indexOf(filter.toLowerCase()) < 0) {
                        fi = fi + 1;
                        continue;
                    }
                }
                var row = document.createElement("div");
                row.style.padding = "3px 6px";
                row.style.fontSize = "12px";
                row.style.color = "#aaa";
                row.textContent = c;
                leftList.appendChild(row);
                fi = fi + 1;
            }
        }
        renderLeftList("");
        leftSearch.addEventListener("input", function () {
            renderLeftList(leftSearch.value);
        });
        leftPanel.appendChild(leftList);
        panelBody.appendChild(leftPanel);

        var rightPanel = document.createElement("div");
        rightPanel.style.flex = "1";
        rightPanel.style.display = "flex";
        rightPanel.style.flexDirection = "column";
        rightPanel.style.padding = "12px";
        rightPanel.style.overflow = "hidden";

        var rightTitle = document.createElement("div");
        rightTitle.textContent = "Mapped Check Items";
        rightTitle.style.fontWeight = "600";
        rightTitle.style.marginBottom = "8px";
        rightTitle.style.fontSize = "13px";
        rightPanel.appendChild(rightTitle);

        var rightSearch = document.createElement("input");
        rightSearch.type = "text";
        rightSearch.placeholder = "Search mappings...";
        rightSearch.style.width = "100%";
        rightSearch.style.padding = "6px 8px";
        rightSearch.style.marginBottom = "8px";
        rightSearch.style.background = "#222";
        rightSearch.style.color = "#fff";
        rightSearch.style.border = "1px solid #555";
        rightSearch.style.borderRadius = "4px";
        rightSearch.style.boxSizing = "border-box";
        rightSearch.style.fontSize = "12px";
        rightPanel.appendChild(rightSearch);

        var rightList = document.createElement("div");
        rightList.style.flex = "1";
        rightList.style.overflowY = "auto";
        rightList.style.paddingRight = "4px";

        var allCheckboxes = [];
        var planCheckboxes = [];
        var saCheckboxes = [];
        var itemCheckboxes = [];

        function updateCounter() {
            var count = 0;
            var ci = 0;
            while (ci < itemCheckboxes.length) {
                if (itemCheckboxes[ci].cb.checked && !itemCheckboxes[ci].cb.disabled) {
                    count = count + 1;
                }
                ci = ci + 1;
            }
            counterEl.textContent = "Selected: " + String(count);
            if (count > 0) {
                confirmBtn.disabled = false;
                confirmBtn.style.opacity = "1";
            } else {
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = "0.5";
            }
        }

        function updateParentState(planIdx) {
            var pEntry = planCheckboxes[planIdx];
            if (!pEntry) {
                return;
            }
            var allChecked = true;
            var anyChecked = false;
            var sai = 0;
            while (sai < pEntry.saIndices.length) {
                var saIdx = pEntry.saIndices[sai];
                var saEntry = saCheckboxes[saIdx];
                var saAllChecked = true;
                var saAnyChecked = false;
                var iti = 0;
                while (iti < saEntry.itemIndices.length) {
                    var itIdx = saEntry.itemIndices[iti];
                    var itEntry = itemCheckboxes[itIdx];
                    if (!itEntry.cb.disabled) {
                        if (itEntry.cb.checked) {
                            saAnyChecked = true;
                        } else {
                            saAllChecked = false;
                        }
                    }
                    iti = iti + 1;
                }
                saEntry.cb.checked = saAnyChecked && saAllChecked;
                saEntry.cb.indeterminate = saAnyChecked && !saAllChecked;
                if (saAnyChecked) {
                    anyChecked = true;
                }
                if (!saAllChecked) {
                    allChecked = false;
                }
                sai = sai + 1;
            }
            pEntry.cb.checked = anyChecked && allChecked;
            pEntry.cb.indeterminate = anyChecked && !allChecked;
            updateCounter();
        }

        function setDescendants(planIdx, checked) {
            var pEntry = planCheckboxes[planIdx];
            var sai = 0;
            while (sai < pEntry.saIndices.length) {
                var saIdx = pEntry.saIndices[sai];
                saCheckboxes[saIdx].cb.checked = checked;
                saCheckboxes[saIdx].cb.indeterminate = false;
                var iti = 0;
                while (iti < saCheckboxes[saIdx].itemIndices.length) {
                    var itIdx = saCheckboxes[saIdx].itemIndices[iti];
                    if (!itemCheckboxes[itIdx].cb.disabled) {
                        itemCheckboxes[itIdx].cb.checked = checked;
                    }
                    iti = iti + 1;
                }
                sai = sai + 1;
            }
            updateParentState(planIdx);
        }

        function setSADescendants(saIdx, checked) {
            var saEntry = saCheckboxes[saIdx];
            var iti = 0;
            while (iti < saEntry.itemIndices.length) {
                var itIdx = saEntry.itemIndices[iti];
                if (!itemCheckboxes[itIdx].cb.disabled) {
                    itemCheckboxes[itIdx].cb.checked = checked;
                }
                iti = iti + 1;
            }
            updateParentState(saEntry.planIdx);
        }

        function renderRightList(filter) {
            rightList.innerHTML = "";
            allCheckboxes = [];
            planCheckboxes = [];
            saCheckboxes = [];
            itemCheckboxes = [];
            var filterLower = (filter + "").toLowerCase();

            var hi = 0;
            while (hi < hierarchy.length) {
                var plan = hierarchy[hi];
                var planIdx = planCheckboxes.length;

                var planRow = document.createElement("div");
                planRow.style.display = "flex";
                planRow.style.alignItems = "center";
                planRow.style.gap = "6px";
                planRow.style.padding = "6px 4px";
                planRow.style.fontWeight = "600";
                planRow.style.fontSize = "13px";
                planRow.style.borderBottom = "1px solid #333";
                planRow.style.marginTop = hi > 0 ? "10px" : "0";

                var planCb = document.createElement("input");
                planCb.type = "checkbox";
                planCb.checked = true;
                planCb.style.cursor = "pointer";
                planCb.style.flexShrink = "0";

                var planLabel = document.createElement("span");
                planLabel.textContent = plan.text;
                planLabel.style.cursor = "pointer";

                planRow.appendChild(planCb);
                planRow.appendChild(planLabel);

                var planEntry = { cb: planCb, saIndices: [], row: planRow };
                planCheckboxes.push(planEntry);

                var planVisible = false;

                var sai = 0;
                while (sai < plan.activities.length) {
                    var sa = plan.activities[sai];
                    var saIdx = saCheckboxes.length;
                    planEntry.saIndices.push(saIdx);

                    var saRow = document.createElement("div");
                    saRow.style.display = "flex";
                    saRow.style.alignItems = "center";
                    saRow.style.gap = "6px";
                    saRow.style.padding = "4px 4px 4px 24px";
                    saRow.style.fontSize = "12px";
                    saRow.style.color = "#ccc";

                    var saCb = document.createElement("input");
                    saCb.type = "checkbox";
                    saCb.checked = true;
                    saCb.style.cursor = "pointer";
                    saCb.style.flexShrink = "0";

                    var saLabel = document.createElement("span");
                    saLabel.textContent = "- " + sa.text;
                    saLabel.style.cursor = "pointer";

                    saRow.appendChild(saCb);
                    saRow.appendChild(saLabel);

                    var saEntry = { cb: saCb, itemIndices: [], planIdx: planIdx, row: saRow };
                    saCheckboxes.push(saEntry);

                    var saVisible = false;

                    var iti = 0;
                    while (iti < sa.items.length) {
                        var item = sa.items[iti];
                        var itemIdx = itemCheckboxes.length;
                        saEntry.itemIndices.push(itemIdx);

                        var alreadyExists = existingCodeSet.has(item.code.toUpperCase());
                        var statusText = alreadyExists ? "Already Exist" : "Not Added";

                        var itemRow = document.createElement("div");
                        itemRow.style.display = "flex";
                        itemRow.style.alignItems = "center";
                        itemRow.style.gap = "6px";
                        itemRow.style.padding = "3px 4px 3px 48px";
                        itemRow.style.fontSize = "12px";

                        var itemCb = document.createElement("input");
                        itemCb.type = "checkbox";
                        itemCb.style.cursor = "pointer";
                        itemCb.style.flexShrink = "0";
                        if (alreadyExists) {
                            itemCb.checked = false;
                            itemCb.disabled = true;
                            itemCb.style.cursor = "not-allowed";
                        } else {
                            itemCb.checked = true;
                        }

                        var itemLabel = document.createElement("span");
                        itemLabel.textContent = "-- " + item.checkItemText;
                        itemLabel.style.flex = "1";
                        itemLabel.style.wordBreak = "break-word";

                        var statusBadge = document.createElement("span");
                        statusBadge.style.fontSize = "10px";
                        statusBadge.style.padding = "2px 6px";
                        statusBadge.style.borderRadius = "3px";
                        statusBadge.style.flexShrink = "0";
                        if (alreadyExists) {
                            statusBadge.textContent = "Already Exist";
                            statusBadge.style.background = "#555";
                            statusBadge.style.color = "#aaa";
                        } else {
                            statusBadge.textContent = "Not Added";
                            statusBadge.style.background = "#1a4a1a";
                            statusBadge.style.color = "#6f6";
                        }

                        itemRow.appendChild(itemCb);
                        itemRow.appendChild(itemLabel);
                        itemRow.appendChild(statusBadge);

                        var itemEntry = { cb: itemCb, mapping: item, row: itemRow, planIdx: planIdx, saIdx: saIdx };
                        itemCheckboxes.push(itemEntry);

                        var matchesFilter = true;
                        if (filterLower.length > 0) {
                            var combined = (item.code + " " + item.checkItemText + " " + item.activityPlanText + " " + item.scheduledActivityText).toLowerCase();
                            if (combined.indexOf(filterLower) < 0) {
                                matchesFilter = false;
                            }
                        }
                        if (matchesFilter) {
                            saVisible = true;
                        }
                        itemRow.style.display = matchesFilter ? "flex" : "none";

                        (function (capturedItemIdx, capturedPlanIdx) {
                            itemCb.addEventListener("change", function () {
                                updateParentState(capturedPlanIdx);
                            });
                        })(itemIdx, planIdx);

                        allCheckboxes.push({ type: "item", idx: itemIdx, row: itemRow });
                        iti = iti + 1;
                    }

                    if (saVisible) {
                        planVisible = true;
                    }
                    saRow.style.display = saVisible ? "flex" : "none";

                    (function (capturedSaIdx) {
                        saCb.addEventListener("change", function () {
                            setSADescendants(capturedSaIdx, saCb.checked);
                        });
                    })(saIdx);

                    allCheckboxes.push({ type: "sa", idx: saIdx, row: saRow });
                    sai = sai + 1;
                }

                planRow.style.display = planVisible ? "flex" : "none";

                (function (capturedPlanIdx) {
                    planCb.addEventListener("change", function () {
                        setDescendants(capturedPlanIdx, planCb.checked);
                    });
                })(planIdx);

                allCheckboxes.push({ type: "plan", idx: planIdx, row: planRow });

                rightList.appendChild(planRow);

                var sa2i = 0;
                while (sa2i < plan.activities.length) {
                    var saE = saCheckboxes[planEntry.saIndices[sa2i]];
                    rightList.appendChild(saE.row);
                    var it2i = 0;
                    while (it2i < saE.itemIndices.length) {
                        var itE = itemCheckboxes[saE.itemIndices[it2i]];
                        rightList.appendChild(itE.row);
                        it2i = it2i + 1;
                    }
                    sa2i = sa2i + 1;
                }

                hi = hi + 1;
            }

            var pci = 0;
            while (pci < planCheckboxes.length) {
                updateParentState(pci);
                pci = pci + 1;
            }
        }

        renderRightList("");
        rightSearch.addEventListener("input", function () {
            renderRightList(rightSearch.value);
        });
        rightPanel.appendChild(rightList);

        var counterEl = document.createElement("div");
        counterEl.style.padding = "8px 0";
        counterEl.style.fontWeight = "500";
        counterEl.style.fontSize = "13px";
        counterEl.textContent = "Selected: 0";
        rightPanel.appendChild(counterEl);

        panelBody.appendChild(rightPanel);
        container.appendChild(panelBody);

        var footerBar = document.createElement("div");
        footerBar.style.display = "flex";
        footerBar.style.alignItems = "center";
        footerBar.style.justifyContent = "flex-end";
        footerBar.style.gap = "10px";
        footerBar.style.padding = "12px 16px";
        footerBar.style.borderTop = "1px solid #444";
        footerBar.style.flexShrink = "0";

        var selectAllBtn = document.createElement("button");
        selectAllBtn.textContent = "Select All";
        selectAllBtn.style.background = "#333";
        selectAllBtn.style.color = "#fff";
        selectAllBtn.style.border = "none";
        selectAllBtn.style.padding = "8px 16px";
        selectAllBtn.style.borderRadius = "4px";
        selectAllBtn.style.cursor = "pointer";
        selectAllBtn.addEventListener("click", function () {
            log("ImportIE: Select All clicked");
            var xi = 0;
            while (xi < itemCheckboxes.length) {
                if (!itemCheckboxes[xi].cb.disabled) {
                    itemCheckboxes[xi].cb.checked = true;
                }
                xi = xi + 1;
            }
            var pci2 = 0;
            while (pci2 < planCheckboxes.length) {
                updateParentState(pci2);
                pci2 = pci2 + 1;
            }
        });

        var clearAllBtn = document.createElement("button");
        clearAllBtn.textContent = "Clear All";
        clearAllBtn.style.background = "#333";
        clearAllBtn.style.color = "#fff";
        clearAllBtn.style.border = "none";
        clearAllBtn.style.padding = "8px 16px";
        clearAllBtn.style.borderRadius = "4px";
        clearAllBtn.style.cursor = "pointer";
        clearAllBtn.addEventListener("click", function () {
            log("ImportIE: Clear All clicked");
            var xi2 = 0;
            while (xi2 < itemCheckboxes.length) {
                if (!itemCheckboxes[xi2].cb.disabled) {
                    itemCheckboxes[xi2].cb.checked = false;
                }
                xi2 = xi2 + 1;
            }
            var pci3 = 0;
            while (pci3 < planCheckboxes.length) {
                updateParentState(pci3);
                pci3 = pci3 + 1;
            }
        });

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.style.background = "#0a0";
        confirmBtn.style.color = "#fff";
        confirmBtn.style.border = "none";
        confirmBtn.style.padding = "8px 24px";
        confirmBtn.style.borderRadius = "4px";
        confirmBtn.style.cursor = "pointer";
        confirmBtn.style.fontWeight = "600";

        confirmBtn.addEventListener("click", function () {
            log("ImportIE: Confirm clicked in review panel");
            var selected = [];
            var gi = 0;
            while (gi < itemCheckboxes.length) {
                if (itemCheckboxes[gi].cb.checked && !itemCheckboxes[gi].cb.disabled) {
                    selected.push(itemCheckboxes[gi].mapping);
                }
                gi = gi + 1;
            }
            log("ImportIE: confirmed " + String(selected.length) + " items");
            overlay.remove();
            onConfirm(selected);
        });

        footerBar.appendChild(selectAllBtn);
        footerBar.appendChild(clearAllBtn);
        footerBar.appendChild(confirmBtn);
        container.appendChild(footerBar);

        overlay.appendChild(container);
        document.body.appendChild(overlay);

        updateCounter();
        log("ImportIE: buildImportIEReviewPanel rendered");
    }

    function buildProgressPopup(selectedMappings) {
        log("ImportIE: buildProgressPopup items=" + String(selectedMappings.length));
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "10px";

        var summaryRow = document.createElement("div");
        summaryRow.style.display = "flex";
        summaryRow.style.justifyContent = "space-between";
        summaryRow.style.fontSize = "13px";
        summaryRow.style.padding = "4px 0";
        summaryRow.style.borderBottom = "1px solid #333";

        var totalEl = document.createElement("span");
        totalEl.textContent = "Total: " + String(selectedMappings.length);
        var successEl = document.createElement("span");
        successEl.textContent = "Success: 0";
        successEl.style.color = "#5cb85c";
        var failEl = document.createElement("span");
        failEl.textContent = "Failed: 0";
        failEl.style.color = "#d9534f";
        var statusEl = document.createElement("span");
        statusEl.textContent = "In Progress";
        statusEl.style.color = "#5bc0de";
        statusEl.style.fontWeight = "600";

        summaryRow.appendChild(totalEl);
        summaryRow.appendChild(successEl);
        summaryRow.appendChild(failEl);
        summaryRow.appendChild(statusEl);
        container.appendChild(summaryRow);

        var listEl = document.createElement("div");
        listEl.style.maxHeight = "400px";
        listEl.style.overflowY = "auto";
        listEl.style.fontSize = "12px";

        var rows = [];
        var ri = 0;
        while (ri < selectedMappings.length) {
            var m = selectedMappings[ri];
            var row = document.createElement("div");
            row.style.display = "flex";
            row.style.justifyContent = "space-between";
            row.style.alignItems = "center";
            row.style.padding = "4px 6px";
            row.style.borderBottom = "1px solid #222";

            var labelSpan = document.createElement("span");
            labelSpan.textContent = m.code + " - " + m.checkItemText;
            labelSpan.style.flex = "1";
            labelSpan.style.overflow = "hidden";
            labelSpan.style.textOverflow = "ellipsis";
            labelSpan.style.whiteSpace = "nowrap";
            labelSpan.style.marginRight = "8px";

            var statusSpan = document.createElement("span");
            statusSpan.textContent = "Pending";
            statusSpan.style.color = "#888";
            statusSpan.style.flexShrink = "0";
            statusSpan.style.fontSize = "11px";

            row.appendChild(labelSpan);
            row.appendChild(statusSpan);
            listEl.appendChild(row);
            rows.push({ row: row, statusSpan: statusSpan });
            ri = ri + 1;
        }
        container.appendChild(listEl);

        var popup = createPopup({
            title: "Import I/E - Progress",
            content: container,
            width: "600px",
            height: "auto",
            maxHeight: "80vh"
        });

        return {
            popup: popup,
            rows: rows,
            successEl: successEl,
            failEl: failEl,
            statusEl: statusEl,
            updateItem: function (index, status, msg) {
                if (index >= 0 && index < rows.length) {
                    rows[index].statusSpan.textContent = status;
                    if (status === "Success") {
                        rows[index].statusSpan.style.color = "#5cb85c";
                    } else if (status === "Failed") {
                        rows[index].statusSpan.style.color = "#d9534f";
                        if (msg) {
                            rows[index].statusSpan.title = msg;
                        }
                    }
                    listEl.scrollTop = listEl.scrollHeight;
                }
            },
            updateSummary: function (successes, failures) {
                successEl.textContent = "Success: " + String(successes);
                failEl.textContent = "Failed: " + String(failures);
            },
            setCompleted: function () {
                statusEl.textContent = "Completed";
                statusEl.style.color = "#5cb85c";
            }
        };
    }

    async function executeSelectedMappings(selectedMappings, existingCodeSet) {
        log("ImportIE: executeSelectedMappings start count=" + String(selectedMappings.length));
        var progress = buildProgressPopup(selectedMappings);
        var successes = 0;
        var failures = 0;

        var mi = 0;
        while (mi < selectedMappings.length) {
            var mapping = selectedMappings[mi];
            log("ImportIE: processing item " + String(mi + 1) + "/" + String(selectedMappings.length) + " code=" + String(mapping.code));

            var modalOpen = document.querySelector("#ajaxModal .modal-content");
            if (!modalOpen) {
                log("ImportIE: modal not open, clicking #addEligButton");
                var addBtn = document.querySelector("a#addEligButton");
                if (!addBtn) {
                    addBtn = document.querySelector("#addEligButton");
                }
                if (!addBtn) {
                    log("ImportIE: #addEligButton not found");
                    progress.updateItem(mi, "Failed", "Add button not found");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
                if (addBtn.hasAttribute("disabled")) {
                    log("ImportIE: #addEligButton is disabled");
                    progress.updateItem(mi, "Failed", "Add button disabled");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
                addBtn.click();
                var opened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                if (!opened) {
                    log("ImportIE: modal did not open on first try, retrying");
                    addBtn.click();
                    opened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                }
                if (!opened) {
                    log("ImportIE: modal failed to open after retry");
                    progress.updateItem(mi, "Failed", "Modal did not open");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    mi = mi + 1;
                    continue;
                }
            }
            await sleep(500);

            log("ImportIE: step b - selecting eligibility item for code=" + String(mapping.code));
            var eligSel = document.querySelector("select#eligibilityItemRef");
            if (!eligSel) {
                eligSel = await waitForElement("select#eligibilityItemRef", 10000);
            }
            if (!eligSel) {
                log("ImportIE: eligibilityItemRef not found");
                progress.updateItem(mi, "Failed", "Eligibility select not found");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb1 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb1) {
                    cb1.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            var eligMatched = false;
            var eligOpts = eligSel.querySelectorAll("option");
            var bestMatch = null;
            var bestMatchVal = "";
            var eoi = 0;
            while (eoi < eligOpts.length) {
                var eoVal = (eligOpts[eoi].value + "").trim();
                var eoTxt = (eligOpts[eoi].textContent + "").trim();
                if (eoVal.length > 0) {
                    var eoCode = extractIECodeStrict(eoTxt);
                    if (eoCode.length === 0) {
                        eoCode = parseItemCodeFromEligibilityOptionText(eoTxt);
                    }
                    if (eoCode.toUpperCase() === mapping.code.toUpperCase()) {
                        bestMatch = eligOpts[eoi];
                        bestMatchVal = eoVal;
                        break;
                    }
                    if (!bestMatch) {
                        var eoTxtUpper = eoTxt.toUpperCase();
                        var codeUpper = mapping.code.toUpperCase();
                        var re = new RegExp("\\b" + codeUpper.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i");
                        if (re.test(eoTxtUpper)) {
                            bestMatch = eligOpts[eoi];
                            bestMatchVal = eoVal;
                        }
                    }
                }
                eoi = eoi + 1;
            }

            if (!bestMatch) {
                log("ImportIE: no eligibility item match for code=" + String(mapping.code));
                progress.updateItem(mi, "Failed", "No eligibility item match");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb2 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb2) {
                    cb2.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            eligSel.value = bestMatchVal;
            select2TriggerChange(eligSel);
            log("ImportIE: eligibility item selected value='" + String(bestMatchVal) + "'");
            await sleep(importIERandomDelay() + 200);
            eligMatched = true;

            log("ImportIE: step c - setting sex option");
            var eligItemText = (bestMatch.textContent + "").trim();
            var sex = parseSexFromEligibilityText(eligItemText);
            var sexSel = document.querySelector("select#sexOption");
            if (sexSel) {
                sexSel.value = sex;
                select2TriggerChange(sexSel);
                log("ImportIE: sex set to '" + String(sex) + "'");
                await sleep(importIERandomDelay());
            } else {
                log("ImportIE: sexOption select not found, skipping sex step");
            }

            log("ImportIE: step d - selecting activity plan value='" + String(mapping.ids.activityPlanValue) + "'");
            var planOk = await select2SelectByValue("select#activityPlan", mapping.ids.activityPlanValue);
            if (!planOk) {
                planOk = await select2SelectByText("select#activityPlan", mapping.activityPlanText);
            }
            if (!planOk) {
                log("ImportIE: activity plan selection failed");
                progress.updateItem(mi, "Failed", "Activity plan selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb3 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb3) {
                    cb3.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(importIERandomDelay() + 300);

            log("ImportIE: step e - selecting scheduled activity value='" + String(mapping.ids.scheduledActivityValue) + "'");
            var schedSelE = document.querySelector("select#scheduledActivity");
            if (schedSelE) {
                await waitForSelectOptions(schedSelE, 1, 8000);
            }
            var saOk = await select2SelectByValue("select#scheduledActivity", mapping.ids.scheduledActivityValue);
            if (!saOk) {
                saOk = await select2SelectByText("select#scheduledActivity", mapping.scheduledActivityText);
            }
            if (!saOk) {
                log("ImportIE: scheduled activity selection failed");
                progress.updateItem(mi, "Failed", "Scheduled activity selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb4 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb4) {
                    cb4.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(importIERandomDelay() + 300);

            log("ImportIE: step f - selecting check item value='" + String(mapping.ids.checkItemValue) + "'");
            var itemRefSelF = document.querySelector("select#itemRef");
            if (itemRefSelF) {
                await waitForSelectOptions(itemRefSelF, 1, 8000);
            }
            var prevSigF = getItemRefOptionsSignature();
            await waitForItemRefReload(prevSigF, 5000);
            var itemOk = await select2SelectByValue("select#itemRef", mapping.ids.checkItemValue);
            if (!itemOk) {
                itemOk = await select2SelectByText("select#itemRef", mapping.checkItemText);
            }
            if (!itemOk) {
                log("ImportIE: check item selection failed");
                progress.updateItem(mi, "Failed", "Check item selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb5 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb5) {
                    cb5.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }
            await sleep(importIERandomDelay() + 200);

            log("ImportIE: step g - setting comparator to EQ");
            var compOk = await setComparatorEQ();
            if (!compOk) {
                await stabilizeSelectionBeforeComparator(4000);
                compOk = await setComparatorEQ();
            }
            if (!compOk) {
                log("ImportIE: comparator EQ failed");
                progress.updateItem(mi, "Failed", "Comparator EQ failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb6 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb6) {
                    cb6.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            log("ImportIE: step h - selecting code list item containing SF");
            var sfOk = await selectCodeListValueContainingSF();
            if (!sfOk) {
                log("ImportIE: SF selection failed");
                progress.updateItem(mi, "Failed", "Code list SF selection failed");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                var cb7 = document.querySelector("#ajaxModal .modal-content button.close");
                if (cb7) {
                    cb7.click();
                    await waitForModalClose(5000);
                }
                mi = mi + 1;
                continue;
            }

            log("ImportIE: step i - clicking Save");
            var saveBtn = document.querySelector("button#actionButton");
            if (!saveBtn) {
                saveBtn = await waitForElement("button#actionButton", 5000);
            }
            if (!saveBtn) {
                log("ImportIE: Save button not found");
                progress.updateItem(mi, "Failed", "Save button not found");
                failures = failures + 1;
                progress.updateSummary(successes, failures);
                mi = mi + 1;
                continue;
            }
            saveBtn.click();
            log("ImportIE: Save clicked, waiting for modal close");
            var modalClosed = await waitForModalClose(15000);
            if (!modalClosed) {
                log("ImportIE: modal did not close after save, checking for errors");
                var errorEl = document.querySelector("#ajaxModal .modal-content .alert-danger, #ajaxModal .modal-content .error-message, #ajaxModal .modal-content .has-error");
                if (errorEl) {
                    var errorText = (errorEl.textContent + "").trim().replace(/\s+/g, " ");
                    log("ImportIE: validation error detected: " + String(errorText));
                    progress.updateItem(mi, "Failed", errorText);
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    var cb8 = document.querySelector("#ajaxModal .modal-content button.close");
                    if (cb8) {
                        cb8.click();
                        await waitForModalClose(5000);
                    }
                    mi = mi + 1;
                    continue;
                }
                log("ImportIE: retrying save");
                saveBtn = document.querySelector("button#actionButton");
                if (saveBtn) {
                    saveBtn.click();
                    modalClosed = await waitForModalClose(10000);
                }
                if (!modalClosed) {
                    log("ImportIE: save retry failed");
                    progress.updateItem(mi, "Failed", "Modal did not close after save");
                    failures = failures + 1;
                    progress.updateSummary(successes, failures);
                    var cb9 = document.querySelector("#ajaxModal .modal-content button.close");
                    if (cb9) {
                        cb9.click();
                        await waitForModalClose(5000);
                    }
                    mi = mi + 1;
                    continue;
                }
            }

            log("ImportIE: item saved successfully code=" + String(mapping.code));
            progress.updateItem(mi, "Success");
            successes = successes + 1;
            progress.updateSummary(successes, failures);
            existingCodeSet.add(mapping.code.toUpperCase());
            await sleep(importIERandomDelay() + 500);

            mi = mi + 1;
        }

        progress.setCompleted();
        log("ImportIE: executeSelectedMappings done successes=" + String(successes) + " failures=" + String(failures));
    }

    function startImportEligibilityMapping() {
        log("ImportIE: startImportEligibilityMapping invoked");

        if (!isValidEligibilityPage()) {
            showWarningPopup("Import I/E - Wrong Page", "You must navigate to the Eligibility list page before using Import I/E. Please go to: " + getBaseUrl() + ELIGIBILITY_LIST_PATH);
            log("ImportIE: invalid page, stopping");
            return;
        }

        log("ImportIE: page validated, starting collection flow");

        var loadingEl = document.createElement("div");
        loadingEl.style.textAlign = "center";
        loadingEl.style.fontSize = "15px";
        loadingEl.style.color = "#fff";
        loadingEl.style.padding = "20px";
        loadingEl.textContent = "Collecting existing table codes and mappings...";

        var loadingPopup = createPopup({
            title: "Import I/E - Scanning",
            content: loadingEl,
            width: "500px",
            height: "auto"
        });

        var dots = 1;
        var loadingInterval = setInterval(function () {
            dots = dots + 1;
            if (dots > 3) {
                dots = 1;
            }
            var t = "Scanning";
            var di = 0;
            while (di < dots) {
                t = t + ".";
                di = di + 1;
            }
            loadingEl.textContent = t;
        }, 400);

        setTimeout(async function () {
            try {
                log("ImportIE: step 1 - collecting existing codes from table");
                loadingEl.textContent = "Step 1: Collecting existing table codes...";
                var existingCodeSet = await collectAllTableCodes();
                log("ImportIE: existing codes collected count=" + String(existingCodeSet.size));

                log("ImportIE: step 2 - checking Add button");
                loadingEl.textContent = "Step 2: Checking Add button...";
                var addBtn = document.querySelector("a#addEligButton");
                if (!addBtn) {
                    addBtn = document.querySelector("#addEligButton");
                }
                if (!addBtn) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Error", "The Add button (#addEligButton) was not found on this page.");
                    log("ImportIE: add button not found, stopping");
                    return;
                }
                if (addBtn.hasAttribute("disabled")) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Add Button Disabled", "The Add button is currently disabled. Please ensure you have permission to add eligibility items and that the mapping is unlocked.");
                    log("ImportIE: add button disabled, stopping");
                    return;
                }

                log("ImportIE: step 3 - clicking Add to open modal");
                loadingEl.textContent = "Step 3: Opening modal...";
                addBtn.click();
                var modalOpened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                if (!modalOpened) {
                    log("ImportIE: modal did not open on first try, retrying");
                    addBtn.click();
                    modalOpened = await waitForModalOpen(IMPORT_IE_MODAL_TIMEOUT);
                }
                if (!modalOpened) {
                    clearInterval(loadingInterval);
                    loadingPopup.close();
                    showWarningPopup("Import I/E - Error", "The Eligibility Management modal did not open. Please try again.");
                    log("ImportIE: modal failed to open, stopping");
                    return;
                }
                await sleep(800);

                log("ImportIE: step 4 - collecting mappings from modal");
                loadingEl.textContent = "Step 4: Scanning all Activity Plans, Scheduled Activities, and Check Items...";
                var rawMappings = await collectMappingsFromModal(existingCodeSet);
                log("ImportIE: raw mappings collected count=" + String(rawMappings.length));

                var mappings = deduplicateMappings(rawMappings);
                log("ImportIE: deduplicated mappings count=" + String(mappings.length));

                log("ImportIE: closing modal after collection");
                var closeModalBtn = document.querySelector("#ajaxModal .modal-content button.close");
                if (closeModalBtn) {
                    closeModalBtn.click();
                    await waitForModalClose(5000);
                }

                clearInterval(loadingInterval);
                loadingPopup.close();

                if (mappings.length === 0) {
                    showWarningPopup("Import I/E - No Mappings", "No INC/EXC check items were found across any Activity Plans and Scheduled Activities.");
                    log("ImportIE: no mappings found, stopping");
                    return;
                }

                log("ImportIE: step 5 - showing review panel");
                buildImportIEReviewPanel(existingCodeSet, mappings, function (selectedMappings) {
                    log("ImportIE: user confirmed " + String(selectedMappings.length) + " mappings");
                    if (selectedMappings.length === 0) {
                        log("ImportIE: no items selected, stopping");
                        return;
                    }
                    executeSelectedMappings(selectedMappings, existingCodeSet);
                });

            } catch (err) {
                clearInterval(loadingInterval);
                if (loadingPopup && loadingPopup.close) {
                    loadingPopup.close();
                }
                log("ImportIE: error in startImportEligibilityMapping: " + String(err));
                showWarningPopup("Import I/E - Error", "An unexpected error occurred: " + String(err));
            }
        }, 300);
    }

    function updateRunAllPopupStatus(statusText) {
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

    function BarcodeFunctions() {}

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

    function clearBarcodeSubjectId() {
        try {
            localStorage.removeItem(STORAGE_BARCODE_SUBJECT_ID);
        } catch (e) { }
    }

    function setBarcodeResult(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_RESULT, String(t));
        } catch (e) { }
    }

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


    function setBarcodeSubjectText(t) {
        try {
            localStorage.setItem(STORAGE_BARCODE_SUBJECT_TEXT, String(t));
        } catch (e) { }
    }

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
    function normalizeSubjectString(t) {
        if (typeof t !== "string") {
            return "";
        }
        var s = t.replace(/\u00A0/g, " ");
        s = s.trim();
        s = s.replace(/\s+/g, " ");
        return s;
    }

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

    function isBarcodeSubjectsPage() {
        var path = location.pathname;
        if (path === "/secure/barcodeprinting/subjects") {
            return true;
        }
        return false;
    }



    function SharedUtilityFunctions() {}

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


    function log(msg) {
        try {
            var d = new Date();
            var hours = d.getHours();
            var minutes = d.getMinutes();
            var seconds = d.getSeconds();

            var ampm = "AM";
            if (hours >= 12) {
                ampm = "PM";
            }

            if (hours === 0) {
                hours = 12;
            } else {
                if (hours > 12) {
                    hours = hours - 12;
                }
            }

            if (minutes < 10) {
                minutes = "0" + minutes;
            }

            if (seconds < 10) {
                seconds = "0" + seconds;
            }

            var ts = hours + ":" + minutes + ":" + seconds + " " + ampm;
            var line = "[" + ts + "] " + String(msg);
            console.log("[APS] " + line);
        } catch (e) {
        }

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
        } catch (e3) {
        }
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

    function importIERandomDelay() {
        var range = IMPORT_IE_SHORT_DELAY_MAX - IMPORT_IE_SHORT_DELAY_MIN;
        var val = IMPORT_IE_SHORT_DELAY_MIN + Math.floor(Math.random() * range);
        return val;
    }

    async function delay(ms) {
        log("ImportIE: delay " + String(ms) + "ms");
        await sleep(ms);
    }

    async function waitForElement(selector, timeoutMs) {
        log("ImportIE: waitForElement selector='" + String(selector) + "' timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_HELPER_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var el = document.querySelector(selector);
            if (el) {
                log("ImportIE: waitForElement found selector='" + String(selector) + "' in " + String(Date.now() - start) + "ms");
                return el;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForElement timeout selector='" + String(selector) + "'");
        return null;
    }

    async function waitForSelectOptions(selectEl, minCount, timeoutMs) {
        log("ImportIE: waitForSelectOptions minCount=" + String(minCount) + " timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_HELPER_TIMEOUT;
        var minC = typeof minCount === "number" ? minCount : 1;
        while (Date.now() - start < maxMs) {
            if (selectEl) {
                var opts = selectEl.querySelectorAll("option");
                var nonEmpty = 0;
                var oi = 0;
                while (oi < opts.length) {
                    var v = (opts[oi].value + "").trim();
                    if (v.length > 0) {
                        nonEmpty = nonEmpty + 1;
                    }
                    oi = oi + 1;
                }
                if (nonEmpty >= minC) {
                    log("ImportIE: waitForSelectOptions got " + String(nonEmpty) + " options in " + String(Date.now() - start) + "ms");
                    return true;
                }
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForSelectOptions timeout");
        return false;
    }

    async function waitForModalOpen(timeoutMs) {
        log("ImportIE: waitForModalOpen timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_MODAL_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var modal = document.querySelector("#ajaxModal .modal-content");
            if (modal) {
                var header = modal.querySelector(".modal-header");
                if (header) {
                    var txt = (header.textContent + "").trim();
                    if (txt.indexOf("Eligibility Management") >= 0) {
                        log("ImportIE: waitForModalOpen modal found in " + String(Date.now() - start) + "ms");
                        return true;
                    }
                }
                log("ImportIE: waitForModalOpen modal content found in " + String(Date.now() - start) + "ms");
                return true;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForModalOpen timeout");
        return false;
    }

    async function waitForModalClose(timeoutMs) {
        log("ImportIE: waitForModalClose timeout=" + String(timeoutMs));
        var start = Date.now();
        var maxMs = typeof timeoutMs === "number" ? timeoutMs : IMPORT_IE_MODAL_TIMEOUT;
        while (Date.now() - start < maxMs) {
            var modal = document.querySelector("#ajaxModal .modal-content");
            if (!modal) {
                log("ImportIE: waitForModalClose modal closed in " + String(Date.now() - start) + "ms");
                return true;
            }
            var display = window.getComputedStyle(modal.closest("#ajaxModal") || modal).display;
            if (display === "none") {
                log("ImportIE: waitForModalClose modal hidden in " + String(Date.now() - start) + "ms");
                return true;
            }
            await sleep(IMPORT_IE_POLL_INTERVAL);
        }
        log("ImportIE: waitForModalClose timeout");
        return false;
    }

    async function clickWithRetry(elOrSelector) {
        log("ImportIE: clickWithRetry start");
        var el = null;
        if (typeof elOrSelector === "string") {
            el = document.querySelector(elOrSelector);
        } else {
            el = elOrSelector;
        }
        if (!el) {
            log("ImportIE: clickWithRetry element not found on first attempt");
            await sleep(500);
            if (typeof elOrSelector === "string") {
                el = document.querySelector(elOrSelector);
            }
            if (!el) {
                log("ImportIE: clickWithRetry element not found on retry");
                return false;
            }
        }
        try {
            el.click();
            log("ImportIE: clickWithRetry clicked");
            return true;
        } catch (e) {
            log("ImportIE: clickWithRetry first click error=" + String(e));
            await sleep(300);
            try {
                el.click();
                log("ImportIE: clickWithRetry retry click succeeded");
                return true;
            } catch (e2) {
                log("ImportIE: clickWithRetry retry click failed=" + String(e2));
                return false;
            }
        }
    }

    function select2TriggerChange(selectEl) {
        log("ImportIE: select2TriggerChange");
        var evt = new Event("change", { bubbles: true });
        selectEl.dispatchEvent(evt);
        if (typeof jQuery !== "undefined" && jQuery && jQuery.fn) {
            try {
                jQuery(selectEl).trigger("change");
                log("ImportIE: select2TriggerChange jQuery trigger fired");
            } catch (e) {
                log("ImportIE: select2TriggerChange jQuery trigger error=" + String(e));
            }
        }
    }

    async function select2SelectByValue(containerOrSelect, value) {
        log("ImportIE: select2SelectByValue value='" + String(value) + "'");
        var sel = null;
        if (containerOrSelect && containerOrSelect.tagName && containerOrSelect.tagName.toLowerCase() === "select") {
            sel = containerOrSelect;
        } else if (typeof containerOrSelect === "string") {
            sel = document.querySelector(containerOrSelect);
        } else {
            var under = containerOrSelect.querySelector("select");
            if (under) {
                sel = under;
            }
        }
        if (!sel) {
            log("ImportIE: select2SelectByValue select element not found");
            return false;
        }
        sel.value = String(value);
        select2TriggerChange(sel);
        await sleep(importIERandomDelay());
        var after = (sel.value + "").trim();
        if (after === String(value).trim()) {
            log("ImportIE: select2SelectByValue confirmed value='" + String(after) + "'");
            return true;
        }
        log("ImportIE: select2SelectByValue value mismatch expected='" + String(value) + "' got='" + String(after) + "'");
        sel.value = String(value);
        select2TriggerChange(sel);
        await sleep(300);
        after = (sel.value + "").trim();
        log("ImportIE: select2SelectByValue retry result='" + String(after) + "'");
        return after === String(value).trim();
    }

    async function select2SelectByText(containerOrSelect, text) {
        log("ImportIE: select2SelectByText text='" + String(text) + "'");
        var sel = null;
        if (containerOrSelect && containerOrSelect.tagName && containerOrSelect.tagName.toLowerCase() === "select") {
            sel = containerOrSelect;
        } else if (typeof containerOrSelect === "string") {
            sel = document.querySelector(containerOrSelect);
        } else {
            var under = containerOrSelect.querySelector("select");
            if (under) {
                sel = under;
            }
        }
        if (!sel) {
            log("ImportIE: select2SelectByText select element not found");
            return false;
        }
        var opts = sel.querySelectorAll("option");
        var normalTarget = normalizeTextForCompare(text);
        var oi = 0;
        while (oi < opts.length) {
            var oTxt = normalizeTextForCompare((opts[oi].textContent + ""));
            if (oTxt === normalTarget) {
                sel.value = opts[oi].value;
                select2TriggerChange(sel);
                await sleep(importIERandomDelay());
                log("ImportIE: select2SelectByText matched exact text index=" + String(oi));
                return true;
            }
            oi = oi + 1;
        }
        log("ImportIE: select2SelectByText no exact match for text='" + String(text) + "'");
        return false;
    }

    function normalizeTextForCompare(t) {
        var s = (t + "").trim();
        s = s.replace(/\s+/g, " ");
        s = s.toLowerCase();
        return s;
    }

    function extractIECode(text) {
        var s = (text + "").trim();
        var match = IE_CODE_REGEX.exec(s);
        if (match) {
            var prefix = match[1].toUpperCase();
            var num = match[2];
            var code = prefix + num;
            return code;
        }
        return "";
    }

    function extractIECodeStrict(text) {
        var s = (text + "").trim();
        var re = /\b(INC|EXC)(\d+)\b/i;
        var match = re.exec(s);
        if (match) {
            var prefix = match[1].toUpperCase();
            var num = match[2];
            return prefix + num;
        }
        var re2 = /\b(INC|EXC)\s+(\d+)\b/i;
        var match2 = re2.exec(s);
        if (match2) {
            var prefix2 = match2[1].toUpperCase();
            var num2 = match2[2];
            return prefix2 + num2;
        }
        return "";
    }

    function isValidEligibilityPage() {
        log("ImportIE: isValidEligibilityPage checking");
        var hostname = location.hostname;
        var path = location.pathname;
        var validHost = false;
        var hi = 0;
        while (hi < ELIGIBILITY_VALID_HOSTNAMES.length) {
            if (hostname === ELIGIBILITY_VALID_HOSTNAMES[hi]) {
                validHost = true;
                break;
            }
            hi = hi + 1;
        }
        if (!validHost) {
            log("ImportIE: invalid hostname='" + String(hostname) + "'");
            return false;
        }
        if (path !== ELIGIBILITY_LIST_PATH) {
            log("ImportIE: invalid path='" + String(path) + "'");
            return false;
        }
        log("ImportIE: valid eligibility page");
        return true;
    }

    function getBaseUrl() {
        return location.protocol + "//" + location.hostname;
    }

    function showWarningPopup(title, message) {
        log("ImportIE: showWarningPopup title='" + String(title) + "'");
        var msgEl = document.createElement("div");
        msgEl.style.textAlign = "center";
        msgEl.style.fontSize = "15px";
        msgEl.style.color = "#fff";
        msgEl.style.padding = "20px";
        msgEl.style.lineHeight = "1.6";
        msgEl.textContent = message;
        createPopup({
            title: title,
            content: msgEl,
            width: "460px",
            height: "auto"
        });
    }

    async function collectAllTableCodes() {
        log("ImportIE: collectAllTableCodes start");
        var codeSet = new Set();
        var tbody = document.querySelector("tbody#eligibilityreftablebody");
        if (!tbody) {
            tbody = document.querySelector("tbody#eligibilityRefTableBody");
            if (tbody) {
                log("ImportIE: collectAllTableCodes fallback selector tbody#eligibilityRefTableBody matched");
            }
        }
        if (!tbody) {
            tbody = document.querySelector("#eligibilityRefTable tbody");
            if (tbody) {
                log("ImportIE: collectAllTableCodes fallback selector #eligibilityRefTable tbody matched");
            }
        }
        if (!tbody) {
            log("ImportIE: collectAllTableCodes no table body found");
            return codeSet;
        }

        var paginationExists = document.querySelector(".pagination") || document.querySelector("[data-page]") || document.querySelector(".dataTables_paginate");
        if (paginationExists) {
            log("ImportIE: collectAllTableCodes pagination detected, iterating pages");
            var maxPages = 50;
            var pageNum = 0;
            while (pageNum < maxPages) {
                var rows = tbody.querySelectorAll("tr");
                var ri = 0;
                while (ri < rows.length) {
                    var tr = rows[ri];
                    var anchor = tr.querySelector("a[href*='/secure/crfdesign/studylibrary/show/item/']");
                    if (anchor) {
                        var anchorText = (anchor.textContent + "").trim();
                        var code = extractIECodeStrict(anchorText);
                        if (code.length > 0) {
                            codeSet.add(code.toUpperCase());
                        } else {
                            var cleaned = anchorText.replace(/\s+/g, "").toUpperCase();
                            if (cleaned.length > 0) {
                                codeSet.add(cleaned);
                            }
                        }
                    } else {
                        var tds = tr.querySelectorAll("td");
                        if (tds && tds.length > 0) {
                            var firstText = (tds[0].textContent + "").trim().replace(/\s+/g, " ");
                            var code2 = extractIECodeStrict(firstText);
                            if (code2.length > 0) {
                                codeSet.add(code2.toUpperCase());
                            }
                        }
                    }
                    ri = ri + 1;
                }
                var nextBtn = document.querySelector(".pagination .next:not(.disabled) a");
                if (!nextBtn) {
                    nextBtn = document.querySelector(".dataTables_paginate .next:not(.disabled)");
                }
                if (!nextBtn) {
                    log("ImportIE: collectAllTableCodes no more pages");
                    break;
                }
                nextBtn.click();
                await sleep(1500);
                tbody = document.querySelector("tbody#eligibilityreftablebody") || document.querySelector("tbody#eligibilityRefTableBody") || document.querySelector("#eligibilityRefTable tbody");
                if (!tbody) {
                    log("ImportIE: collectAllTableCodes tbody lost after page change");
                    break;
                }
                pageNum = pageNum + 1;
            }
        } else {
            var loadMoreBtn = document.querySelector(".load-more, [data-load-more], button:contains('Load More')");
            if (loadMoreBtn) {
                log("ImportIE: collectAllTableCodes load-more detected");
                var clickCount = 0;
                var maxClicks = 50;
                while (clickCount < maxClicks) {
                    loadMoreBtn.click();
                    await sleep(1500);
                    loadMoreBtn = document.querySelector(".load-more, [data-load-more]");
                    if (!loadMoreBtn) {
                        break;
                    }
                    clickCount = clickCount + 1;
                }
            }
            var rows2 = tbody.querySelectorAll("tr");
            log("ImportIE: collectAllTableCodes rows=" + String(rows2.length));
            var ri2 = 0;
            while (ri2 < rows2.length) {
                var tr2 = rows2[ri2];
                var anchor2 = tr2.querySelector("a[href*='/secure/crfdesign/studylibrary/show/item/']");
                if (anchor2) {
                    var anchorText2 = (anchor2.textContent + "").trim();
                    var code3 = extractIECodeStrict(anchorText2);
                    if (code3.length > 0) {
                        codeSet.add(code3.toUpperCase());
                    } else {
                        var cleaned2 = anchorText2.replace(/\s+/g, "").toUpperCase();
                        if (cleaned2.length > 0) {
                            codeSet.add(cleaned2);
                        }
                    }
                } else {
                    var tds2 = tr2.querySelectorAll("td");
                    if (tds2 && tds2.length > 0) {
                        var firstText2 = (tds2[0].textContent + "").trim().replace(/\s+/g, " ");
                        var code4 = extractIECodeStrict(firstText2);
                        if (code4.length > 0) {
                            codeSet.add(code4.toUpperCase());
                        }
                    }
                }
                ri2 = ri2 + 1;
            }
        }
        log("ImportIE: collectAllTableCodes done count=" + String(codeSet.size));
        log("ImportIE: collectAllTableCodes codes=" + JSON.stringify(Array.from(codeSet)));
        return codeSet;
    }


    function clearSelectedVolunteerIds() {
        try {
            localStorage.removeItem(STORAGE_SELECTED_IDS);
            log("SelectedIds cleared");
        } catch (e) {}
    }

    function clearPendingIds() {
        try {
            localStorage.removeItem(STORAGE_PENDING);
            log("Pending IDs cleared");
        } catch (e) {}
    }

    function clearRunMode() {
        try {
            localStorage.removeItem(STORAGE_RUN_MODE);
            log("RunMode cleared");
        } catch (e) {}
    }

    function clearContinueEpoch() {
        try {
            localStorage.removeItem(STORAGE_CONTINUE_EPOCH);
            log("ContinueEpoch cleared");
        } catch (e) {}
    }
    function clearCohortGuard() {
        try {
            localStorage.removeItem("activityPlanState.cohortAdd.guard");
            localStorage.removeItem("activityPlanState.cohortAdd.editDoneMap");
            log("CohortGuard and editDoneMap cleared");
        } catch (e) {}
    }

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

    function clearConsentScanIndex() {
        try {
            localStorage.removeItem(STORAGE_CONSENT_SCAN_INDEX);
        } catch (e) {}
    }
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

    function SharedPanelFunctions() {}

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
        title.textContent = "ClinSpark Eligibility Mapping Automator";
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
        runBarcodeBtn.textContent = "8. Run Barcode";
        runBarcodeBtn.style.background = "#9df";
        runBarcodeBtn.style.color = "#000";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
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
        clearMappingBtn.style.background = "#f99";
        clearMappingBtn.style.color = "#000";
        clearMappingBtn.style.border = "none";
        clearMappingBtn.style.borderRadius = "6px";
        clearMappingBtn.style.padding = "8px";
        clearMappingBtn.style.cursor = "pointer";


        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(pauseBtn);
        btnRow.appendChild(clearLogsBtn);
        btnRow.appendChild(toggleLogsBtn);
        btnRow.appendChild(importEligBtn);
        btnRow.appendChild(clearMappingBtn);

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
        } catch (e) {
        }
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

        importEligBtn.addEventListener("click", function () {
            log("ImportElig: button clicked");
            startImportEligibilityMapping();
        });


        clearMappingBtn.addEventListener("click", function () {
            log("ClearMapping: button clicked");
            startClearMapping();
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
            } catch (e) {
            }
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
            } catch (e3) {
            }
            try {
                localStorage.setItem("activityPlanState.panel.right", panel.style.right);
            } catch (e4) {
            }
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



    function init() {
        log("Init: starting");
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
            log("Init: barcode subjects page detected; processing");
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
            log("ImportIE: init detected RUNMODE_ELIG_IMPORT, clearing stale run mode");
            try {
                localStorage.removeItem(STORAGE_RUN_MODE);
            } catch (e) {
            }
            try {
                localStorage.removeItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
            } catch (e) {
            }
            return;
        }

        if (runModeRaw === RUNMODE_CLEAR_MAPPING) {
            if (!isEligibilityListPage()) {
                log("ClearMapping: run mode set but not on list page; redirecting");
                location.href = getBaseUrl() + ELIGIBILITY_LIST_PATH;
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
