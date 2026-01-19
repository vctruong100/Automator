
// ==UserScript==
// @name        ClinSpark Automator
// @namespace   vinh.activity.plan.state
// @version     1.3.1
// @description Retain only Barcode feature; production environment only
// @match       https://cenexel.clinspark.com/*
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
    var AE_LIST_TEST_BASE_URL = "https://cenexeltest.clinspark.com/secure/study/data/list?search=true";
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
        pauseBtn.style.background = "#ccc";
        pauseBtn.style.color = "#000";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = "6px";
        pauseBtn.style.padding = "8px";
        pauseBtn.style.cursor = "pointer";
        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Run Barcode";
        runBarcodeBtn.style.background = "#9df";
        runBarcodeBtn.style.color = "#000";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = "6px";
        runBarcodeBtn.style.padding = "8px";
        runBarcodeBtn.style.cursor = "pointer";
        var findAeBtn = document.createElement("button");
        findAeBtn.textContent = "Find Adverse Event";
        findAeBtn.style.background = "#e66";
        findAeBtn.style.color = "#fff";
        findAeBtn.style.border = "none";
        findAeBtn.style.borderRadius = "6px";
        findAeBtn.style.padding = "8px";
        findAeBtn.style.cursor = "pointer";
        var findFormBtn = document.createElement("button");
        findFormBtn.textContent = "Find Form";
        findFormBtn.style.background = "#6c8";
        findFormBtn.style.color = "#000";
        findFormBtn.style.border = "none";
        findFormBtn.style.borderRadius = "6px";
        findFormBtn.style.padding = "8px";
        findFormBtn.style.cursor = "pointer";
        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#555";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = "6px";
        toggleLogsBtn.style.padding = "8px";
        toggleLogsBtn.style.cursor = "pointer";
        btnRow.appendChild(runBarcodeBtn);
        btnRow.appendChild(findAeBtn);
        btnRow.appendChild(findFormBtn);
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
        }
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


    function isBarcodeSubjectsPageRouteOnly() {
        var path = location.pathname;
        if (path === "/secure/barcodeprinting/subjects") {
            return true;
        }
        return false;
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
    }


    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();

