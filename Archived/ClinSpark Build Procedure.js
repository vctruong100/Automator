
// ==UserScript==
// @name ClinSpark Build Procedure Log
// @namespace vinh.activity.plan.state
// @version 1.0.0
// @description Build Procedure Log automation
// @match https://cenexeltest.clinspark.com/*
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// ==/UserScript==

(function () {
    var STORAGE_PANEL_WIDTH = "activityPlanState.panel.width";
    var STORAGE_PANEL_HEIGHT = "activityPlanState.panel.height";
    // UI Scale Constants
    var UI_SCALE = 1.0; // Master scale factor (will be initialized after function definitions)
    var PANEL_DEFAULT_WIDTH = 340;
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 50;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 60;
    var PANEL_PADDING_PX = 12;
    var PANEL_BORDER_RADIUS_PX = 8;
    var PANEL_FONT_SIZE_PX = 14;
    var HEADER_FONT_WEIGHT = "600";
    var HEADER_TITLE_OFFSET_PX = 16;
    var HEADER_LEFT_SPACER_WIDTH_PX = 16;
    var BUTTON_BORDER_RADIUS_PX = 6;
    var BUTTON_PADDING_PX = 8;
    var BUTTON_GAP_PX = 8;
    var STATUS_MARGIN_TOP_PX = 10;
    var STATUS_PADDING_PX = 6;
    var STATUS_FONT_SIZE_PX = 13;
    var STATUS_BORDER_RADIUS_PX = 6;
    var LOG_MARGIN_TOP_PX = 8;
    var LOG_HEIGHT_PX = 220;
    var LOG_PADDING_PX = 6;
    var LOG_FONT_SIZE_PX = 12;
    var LOG_BORDER_RADIUS_PX = 6;
    var CLOSE_BTN_TEXT = "✕";
    var STORAGE_UI_SCALE = "activityPlanState.ui.scale";

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
    const STORAGE_PANEL_HOTKEY = "activityPlanState.panel.hotkey";
    const PANEL_TOGGLE_KEY = "F2";

    // Run Subject Eligibility
    const ELIGIBILITY_LIST_URL = "https://cenexeltest.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const STORAGE_ELIG_IMPORTED = "activityPlanState.eligibility.importedItems";
    const RUNMODE_ELIG_IMPORT = "eligibilityImport";
    const STORAGE_ELIG_CHECKITEM_CACHE = "activityPlanState.eligibility.checkItemCache";
    const STORAGE_ELIG_IMPORT_PENDING_POPUP = "activityPlanState.eligibility.importPendingPopup";
    const ELIGIBILITY_LIST_URL_PROD = "https://cenexel.clinspark.com/secure/crfdesign/studylibrary/eligibility/list";
    const ELIGIBILITY_VALID_HOSTNAMES = ["cenexeltest.clinspark.com", "cenexel.clinspark.com"];
    const ELIGIBILITY_LIST_PATH = "/secure/crfdesign/studylibrary/eligibility/list";
    const IE_CODE_REGEX = /\b(INC|EXC)\s*(\d+)([a-zA-Z]?)\b/i;
    const IE_CODE_REGEX_GLOBAL = /\b(INC|EXC)\s*(\d+)([a-zA-Z]?)\b/gi;
    const IMPORT_IE_HELPER_TIMEOUT = 15000;
    const IMPORT_IE_POLL_INTERVAL = 120;
    const IMPORT_IE_MODAL_TIMEOUT = 12000;
    const IMPORT_IE_SHORT_DELAY_MIN = 150;
    const IMPORT_IE_SHORT_DELAY_MAX = 400;
    var IMPORT_IE_CANCELED = false;
    var IMPORT_IE_COLLECTION_STOPPED = false;
    var RIGHT_PANEL_MODE = { HIERARCHY: "hierarchy", CODE: "code" };
    var CODE_SORT_TOGGLE_BUTTON_ID = "ieCodeSortToggle";
    var RIGHT_PANEL_BODY_ID = "ieRightPanelBody";
    var CODE_TOKEN_REGEX = /\b(INC|EXC)\s*[-_:/#]?\s*0*(\d+)([a-zA-Z]?)\b/i;
    var CODE_PAD_WIDTH = 3;
    var CODE_TOKEN_REGEX_POOL = /\b(INC|EXC)\s*[-_:/#]?\s*0*(\d+)([a-zA-Z]?)\b/i;
    var CODE_PAD_WIDTH_POOL = 3;
    var ELIG_POOL_PANEL_ID = "ieEligPoolPanel";
    var ELIG_POOL_LIST_ID = "ieEligPoolList";
    var ELIG_DROPBOX_CLASS = "ie-elig-dropbox";
    var ELIG_DRAG_TYPE = "application/x-elig-item-value";
    var ELIG_POOL_LABEL_WORD_LIMIT = 10;
    var GENDER_BOTH = "Both";
    var GENDER_MALE = "Male";
    var GENDER_FEMALE = "Female";
    var GENDER_CONTROL_CLASS = "ieGenderControl";
    var CODE_VIEW_DUP_TOGGLE_BUTTON_ID = "ieCodeViewDupToggle";

    // Theme Mode
    var STORAGE_THEME_MODE = "activityPlanState.themeMode";
    var THEME_MODE_BLACK = "black";
    var THEME_MODE_GLASS = "glass";

    // Glassmorphism Theme Variables (from Test Automator 2)
    var THEME_GRADIENT_START = "#667eea";
    var THEME_GRADIENT_END = "#764ba2";
    var THEME_GRADIENT_BG = "linear-gradient(135deg, #667eea 0%, #764ba2 100%)";
    var THEME_SURFACE_BG = "rgba(15,10,40,0.55)";
    var THEME_SURFACE_BG_HEAVY = "rgba(15,10,40,0.7)";
    var THEME_SURFACE_BORDER = "rgba(255,255,255,0.25)";
    var THEME_SURFACE_INNER_BORDER = "rgba(255,255,255,0.35)";
    var THEME_BLUR_PX = 16;
    var THEME_TEXT_PRIMARY = "#ffffff";
    var THEME_TEXT_MUTED = "rgba(255,255,255,0.75)";
    var THEME_TEXT_INVERSE = "#0f1020";
    var THEME_ACCENT = "#a78bfa";
    var THEME_ACCENT_HOVER = "#c4b5fd";
    var THEME_DANGER = "#ef4444";
    var THEME_DANGER_DARK = "#b91c1c";
    var THEME_WARNING = "#f59e0b";
    var THEME_WARNING_DARK = "#d97706";
    var THEME_SUCCESS = "#10b981";
    var THEME_SUCCESS_DARK = "#059669";
    var THEME_DISABLED_OPACITY = 0.5;
    var THEME_SELECT_OPTION_BG = "#2d1b69";
    var THEME_SHADOW = "0 8px 30px rgba(0,0,0,0.3)";
    var THEME_RADIUS = 14;
    var THEME_OUTLINE_FOCUS = "0 0 0 3px rgba(199,210,254,0.6)";
    var THEME_SCROLLBAR_TRACK = "rgba(255,255,255,0.12)";
    var THEME_SCROLLBAR_THUMB = "rgba(255,255,255,0.35)";
    var THEME_STYLE_TAG_ID = "ieThemeStyles";
    var THEME_SCOPE_CLASS = "ie-theme-scope";
    var THEME_Z_BASE = 999999;
    var THEME_Z_OVERLAY = 1000000;
    var THEME_Z_TOAST = 1000001;

    function getThemeMode() {
        try {
            var saved = localStorage.getItem(STORAGE_THEME_MODE);
            if (saved === THEME_MODE_GLASS) return THEME_MODE_GLASS;
        } catch (e) {}
        return THEME_MODE_BLACK;
    }

    function setThemeMode(mode) {
        try {
            localStorage.setItem(STORAGE_THEME_MODE, mode);
        } catch (e) {}
    }

    function isGlassTheme() {
        return getThemeMode() === THEME_MODE_GLASS;
    }

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
    var STUDY_EVENT_POPUP_DESCRIPTION = "Auto-navigate to Study Events data page based on keywords and status";
    var STUDY_EVENT_POPUP_KEYWORD_LABEL = "Study Event Keyword";
    var STUDY_EVENT_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var STUDY_EVENT_POPUP_OK_TEXT = "Continue";
    var STUDY_EVENT_POPUP_CANCEL_TEXT = "Cancel";

    var STUDY_EVENT_NO_MATCH_TITLE = "Find Study Events";
    var STUDY_EVENT_NO_MATCH_MESSAGE = "No study event is found.";

    // Run Find Form
    var FORM_LIST_URL = "https://cenexeltest.clinspark.com/secure/study/data/list";
    var FORM_POPUP_TITLE = "Find Form";
    var FORM_POPUP_DESCRIPTION = "Auto-navigate to Form data page based on keywords and status";
    var FORM_POPUP_KEYWORD_LABEL = "Form Keyword";
    var FORM_POPUP_SUBJECT_LABEL = "Subject Identifier";
    var FORM_POPUP_OK_TEXT = "Continue";
    var FORM_POPUP_CANCEL_TEXT = "Cancel";

    var FORM_NO_MATCH_TITLE = "Find Form";
    var FORM_NO_MATCH_MESSAGE = "No form is found.";

    var BARCODE_BG_TAB = null;
    var BARCODE_ADMIN_PATH = "/secure/administration/studies/show";
    var SUBJECT_SHOW_PATH_PREFIX = "/secure/study/subjects/show/";
    var MODAL_ID = "dataCollectionModal";
    var BARCODE_ICON_ID = "icfBarcodeIcon";
    var BARCODE_VALIDATE_REGEX = /^[A-Za-z0-9]{3,}$/;
    var OPEN_MODAL_TIMEOUT_MS = 12000;
    var FETCH_BARCODE_TIMEOUT_MS = 12000;
    var CLICK_RETRY_MAX = 2;
    var RETRY_DELAY_MS_MIN = 150;
    var RETRY_DELAY_MS_MAX = 400;
    var ICF_ALLOWED_HOSTNAMES = ["cenexeltest.clinspark.com", "cenexel.clinspark.com"];
    const RUNMODE_CLEAR_MAPPING = "clearMapping";

    // Run Parse Method
    var PARSE_METHOD_POPUP_TITLE = "Parse Method";
    var PARSE_METHOD_POPUP_DESCRIPTION = "Auto-navigate to data page with forms that has methods that pulls this item.";
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
    var RUN_ALL_POPUP_TITLE = "Run All";
    var RUN_ALL_POPUP_DESCRIPTION = "Auto-navigate to Run All data page based on keywords and status";
    const STORAGE_RUN_ALL_POPUP = "activityPlanState.runAllPopup";
    const STORAGE_ICF_BARCODE_POPUP = "activityPlanState.icfBarcodePopup";
    const STORAGE_RUN_ALL_STATUS = "activityPlanState.runAllStatus";
    const STORAGE_CLEAR_MAPPING_POPUP = "activityPlanState.clearMappingPopup";
    const STORAGE_IMPORT_ELIG_POPUP = "activityPlanState.importEligPopup";
    const STORAGE_IMPORT_COHORT_POPUP = "activityPlanState.importCohortPopup";
    const STORAGE_ADD_COHORT_POPUP = "activityPlanState.addCohortPopup";

    // Run Subject Eligibility
    var ELIGIBILITY_POPUP_TITLE = "Run Subject Eligibility";
    var ELIGIBILITY_POPUP_DESCRIPTION = "Auto-navigate to Run Subject Eligibility data page based on keywords and status";
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

    var CLEAR_MAPPING_CANCELED = false;
    var CLEAR_MAPPING_PAUSE = false;

    // Lab Barcode Feature

    var BARCODE_SELECTORS = {
        featureButtonTarget: '.main-gui-panel',
        presenceCheckAnyIcon: 'i.fa-barcode, i.fa.fa-barcode',
        barcodeIcon: 'i.fa-barcode[class*="itemDataBarcodeIcon_"], i.fa.fa-barcode[class*="itemDataBarcodeIcon_"]',
        verifiedClass: 'barcodeVerifiedClass',
        tooltipAttr: 'data-original-title',
        modalInput: 'input.bootbox-input.bootbox-input-text.form-control',
        modalConfirmPrimary: 'button[data-bb-handler="confirm"].btn.btn-primary',
        modalAnyPrimary: 'div.bootbox.modal.in .modal-footer .btn.btn-primary, div.bootbox.modal.show .modal-footer .btn.btn-primary',
        bootboxVisibleModal: 'div.bootbox.modal.in, div.bootbox.modal.show',
        bootboxBody: '.bootbox-body',
        bootboxOkBtn: 'button[data-bb-handler="ok"].btn.btn-primary',
        itemRow: 'tr',
        itemTextCell: 'td.itemText',
        itemHelpText: '.itemHelpText',
        ariaLiveRegion: '.aria-live-region'
    };
    var BARCODE_TIMEOUTS = {
        waitProgressPanelMs: 10000,
        waitCollectIconsMs: 8000,
        waitModalOpenMs: 8000,
        waitModalCloseMs: 6000,
        waitSettleMs: 200,
        waitVerifyIconMs: 1200,
        idleBetweenItemsMs: 120,
        maxTotalDurationMs: 300000
    };
    var BARCODE_RETRY = {
        collectRetries: 2,
        openModalRetries: 2,
        fillConfirmRetries: 2,
        verifyRetries: 2
    };
    var BARCODE_LABELS = {
        featureButton: 'Pull Lab Barcode',
        progressTitle: 'Pull Lab Barcode',
        statusPending: 'Pending',
        statusProcessing: 'Processing',
        statusVerified: 'Verified',
        statusSkippedVerified: 'Skipped (Already Verified)',
        statusFailed: 'Failed',
        statusInvalidBarcode: 'Failed (Invalid Barcode)',
        statusStopped: 'Stopped',
        scanning: 'Collecting barcode icons',
        filling: 'Filling barcode',
        confirming: 'Confirming',
        verifying: 'Verifying',
        done: 'Completed'
    };
    var BARCODE_COUNTERS = {
        total: 0,
        processed: 0,
        verified: 0,
        skipped: 0,
        failures: 0,
        pending: 0
    };
    var BARCODE_REGEX = {
        barcodeClassToken: /\bitemDataBarcodeIcon_([A-Za-z0-9_-]+)\b/,
        requiredPrefix: /^Scan Required:?/i,
        verifiedPrefix: /^Verified:/i,
        okText: /^(ok|confirm)$/i
    };
    var BARCODE_ATTRS = {
        ariaBusyTarget: 'body',
        ariaBusyAttr: 'aria-busy'
    };
    var PULL_LAB_BARCODE_STOPPED = false;
    var PULL_LAB_BARCODE_RUNNING = false;
    var PULL_LAB_BARCODE_POPUP_REF = null;
    var PULL_LAB_BARCODE_TIMEOUTS_LIST = [];
    var PULL_LAB_BARCODE_INTERVALS_LIST = [];
    var PULL_LAB_BARCODE_OBSERVERS_LIST = [];
    var PULL_LAB_BARCODE_RAF_IDS = [];
    var PULL_LAB_BARCODE_IDLE_IDS = [];
    var PULL_LAB_BARCODE_BUTTON_REF = null;
    var PULL_LAB_BARCODE_ARIA_LIVE_EL = null;
    var PULL_LAB_BARCODE_LEFT_LIST_EL = null;
    var PULL_LAB_BARCODE_RIGHT_LIST_EL = null;
    var PULL_LAB_BARCODE_SUMMARY_EL = null;
    var PULL_LAB_BARCODE_STATUS_MAP = {};

    //==========================
    // BUILD PROCEDURE LOG FEATURE
    //==========================
    // This section contains all functions related to the BUILD PROCEDURE LOG.
    // This feature automates adding procedure log entries.
    //==========================


    
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
    var STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES = "activityPlanState.saBuilder.segmentCheckboxes";
    var SA_BUILDER_POPUP_REF = null;
    var SA_BUILDER_PROGRESS_POPUP_REF = null;
    var SA_BUILDER_CANCELLED = false;
    var SA_BUILDER_PAUSE = false;
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
            localStorage.removeItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES);
        } catch (e) {}
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

    async function selectSelect2Value(selectId, value) {
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled during selectSelect2Value");
            return false;
        }

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

        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled after selectSelect2Value sleep");
            return false;
        }

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
        var segmentColumnWrapper = document.createElement("div");
        segmentColumnWrapper.style.cssText = "display:flex;flex-direction:column;border:1px solid #333;border-radius:4px;background:#1a1a1a;height:100%;overflow:hidden;";

        // Select All header for segments
        var segmentHeader = document.createElement("div");
        segmentHeader.style.cssText = "display:flex;align-items:center;gap:8px;padding:8px;border-bottom:1px solid #333;background:#222;";

        var selectAllCheckbox = document.createElement("input");
        selectAllCheckbox.type = "checkbox";
        selectAllCheckbox.id = "saBuilderSelectAllSegments";
        selectAllCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;";

        var selectAllLabel = document.createElement("span");
        selectAllLabel.textContent = "Select All";
        selectAllLabel.style.cssText = "font-weight:600;font-size:13px;";

        segmentHeader.appendChild(selectAllCheckbox);
        segmentHeader.appendChild(selectAllLabel);

        var segmentColumn = document.createElement("div");
        segmentColumn.style.cssText = "overflow-y:auto;padding:8px;flex:1;";
        segmentColumn.id = "saBuilderSegments";

        segmentColumnWrapper.appendChild(segmentHeader);
        segmentColumnWrapper.appendChild(segmentColumn);

        // Study Events column (draggable items)
        var eventColumn = document.createElement("div");
        eventColumn.style.cssText = "overflow-y:auto;border:1px solid #444;border-radius:6px;padding:10px;background:#1e1e1e;box-shadow:inset 0 1px 3px rgba(0,0,0,0.3);";
        eventColumn.id = "saBuilderEvents";

        // Make Study Events column a drop zone for restoring events
        eventColumn.addEventListener("dragover", function(e) {
            e.preventDefault();
            this.style.background = "#252525";
            this.style.border = "2px dashed #007bff";
            this.style.transition = "all 0.2s ease";
        });

        eventColumn.addEventListener("dragleave", function(e) {
            e.preventDefault();
            this.style.background = "#1e1e1e";
            this.style.border = "1px solid #444";
            this.style.transition = "all 0.2s ease";
        });

        eventColumn.addEventListener("drop", function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.style.background = "#1e1e1e";
            this.style.border = "1px solid #444";
            this.style.transition = "all 0.2s ease";

            var eventData = e.dataTransfer.getData("text/plain");
            if (eventData) {
                var data = JSON.parse(eventData);
                if (data.fromSegment) {
                    // Remove from segment
                    if (data.segmentValue && segmentEventMap[data.segmentValue]) {
                        var events = segmentEventMap[data.segmentValue];
                        var index = events.findIndex(function(ev) {
                            return ev.value === data.value;
                        });
                        if (index > -1) {
                            events.splice(index, 1);
                            segmentEventMap[data.segmentValue] = events;
                            // Auto-uncheck segment if no events remain, otherwise just refresh
                            syncSegmentCheckbox(data.segmentValue);
                        }
                    }
                    // Restore to column
                    restoreEventToStudyEventsColumn(data);
                    log("SA Builder: Event '" + data.text + "' moved back to Study Events column");
                }
            }
        });

        // Forms column (checkboxes)
        var formColumn = document.createElement("div");
        formColumn.style.cssText = "overflow-y:auto;border:1px solid #333;border-radius:4px;padding:8px;background:#1a1a1a;";
        formColumn.id = "saBuilderForms";

        // Time inputs column
        var timeColumn = document.createElement("div");
        timeColumn.style.cssText = "display:flex;flex-direction:column;gap:12px;padding:8px;border:1px solid #333;border-radius:4px;background:#1a1a1a;height:100%;overflow-y:auto;overflow-x:hidden;";
        timeColumn.id = "saBuilderTime";

        // Track segment-event assignments
        var segmentEventMap = {};

        // Track segment checkbox states separately to prevent resets
        var segmentCheckboxStates = {};

        // Auto-check/uncheck a segment checkbox based on whether it has study events
        function syncSegmentCheckbox(segVal) {
            var hasEvents = segmentEventMap[segVal] && segmentEventMap[segVal].length > 0;
            segmentCheckboxStates[segVal] = hasEvents;
            try {
                localStorage.setItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES, JSON.stringify(segmentCheckboxStates));
            } catch (e) {}
            renderSegments(segmentSearch.value);
            updateSelectAllCheckbox();
        }

        // Load saved checkbox states
        try {
            var savedStates = localStorage.getItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES);
            if (savedStates) {
                segmentCheckboxStates = JSON.parse(savedStates);
            }
        } catch (e) {}

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

                // Restore checkbox state from saved states
                if (segmentCheckboxStates[seg.value]) {
                    checkbox.checked = true;
                }

                // Save checkbox state when changed
                checkbox.addEventListener("change", function() {
                    segmentCheckboxStates[this.dataset.segmentValue] = this.checked;
                    try {
                        localStorage.setItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES, JSON.stringify(segmentCheckboxStates));
                    } catch (e) {}
                    updateSelectAllCheckbox();
                });

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

                            var attachedItem = document.createElement("div");
                            attachedItem.style.cssText = "display:flex;align-items:center;justify-content:space-between;padding:8px 10px;margin-bottom:4px;border:1px solid #666;border-radius:5px;background:#3a3a3a;color:#fff;font-size:13px;font-weight:500;cursor:move;transition:all 0.2s ease;box-shadow:0 1px 3px rgba(0,0,0,0.2);";
                            attachedItem.draggable = true;
                            attachedItem.dataset.eventValue = ev.value;
                            attachedItem.dataset.eventText = ev.text;

                            var eventLabel = document.createElement("span");
                            eventLabel.textContent = ev.text;
                            eventLabel.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";

                            // Make attached events draggable back to Study Events column
                            attachedItem.addEventListener("dragstart", function(e) {
                                e.dataTransfer.setData("text/plain", JSON.stringify({
                                    value: this.dataset.eventValue,
                                    text: this.dataset.eventText,
                                    fromSegment: true,
                                    segmentValue: segVal
                                }));
                                e.dataTransfer.effectAllowed = "move";
                                this.style.opacity = "0.7";
                                this.style.transform = "scale(0.95)";
                                this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
                            });

                            attachedItem.addEventListener("dragend", function(e) {
                                this.style.opacity = "1";
                                this.style.transform = "scale(1)";
                                this.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
                            });

                            attachedItem.addEventListener("mouseenter", function() {
                                this.style.background = "#4a4a4a";
                                this.style.border = "1px solid #777";
                                this.style.transform = "translateY(-1px)";
                                this.style.boxShadow = "0 2px 5px rgba(0,0,0,0.3)";
                            });

                            attachedItem.addEventListener("mouseleave", function() {
                                this.style.background = "#3a3a3a";
                                this.style.border = "1px solid #666";
                                this.style.transform = "translateY(0)";
                                this.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
                            });

                            var removeBtn = document.createElement("button");
                            removeBtn.textContent = "×";
                            removeBtn.style.cssText = "background:transparent;border:none;color:#ff6b6b;cursor:pointer;font-size:18px;font-weight:bold;padding:0;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;transition:all 0.2s ease;flex-shrink:0;margin-left:8px;";

                            removeBtn.addEventListener("mouseenter", function() {
                                this.style.background = "rgba(255,107,107,0.2)";
                                this.style.color = "#ff5252";
                                this.style.transform = "scale(1.1)";
                            });

                            removeBtn.addEventListener("mouseleave", function() {
                                this.style.background = "transparent";
                                this.style.color = "#ff6b6b";
                                this.style.transform = "scale(1)";
                            });

                            removeBtn.addEventListener("click", function(e) {
                                e.stopPropagation();
                                var index = events.indexOf(ev);
                                if (index > -1) {
                                    events.splice(index, 1);
                                    segmentEventMap[segVal] = events;
                                    renderAttachedEvents(dz, segVal);
                                    log("SA Builder: Event '" + ev.text + "' removed from segment");

                                    // Restore event to Study Events column
                                    restoreEventToStudyEventsColumn({ value: ev.value, text: ev.text });

                                    // Auto-uncheck segment if no events remain
                                    syncSegmentCheckbox(segVal);
                                }
                            });

                            attachedItem.appendChild(eventLabel);
                            attachedItem.appendChild(removeBtn);
                            dz.appendChild(attachedItem);

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
                }

                renderAttachedEvents(dropzone, seg.value);

                // Handle dragover
                dropzone.addEventListener("dragover", function(e) {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = "copy";
                    this.style.background = "#2a2a4a";
                });

                dropzone.addEventListener("dragleave", function(e) {
                    this.style.background = "#1a1a1a";
                });

                // Handle drop
                dropzone.addEventListener("drop", (function(segVal) {
                    return function(e) {
                        e.preventDefault();
                        this.style.background = "#1a1a1a";
                        try {
                            var data = JSON.parse(e.dataTransfer.getData("text/plain"));
                            if (data && data.value && data.text) {
                                // Remove from old segment if exists
                                if (data.fromSegment && segmentEventMap[data.fromSegment]) {
                                    segmentEventMap[data.fromSegment] = segmentEventMap[data.fromSegment].filter(function(x) {
                                        return x.value !== data.value;
                                    });
                                }
                                // Check for duplicate in this segment
                                var existing = (segmentEventMap[segVal] || []).find(function(x) { return x.value === data.value; });
                                if (!existing) {
                                    if (!segmentEventMap[segVal]) segmentEventMap[segVal] = [];
                                    segmentEventMap[segVal].push({ value: data.value, text: data.text });

                                    removeEventFromStudyEventsColumn(data.value);
                                }
                                // Auto-uncheck old segment if it lost its last event
                                if (data.fromSegment && data.fromSegment !== segVal) {
                                    syncSegmentCheckbox(data.fromSegment);
                                }
                                // Auto-check this segment since it now has events
                                syncSegmentCheckbox(segVal);
                            }
                        } catch (err) {
                            log("SA Builder: drop error " + err);
                        }
                    };
                })(seg.value));

                segItem.appendChild(dropzone);
                segmentColumn.appendChild(segItem);
            }
        }

        // Populate study events (draggable)
        function renderStudyEvents(filter) {
            eventColumn.innerHTML = "";
            var filterLower = (filter || "").toLowerCase();
            for (var i = 0; i < studyEvents.length; i++) {
                var ev = studyEvents[i];
                if (filterLower && ev.text.toLowerCase().indexOf(filterLower) === -1) continue;

                var evItem = document.createElement("div");
                evItem.style.cssText = "padding:10px 12px;margin-bottom:6px;border:1px solid #555;border-radius:6px;background:#2a2a2a;cursor:grab;color:#fff;font-size:14px;font-weight:500;transition:all 0.2s ease;box-shadow:0 2px 4px rgba(0,0,0,0.2);";
                evItem.textContent = ev.text;
                evItem.draggable = true;
                evItem.dataset.eventValue = ev.value;
                evItem.dataset.eventText = ev.text;

                evItem.addEventListener("mouseenter", function() {
                    this.style.background = "#333";
                    this.style.border = "1px solid #666";
                    this.style.transform = "translateY(-1px)";
                    this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
                });

                evItem.addEventListener("mouseleave", function() {
                    this.style.background = "#2a2a2a";
                    this.style.border = "1px solid #555";
                    this.style.transform = "translateY(0)";
                    this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
                });

                evItem.addEventListener("dragstart", function(e) {
                    e.dataTransfer.setData("text/plain", JSON.stringify({
                        value: this.dataset.eventValue,
                        text: this.dataset.eventText
                    }));
                    e.dataTransfer.effectAllowed = "copy";
                    this.style.opacity = "0.6";
                    this.style.transform = "scale(0.95)";
                    this.style.boxShadow = "0 6px 12px rgba(0,0,0,0.4)";
                });

                evItem.addEventListener("dragend", function(e) {
                    this.style.opacity = "1";
                    this.style.transform = "scale(1)";
                    this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
                });

                eventColumn.appendChild(evItem);
                sortStudyEventsColumn();
            }
        }

        // Remove event from Study Events column
        function removeEventFromStudyEventsColumn(eventValue) {
            var eventItems = eventColumn.querySelectorAll("[data-event-value]");
            for (var i = 0; i < eventItems.length; i++) {
                var item = eventItems[i];
                if (item.dataset.eventValue === eventValue) {
                    item.remove();
                    log("SA Builder: Event removed from Study Events column");
                    break;
                }
            }
        }
        // Restore event to Study Events column
        function restoreEventToStudyEventsColumn(eventData) {
            // Check if event already exists in column
            var existingItems = eventColumn.querySelectorAll("[data-event-value]");
            for (var i = 0; i < existingItems.length; i++) {
                if (existingItems[i].dataset.eventValue === eventData.value) {
                    log("SA Builder: Event already exists in Study Events column");
                    return;
                }
            }

            // Create the event item
            var evItem = document.createElement("div");
            evItem.textContent = eventData.text;
            evItem.style.cssText = "padding:10px 12px;margin-bottom:6px;border:1px solid #555;border-radius:6px;background:#2a2a2a;cursor:grab;color:#fff;font-size:14px;font-weight:500;transition:all 0.2s ease;box-shadow:0 2px 4px rgba(0,0,0,0.2);animation:slideIn 0.3s ease;";
            evItem.draggable = true;
            evItem.dataset.eventValue = eventData.value;
            evItem.dataset.eventText = eventData.text;

            evItem.addEventListener("mouseenter", function() {
                this.style.background = "#333";
                this.style.border = "1px solid #666";
                this.style.transform = "translateY(-1px)";
                this.style.boxShadow = "0 4px 8px rgba(0,0,0,0.3)";
            });

            evItem.addEventListener("mouseleave", function() {
                this.style.background = "#2a2a2a";
                this.style.border = "1px solid #555";
                this.style.transform = "translateY(0)";
                this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
            });

            evItem.addEventListener("dragstart", function(e) {
                e.dataTransfer.setData("text/plain", JSON.stringify({
                    value: this.dataset.eventValue,
                    text: this.dataset.eventText
                }));
                e.dataTransfer.effectAllowed = "copy";
                this.style.opacity = "0.6";
                this.style.transform = "scale(0.95)";
                this.style.boxShadow = "0 6px 12px rgba(0,0,0,0.4)";
            });

            evItem.addEventListener("dragend", function(e) {
                this.style.opacity = "1";
                this.style.transform = "scale(1)";
                this.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
            });

            // Add the event to the column and sort
            eventColumn.appendChild(evItem);
            sortStudyEventsColumn();
            log("SA Builder: Event '" + eventData.text + "' restored to Study Events column");
        }

        // Natural sort comparator: splits text into alphabetic and numeric chunks, compares numerically where possible
        function naturalSortCompare(strA, strB) {
            var re = /(\d+|\D+)/g;
            var partsA = strA.match(re) || [""];
            var partsB = strB.match(re) || [""];
            for (var i = 0; i < Math.max(partsA.length, partsB.length); i++) {
                var a = partsA[i] || "";
                var b = partsB[i] || "";
                var numA = parseInt(a, 10);
                var numB = parseInt(b, 10);
                if (!isNaN(numA) && !isNaN(numB)) {
                    if (numA !== numB) return numA - numB;
                } else {
                    var cmp = a.toLowerCase().localeCompare(b.toLowerCase());
                    if (cmp !== 0) return cmp;
                }
            }
            return 0;
        }

        // Sort the Study Events column using natural sort order
        function sortStudyEventsColumn() {
            var allItems = eventColumn.querySelectorAll("[data-event-value]");
            var itemsArray = Array.prototype.slice.call(allItems);

            // Sort by text content using natural sort (so Day 2 < Day 11)
            itemsArray.sort(function(a, b) {
                var textA = a.dataset.eventText || a.textContent || "";
                var textB = b.dataset.eventText || b.textContent || "";
                return naturalSortCompare(textA, textB);
            });

            // Clear and re-append in sorted order
            eventColumn.innerHTML = "";
            itemsArray.forEach(function(item) {
                eventColumn.appendChild(item);
            });

            log("SA Builder: Study Events column sorted (natural order)");
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
                checkbox.style.cssText = "width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-top:2px;accent-color:#007bff;";

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

        // Hidden checkbox row
        var hiddenRow = document.createElement("div");
        hiddenRow.style.cssText = "display:flex;align-items:center;gap:8px;padding-top:12px;margin-top:12px;border-top:1px solid #444;";

        var hiddenCheckbox = document.createElement("input");
        hiddenCheckbox.type = "checkbox";
        hiddenCheckbox.id = "saBuilderHidden";
        hiddenCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var hiddenLabel = document.createElement("label");
        hiddenLabel.textContent = "Hidden?";
        hiddenLabel.setAttribute("for", "saBuilderHidden");
        hiddenLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        hiddenRow.appendChild(hiddenCheckbox);
        hiddenRow.appendChild(hiddenLabel);
        timeColumn.appendChild(hiddenRow);

        // Mandatory checkbox row
        var mandatoryRow = document.createElement("div");
        mandatoryRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;";

        var mandatoryCheckbox = document.createElement("input");
        mandatoryCheckbox.type = "checkbox";
        mandatoryCheckbox.id = "saBuilderMandatory";
        mandatoryCheckbox.checked = true;
        mandatoryCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var mandatoryLabel = document.createElement("label");
        mandatoryLabel.textContent = "Mandatory";
        mandatoryLabel.setAttribute("for", "saBuilderMandatory");
        mandatoryLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        mandatoryRow.appendChild(mandatoryCheckbox);
        mandatoryRow.appendChild(mandatoryLabel);
        timeColumn.appendChild(mandatoryRow);

        // Enforce Data Collection Order checkbox row
        var enforceRow = document.createElement("div");
        enforceRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;";

        var enforceCheckbox = document.createElement("input");
        enforceCheckbox.type = "checkbox";
        enforceCheckbox.id = "saBuilderEnforce";
        enforceCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var enforceLabel = document.createElement("label");
        enforceLabel.textContent = "Enforce Data Collection Order";
        enforceLabel.setAttribute("for", "saBuilderEnforce");
        enforceLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        enforceRow.appendChild(enforceCheckbox);
        enforceRow.appendChild(enforceLabel);
        timeColumn.appendChild(enforceRow);

        // Pre-Window input row
        var preWindowRow = document.createElement("div");
        preWindowRow.style.cssText = "display:flex;flex-direction:column;gap:4px;margin-top:8px;";

        var preWindowLabel = document.createElement("label");
        preWindowLabel.textContent = "Pre-Window";
        preWindowLabel.style.cssText = "font-size:12px;color:#aaa;";

        var preWindowInput = document.createElement("input");
        preWindowInput.type = "text";
        preWindowInput.id = "saBuilderPreWindow";
        preWindowInput.placeholder = "";
        preWindowInput.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;width:100%;";

        preWindowRow.appendChild(preWindowLabel);
        preWindowRow.appendChild(preWindowInput);
        timeColumn.appendChild(preWindowRow);

        // Post-Window input row
        var postWindowRow = document.createElement("div");
        postWindowRow.style.cssText = "display:flex;flex-direction:column;gap:4px;margin-top:8px;";

        var postWindowLabel = document.createElement("label");
        postWindowLabel.textContent = "Post-Window";
        postWindowLabel.style.cssText = "font-size:12px;color:#aaa;";

        var postWindowInput = document.createElement("input");
        postWindowInput.type = "text";
        postWindowInput.id = "saBuilderPostWindow";
        postWindowInput.placeholder = "";
        postWindowInput.style.cssText = "padding:8px;border-radius:4px;border:1px solid #444;background:#222;color:#fff;width:100%;";

        postWindowRow.appendChild(postWindowLabel);
        postWindowRow.appendChild(postWindowInput);
        timeColumn.appendChild(postWindowRow);

        // Reference Activity checkbox row
        var refActivityRow = document.createElement("div");
        refActivityRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-top:8px;padding-top:8px;border-top:1px solid #444;";

        var refActivityCheckbox = document.createElement("input");
        refActivityCheckbox.type = "checkbox";
        refActivityCheckbox.id = "saBuilderRefActivity";
        refActivityCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var refActivityLabel = document.createElement("label");
        refActivityLabel.textContent = "Reference Activity";
        refActivityLabel.setAttribute("for", "saBuilderRefActivity");
        refActivityLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        refActivityRow.appendChild(refActivityCheckbox);
        refActivityRow.appendChild(refActivityLabel);
        timeColumn.appendChild(refActivityRow);

        // Reference Activity disables time inputs and Pre-Reference checkbox
        refActivityCheckbox.addEventListener("change", function() {
            var isChecked = this.checked;
            var daysEl = document.getElementById("saBuilderDays");
            var hoursEl = document.getElementById("saBuilderHours");
            var minutesEl = document.getElementById("saBuilderMinutes");
            var secondsEl = document.getElementById("saBuilderSeconds");
            var preRefEl = document.getElementById("saBuilderPreReference");

            if (daysEl) {
                daysEl.disabled = isChecked;
                daysEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (hoursEl) {
                hoursEl.disabled = isChecked;
                hoursEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (minutesEl) {
                minutesEl.disabled = isChecked;
                minutesEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (secondsEl) {
                secondsEl.disabled = isChecked;
                secondsEl.style.opacity = isChecked ? "0.5" : "1";
            }
            if (preRefEl) {
                preRefEl.disabled = isChecked;
                preRefEl.style.opacity = isChecked ? "0.5" : "1";
                if (isChecked) {
                    preRefEl.checked = false;
                }
            }
        });

        // Time inputs section (moved below other inputs)
        var timeInputsSection = document.createElement("div");
        timeInputsSection.style.cssText = "padding-top:12px;margin-top:12px;border-top:1px solid #444;";

        // Pre-Reference checkbox row (positioned right above time inputs)
        var preRefRow = document.createElement("div");
        preRefRow.style.cssText = "display:flex;align-items:center;gap:8px;margin-bottom:12px;";

        var preRefCheckbox = document.createElement("input");
        preRefCheckbox.type = "checkbox";
        preRefCheckbox.id = "saBuilderPreReference";
        preRefCheckbox.style.cssText = "width:18px;height:18px;cursor:pointer;accent-color:#007bff;";

        var preRefLabel = document.createElement("label");
        preRefLabel.textContent = "Pre-Reference?";
        preRefLabel.setAttribute("for", "saBuilderPreReference");
        preRefLabel.style.cssText = "font-size:13px;font-weight:600;cursor:pointer;";

        preRefRow.appendChild(preRefCheckbox);
        preRefRow.appendChild(preRefLabel);
        timeInputsSection.appendChild(preRefRow);

        timeInputsSection.appendChild(createTimeInput("Days", "saBuilderDays", 200));
        timeInputsSection.appendChild(createTimeInput("Hours", "saBuilderHours", 23));
        timeInputsSection.appendChild(createTimeInput("Minutes", "saBuilderMinutes", 59));
        timeInputsSection.appendChild(createTimeInput("Seconds", "saBuilderSeconds", 59));

        timeColumn.appendChild(timeInputsSection);

        // Initial render
        renderSegments("");
        renderStudyEvents("");
        renderForms("");

        // Restore saved time offset values
        try {
            var savedTimeOffset = localStorage.getItem(STORAGE_SA_BUILDER_TIME_OFFSET);
            if (savedTimeOffset) {
                var timeData = JSON.parse(savedTimeOffset);
                if (timeData.days !== undefined) document.getElementById("saBuilderDays").value = timeData.days;
                if (timeData.hours !== undefined) document.getElementById("saBuilderHours").value = timeData.hours;
                if (timeData.minutes !== undefined) document.getElementById("saBuilderMinutes").value = timeData.minutes;
                if (timeData.seconds !== undefined) document.getElementById("saBuilderSeconds").value = timeData.seconds;
            }
        } catch (e) {}

        // Select All checkbox handler
        function updateSelectAllCheckbox() {
            var allChecked = true;
            for (var i = 0; i < segments.length; i++) {
                if (!segmentCheckboxStates[segments[i].value]) {
                    allChecked = false;
                    break;
                }
            }
            selectAllCheckbox.checked = allChecked;
        }

        selectAllCheckbox.addEventListener("change", function() {
            var isChecked = this.checked;
            for (var i = 0; i < segments.length; i++) {
                segmentCheckboxStates[segments[i].value] = isChecked;
            }
            try {
                localStorage.setItem(STORAGE_SA_BUILDER_SEGMENT_CHECKBOXES, JSON.stringify(segmentCheckboxStates));
            } catch (e) {}
            renderSegments(segmentSearch.value);
        });

        // Initialize Select All checkbox state
        updateSelectAllCheckbox();

        // Search handlers
        segmentSearch.addEventListener("input", function() { renderSegments(this.value); });
        eventSearch.addEventListener("input", function() { renderStudyEvents(this.value); });
        formSearch.addEventListener("input", function() { renderForms(this.value); });

        contentRow.appendChild(segmentColumnWrapper);
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

            var hiddenChecked = document.getElementById("saBuilderHidden").checked;
            var mandatoryChecked = document.getElementById("saBuilderMandatory").checked;
            var enforceChecked = document.getElementById("saBuilderEnforce").checked;
            var preWindowValue = document.getElementById("saBuilderPreWindow").value.trim();
            var postWindowValue = document.getElementById("saBuilderPostWindow").value.trim();
            var refActivityChecked = document.getElementById("saBuilderRefActivity").checked;
            var preReferenceChecked = document.getElementById("saBuilderPreReference").checked;

            // If Reference Activity is checked, void the time offset
            if (refActivityChecked) {
                timeOffset = { days: 0, hours: 0, minutes: 0, seconds: 0 };
            }

            var userSelection = {
                segments: selectedSegments,
                forms: selectedForms,
                timeOffset: timeOffset,
                hidden: hiddenChecked,
                mandatory: mandatoryChecked,
                enforce: enforceChecked,
                preWindow: preWindowValue,
                postWindow: postWindowValue,
                refActivity: refActivityChecked,
                preReference: preReferenceChecked
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
            if (SA_BUILDER_CANCELLED) {
                log("SA Builder: cancelled");
                return;
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
        stopBtn.textContent = "Close";
        stopBtn.style.cssText = "padding:10px 24px;border-radius:6px;border:none;background:#dc3545;color:#fff;font-weight:600;cursor:pointer;";
        stopBtn.addEventListener("click", function() {
            SA_BUILDER_CANCELLED = true;
            log("SA Builder: stopped by user");
            clearSABuilderStorage();
            if (SA_BUILDER_PROGRESS_POPUP_REF && SA_BUILDER_PROGRESS_POPUP_REF.close) {
                try {
                    SA_BUILDER_PROGRESS_POPUP_REF.close();
                } catch (e) {
                    log("SA Builder: error closing popup - " + String(e));
                }
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
        });
        if (SA_BUILDER_CANCELLED) {
            log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
            log("SA Builder: cancelled");
            return;
        }
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
                log("SA Builder: cancelled by user (X button)");
                clearSABuilderStorage();
                SA_BUILDER_PROGRESS_POPUP_REF = null;
            }
        });
        log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled");
            return;
        }
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
                // Check cancellation before starting
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before Add button");
                    break;
                }

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

                // Check cancellation after modal appears
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after modal appeared");
                    break;
                }

                await sleep(1000);

                // Check cancellation before selections
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before selections");
                    break;
                }
                log("SA_BUILDER_CANCELLED: " + SA_BUILDER_CANCELLED);
                // Select segment
                log("SA Builder: selecting segment " + item.segmentText + " (value: " + item.segmentValue + ")");
                await selectSelect2Value("segment", item.segmentValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after segment selection");
                    break;
                }

                // Select study event
                log("SA Builder: selecting study event " + item.eventText + " (value: " + item.eventValue + ")");
                await selectSelect2Value("studyEvent", item.eventValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after study event selection");
                    break;
                }

                // Select form
                log("SA Builder: selecting form " + item.formText + " (value: " + item.formValue + ")");
                await selectSelect2Value("form", item.formValue);
                await sleep(500);

                // Check cancellation
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after form selection");
                    break;
                }
                // Check Hidden checkbox if user selected it
                if (userSelection.hidden) {
                    var hiddenCheckboxEl = document.querySelector("#uniform-hidden span input#hidden.checkbox");
                    if (!hiddenCheckboxEl) {
                        hiddenCheckboxEl = document.getElementById("hidden");
                    }
                    if (hiddenCheckboxEl && !hiddenCheckboxEl.checked) {
                        hiddenCheckboxEl.click();
                        log("SA Builder: Hidden checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Mandatory checkbox (default is checked, uncheck if user unchecked it)
                if (!userSelection.mandatory) {
                    var mandatoryEl = document.querySelector("#uniform-mandatory span input#mandatory.checkbox");
                    if (!mandatoryEl) {
                        mandatoryEl = document.getElementById("mandatory");
                    }
                    if (mandatoryEl && mandatoryEl.checked) {
                        mandatoryEl.click();
                        log("SA Builder: Mandatory checkbox unchecked");
                        await sleep(200);
                    }
                }

                // Handle Enforce Data Collection Order checkbox (default is unchecked, check if user checked it)
                if (userSelection.enforce) {
                    var enforceEl = document.querySelector("#uniform-enforceDataCollectionOrder span input#enforceDataCollectionOrder.checkbox");
                    if (!enforceEl) {
                        enforceEl = document.getElementById("enforceDataCollectionOrder");
                    }
                    if (enforceEl && !enforceEl.checked) {
                        enforceEl.click();
                        log("SA Builder: Enforce Data Collection Order checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Pre-Window and Post-Window fields
                var preWindow = document.getElementById("preWindow");
                var postWindow = document.getElementById("postWindow");
                if (preWindow) {
                    preWindow.value = userSelection.preWindow || "";
                    preWindow.dispatchEvent(new Event("input", { bubbles: true }));
                    log("SA Builder: Pre-Window set to '" + (userSelection.preWindow || "") + "'");
                }
                if (postWindow) {
                    postWindow.value = userSelection.postWindow || "";
                    postWindow.dispatchEvent(new Event("input", { bubbles: true }));
                    log("SA Builder: Post-Window set to '" + (userSelection.postWindow || "") + "'");
                }

                // Handle Reference Activity checkbox
                if (userSelection.refActivity) {
                    var refActivityEl = document.getElementById("referenceActivity");
                    if (refActivityEl && !refActivityEl.checked) {
                        refActivityEl.click();
                        log("SA Builder: Reference Activity checkbox checked");
                        await sleep(200);
                    }
                }

                // Handle Pre-Reference checkbox (must sync with user selection since default inherits from last added item)
                var preRefContainer = document.getElementById("uniform-offset.preReference");
                var preRefEl = document.getElementById("offset.preReference");
                if (preRefEl && preRefContainer) {
                    var preRefSpan = preRefContainer.querySelector("span");
                    var preRefCurrentlyChecked = preRefSpan && preRefSpan.classList.contains("checked");
                    log("SA Builder: Pre-Reference current state: " + (preRefCurrentlyChecked ? "checked" : "unchecked") + ", user wants: " + (userSelection.preReference ? "checked" : "unchecked"));
                    if (userSelection.preReference !== preRefCurrentlyChecked) {
                        // Set the native checked property directly
                        preRefEl.checked = userSelection.preReference;
                        // Sync uniform.js span class
                        if (preRefSpan) {
                            if (userSelection.preReference) {
                                preRefSpan.classList.add("checked");
                            } else {
                                preRefSpan.classList.remove("checked");
                            }
                        }
                        // Try to sync via jQuery uniform if available
                        try {
                            if (window.jQuery && window.jQuery.fn.uniform) {
                                window.jQuery(preRefEl).uniform.update(preRefEl);
                            }
                        } catch (ue) {}
                        // Dispatch change event so the form picks up the new value
                        preRefEl.dispatchEvent(new Event("change", { bubbles: true }));
                        log("SA Builder: Pre-Reference checkbox " + (userSelection.preReference ? "checked" : "unchecked"));
                        await sleep(200);
                    }
                }

                // Set time offset values (skip if Reference Activity is checked)
                if (!userSelection.refActivity) {
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
                } else {
                    log("SA Builder: Skipping time offset values (Reference Activity checked)");
                }

                await sleep(300);

                // Check cancellation before saving
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled before Save button");
                    break;
                }

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

                // Check cancellation after save completes
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after save completed");
                    break;
                }

                // Update status
                mappedItems[idx].status = "Complete";
                progressContent.updateItem(idx, "Complete");
                log("SA Builder: item completed - " + item.key);

                // Wait for table refresh
                await sleep(2000);

                // Check cancellation after table refresh wait
                if (SA_BUILDER_CANCELLED) {
                    log("SA Builder: cancelled after table refresh wait");
                    break;
                }

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

        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled");
            return;
        }
        // Clear running state
        try {
            localStorage.removeItem(STORAGE_SA_BUILDER_RUNNING);
            localStorage.removeItem(STORAGE_SA_BUILDER_CURRENT_INDEX);
        } catch (e) {}
    }

    // Main entry point for SA Builder
    async function runSABuilder() {
        // Check if on correct page
        if (!isOnSABuilderPage()) {
            var errorPopup = createPopup({
                title: "Scheduled Activities Builder",
                content: '<div style="text-align:center;padding:20px;"><p style="color:#f66;font-size:16px;margin-bottom:16px;">⚠️ Wrong Page</p><p>You must be on the Activity Plans Show page to use this feature.</p><p style="margin-top:12px;font-size:12px;color:#ffffffff;">Required URL: ' + SA_BUILDER_TARGET_URL + '</p></div>',
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
                // Only clear storage if user explicitly cancelled, not on auto-close
                if (SA_BUILDER_CANCELLED) {
                    clearSABuilderStorage();
                    log("SA Builder: cancelled after table refresh wait");
                }
                SA_BUILDER_POPUP_REF = null;
            }
        });
        if (SA_BUILDER_CANCELLED) {
            log("SA Builder: cancelled after table refresh wait");
            return;
        }
        log("SA Builder: selection GUI displayed");
    }

    // =========================
    // UI SCALING FEATURE
    // =========================


    function scale(value) {
        if (typeof value === 'string' && value.endsWith('px')) {
            return (parseFloat(value) * UI_SCALE) + 'px';
        }
        return (value * UI_SCALE) + 'px';
    }

    function getStoredUIScale() {
        try {
            var stored = localStorage.getItem(STORAGE_UI_SCALE);
            if (stored) {
                var scale = parseFloat(stored);
                if (!isNaN(scale) && scale >= 0.5 && scale <= 1.0) {
                    return scale;
                }
            }
        } catch (e) {}
        return 1.0;
    }

    /**
     * @param {number} scale
     */
    function setStoredUIScale(scale) {
        try {
            localStorage.setItem(STORAGE_UI_SCALE, String(scale));
        } catch (e) {}
    }

    /**
     * @param {number} newScale
     */
    function updateUIScale(newScale) {
        UI_SCALE = Math.max(0.5, Math.min(1.0, newScale));
        setStoredUIScale(UI_SCALE);
    }

    // Initialize UI scale from storage
    UI_SCALE = getStoredUIScale();

    
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

    // Persist an after-refresh action string.
    function setAfterRefresh(action) {
        try {
            localStorage.setItem(STORAGE_AFTER_REFRESH, String(action));
            log("AfterRefresh=" + String(action));
        } catch (e) {}
    }


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
        // @ts-ignore
        var api = typeof GM !== "undefined" && typeof GM.openInTab === "function";
        if (api) {
            // @ts-ignore
            return GM.openInTab(url, { active: !!active, insert: true, setParent: true });
        }
        // @ts-ignore
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
            w = scale(PANEL_DEFAULT_WIDTH);
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
            panel.style.width = scale(PANEL_DEFAULT_WIDTH);
            panel.style.height = scale(PANEL_HEADER_HEIGHT_PX);
            panel.style.overflow = "hidden";

            if (bodyContainer) {
                bodyContainer.style.display = "none";
            }
            if (resizeHandle) {
                resizeHandle.style.display = "none";
            }
            if (collapseBtn) {
                collapseBtn.textContent = "+";
            }
        } else {
            // Expand: restore last stored size
            var size = getStoredPanelSize();
            panel.style.width = size.width;
            panel.style.height = size.height;
            panel.style.overflow = "hidden";

            if (bodyContainer) {
                bodyContainer.style.display = "block";
            }
            if (resizeHandle) {
                resizeHandle.style.display = "block";
            }
            if (collapseBtn) {
                collapseBtn.textContent = "—";
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

    // Detect if current path is an eligibility form page
    function isEligibilityFormPage() {
        var path = location.pathname;
        var ok = path.indexOf("/secure/crfdesign/studylibrary/show/form/") !== -1;
        return ok;
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
        var maxTop = vh - ph / 2;
        var minRight = -pw / 2;
        var maxRight = vw - pw / 2;
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
        btn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        btn.style.padding = scale(BUTTON_PADDING_PX);
        btn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
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


    // Glassmorphism Theme Style Injection (from Test Automator 2)
    function injectThemeStylesIfNeeded() {
        if (!isGlassTheme()) return;
        var existing = document.getElementById(THEME_STYLE_TAG_ID);
        if (existing) return;
        var style = document.createElement("style");
        style.id = THEME_STYLE_TAG_ID;
        style.textContent = [
            "." + THEME_SCOPE_CLASS + " {",
            "  font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;",
            "  color: " + THEME_TEXT_PRIMARY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-theme-backdrop {",
            "  position: fixed; top:0; left:0; width:100%; height:100%;",
            "  background: linear-gradient(135deg, " + THEME_GRADIENT_START + " 0%, " + THEME_GRADIENT_END + " 100%);",
            "  z-index: " + (THEME_Z_BASE - 1) + ";",
            "  pointer-events: none;",
            "}",
            "." + THEME_SCOPE_CLASS + ".ie-glass-panel,",
            "." + THEME_SCOPE_CLASS + " .ie-glass-panel {",
            "  background: " + THEME_GRADIENT_BG + ";",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  box-shadow: inset 0 0 0 1px " + THEME_SURFACE_INNER_BORDER + ", " + THEME_SHADOW + ";",
            "  border-radius: " + THEME_RADIUS + "px;",
            "  color: " + THEME_TEXT_PRIMARY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + ".ie-glass-panel-header,",
            "." + THEME_SCOPE_CLASS + " .ie-glass-panel-header {",
            "  background: linear-gradient(135deg, " + THEME_GRADIENT_START + " 0%, " + THEME_GRADIENT_END + " 100%);",
            "  border-bottom: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  color: " + THEME_TEXT_PRIMARY + ";",
            "  cursor: move; user-select: none;",
            "}",
            "." + THEME_SCOPE_CLASS + ".ie-glass-panel-header-danger,",
            "." + THEME_SCOPE_CLASS + " .ie-glass-panel-header-danger {",
            "  background: linear-gradient(135deg, " + THEME_DANGER + " 0%, " + THEME_DANGER_DARK + " 100%);",
            "  border-bottom: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  color: " + THEME_TEXT_PRIMARY + ";",
            "  cursor: move; user-select: none;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-primary {",
            "  background: linear-gradient(135deg, " + THEME_GRADIENT_START + " 0%, " + THEME_GRADIENT_END + " 100%);",
            "  color: " + THEME_TEXT_PRIMARY + "; border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 8px 16px; cursor: pointer; font-weight: 600;",
            "  font-size: 14px; transition: opacity 0.2s, box-shadow 0.2s;",
            "  box-shadow: " + THEME_SHADOW + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-primary:hover {",
            "  opacity: 0.9; box-shadow: 0 0 0 2px " + THEME_ACCENT_HOVER + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-primary:focus-visible {",
            "  box-shadow: " + THEME_OUTLINE_FOCUS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-primary:disabled {",
            "  opacity: " + THEME_DISABLED_OPACITY + "; cursor: default;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-secondary {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  color: " + THEME_TEXT_PRIMARY + "; border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 8px 16px; cursor: pointer; font-weight: 500;",
            "  font-size: 13px; transition: background-color 0.2s, box-shadow 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-secondary:hover {",
            "  background-color: " + THEME_SURFACE_BG_HEAVY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-secondary:disabled {",
            "  opacity: " + THEME_DISABLED_OPACITY + "; cursor: default;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-secondary:focus-visible {",
            "  box-shadow: " + THEME_OUTLINE_FOCUS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-danger {",
            "  background: linear-gradient(135deg, " + THEME_DANGER + " 0%, " + THEME_DANGER_DARK + " 100%);",
            "  color: " + THEME_TEXT_PRIMARY + "; border: 1px solid rgba(239,68,68,0.4);",
            "  border-radius: 8px; padding: 8px 16px; cursor: pointer; font-weight: 600;",
            "  font-size: 14px; transition: opacity 0.2s, box-shadow 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-danger:hover {",
            "  opacity: 0.9;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-danger:focus-visible {",
            "  box-shadow: " + THEME_OUTLINE_FOCUS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-success {",
            "  background: linear-gradient(135deg, " + THEME_SUCCESS + " 0%, " + THEME_SUCCESS_DARK + " 100%);",
            "  color: " + THEME_TEXT_PRIMARY + "; border: 1px solid rgba(16,185,129,0.4);",
            "  border-radius: 8px; padding: 8px 16px; cursor: pointer; font-weight: 600;",
            "  font-size: 14px; transition: opacity 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-success:hover { opacity: 0.9; }",
            "." + THEME_SCOPE_CLASS + " .ie-btn-success:disabled {",
            "  opacity: " + THEME_DISABLED_OPACITY + "; cursor: default;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-warning {",
            "  background: linear-gradient(135deg, " + THEME_WARNING + " 0%, " + THEME_WARNING_DARK + " 100%);",
            "  color: " + THEME_TEXT_INVERSE + "; border: 1px solid rgba(245,158,11,0.4);",
            "  border-radius: 8px; padding: 8px 16px; cursor: pointer; font-weight: 600;",
            "  font-size: 14px; transition: opacity 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-btn-warning:hover { opacity: 0.9; }",
            "." + THEME_SCOPE_CLASS + " .ie-btn-warning:disabled {",
            "  opacity: " + THEME_DISABLED_OPACITY + "; cursor: default;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-input {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 10px 12px; color: " + THEME_TEXT_PRIMARY + ";",
            "  font-size: 14px; outline: none; transition: border-color 0.2s, box-shadow 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-input:focus {",
            "  border-color: " + THEME_ACCENT + "; box-shadow: " + THEME_OUTLINE_FOCUS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-input::placeholder { color: " + THEME_TEXT_MUTED + "; }",
            "." + THEME_SCOPE_CLASS + " .ie-select {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 8px 12px; color: " + THEME_TEXT_PRIMARY + ";",
            "  font-size: 14px; cursor: pointer; outline: none;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-select:focus {",
            "  border-color: " + THEME_ACCENT + "; box-shadow: " + THEME_OUTLINE_FOCUS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-select option {",
            "  background: " + THEME_SELECT_OPTION_BG + "; color: " + THEME_TEXT_PRIMARY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-checkbox-row {",
            "  display: flex; align-items: center; gap: 10px;",
            "  padding: 8px 12px; background-color: " + THEME_SURFACE_BG + ";",
            "  border-radius: 8px; cursor: pointer; transition: background-color 0.15s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-checkbox-row:hover {",
            "  background-color: " + THEME_SURFACE_BG_HEAVY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " input[type='checkbox'] {",
            "  accent-color: " + THEME_ACCENT + "; cursor: pointer;",
            "}",
            "." + THEME_SCOPE_CLASS + " input[type='range'] {",
            "  accent-color: " + THEME_ACCENT + "; cursor: pointer;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-success {",
            "  background: rgba(16,185,129,0.25); color: " + THEME_SUCCESS + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 600;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-danger {",
            "  background: rgba(239,68,68,0.25); color: " + THEME_DANGER + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 600;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-warning {",
            "  background: rgba(245,158,11,0.25); color: " + THEME_WARNING + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 600;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-muted {",
            "  background: " + THEME_SURFACE_BG + "; color: " + THEME_TEXT_MUTED + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-accent {",
            "  background: rgba(167,139,250,0.25); color: " + THEME_ACCENT + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 600;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-pill-info {",
            "  background: rgba(102,126,234,0.25); color: " + THEME_GRADIENT_START + ";",
            "  padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 600;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-progress-track {",
            "  width: 100%; height: 8px; background: " + THEME_SURFACE_BG + ";",
            "  border-radius: 4px; overflow: hidden;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-progress-fill {",
            "  height: 100%; transition: width 0.3s;",
            "  background: linear-gradient(135deg, " + THEME_GRADIENT_START + " 0%, " + THEME_GRADIENT_END + " 100%);",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-spinner {",
            "  width: 40px; height: 40px;",
            "  border: 4px solid " + THEME_SURFACE_BORDER + ";",
            "  border-top: 4px solid " + THEME_ACCENT + ";",
            "  border-radius: 50%; margin: 0 auto 16px;",
            "  animation: ieThemeSpin 1s linear infinite;",
            "}",
            "@keyframes ieThemeSpin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }",
            "@media (prefers-reduced-motion: reduce) {",
            "  ." + THEME_SCOPE_CLASS + " .ie-spinner { animation: none; }",
            "  ." + THEME_SCOPE_CLASS + " * { transition-duration: 0s !important; animation-duration: 0s !important; }",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-list-item {",
            "  display: flex; align-items: center; gap: 10px;",
            "  padding: 10px; margin-bottom: 6px;",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; background-color: " + THEME_SURFACE_BG + ";",
            "  cursor: pointer; transition: background-color 0.2s, border-color 0.2s;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-list-item:hover {",
            "  background-color: " + THEME_SURFACE_BG_HEAVY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-list-item.ie-selected {",
            "  background-color: rgba(167,139,250,0.2); border-color: " + THEME_ACCENT + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-list-item.ie-selected-success {",
            "  background-color: rgba(16,185,129,0.15); border-color: " + THEME_SUCCESS + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-list-item.ie-selected-danger {",
            "  background-color: rgba(239,68,68,0.15); border-color: " + THEME_DANGER + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-section-header {",
            "  font-weight: 600; font-size: 14px; color: " + THEME_ACCENT + ";",
            "  padding: 8px; background-color: " + THEME_SURFACE_BG + ";",
            "  border-radius: " + THEME_RADIUS + "px " + THEME_RADIUS + "px 0 0;",
            "  border: 1px solid " + THEME_SURFACE_BORDER + "; border-bottom: none;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-section-body {",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 0 0 " + THEME_RADIUS + "px " + THEME_RADIUS + "px;",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-count-badge {",
            "  font-size: 11px; color: " + THEME_TEXT_MUTED + ";",
            "  background: " + THEME_SURFACE_BG + "; padding: 2px 8px;",
            "  border-radius: 10px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-radio-indicator {",
            "  width: 18px; height: 18px; border: 2px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 50%; flex-shrink: 0; display: flex;",
            "  align-items: center; justify-content: center;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-sub-surface {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-status-area {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 6px; font-size: 13px;",
            "  white-space: pre-wrap; color: " + THEME_TEXT_PRIMARY + ";",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-log-box {",
            "  background-color: rgba(0,0,0,0.25);",
            "  border: 1px solid " + THEME_SURFACE_BORDER + ";",
            "  border-radius: 8px; padding: 6px; font-size: 12px;",
            "  color: " + THEME_TEXT_MUTED + "; white-space: pre-wrap;",
            "  word-break: break-word; overflow-wrap: anywhere;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-resize-handle {",
            "  position: absolute; width: 12px; height: 12px;",
            "  right: 6px; bottom: 6px; cursor: se-resize;",
            "  background: " + THEME_SURFACE_BORDER + "; border-radius: 2px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-error-text {",
            "  color: " + THEME_DANGER + "; text-align: center; font-size: 14px; min-height: 20px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-muted-text {",
            "  color: " + THEME_TEXT_MUTED + "; font-size: 13px;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-selection-info {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border-radius: 8px; padding: 10px; font-size: 13px;",
            "  color: " + THEME_TEXT_MUTED + "; text-align: center;",
            "}",
            "." + THEME_SCOPE_CLASS + " .ie-summary-box {",
            "  background-color: " + THEME_SURFACE_BG + ";",
            "  border-radius: 8px; padding: 12px; text-align: center;",
            "}",
            "." + THEME_SCOPE_CLASS + " ::-webkit-scrollbar { width: 8px; height: 8px; }",
            "." + THEME_SCOPE_CLASS + " ::-webkit-scrollbar-track {",
            "  background: " + THEME_SCROLLBAR_TRACK + "; border-radius: 4px;",
            "}",
            "." + THEME_SCOPE_CLASS + " ::-webkit-scrollbar-thumb {",
            "  background: " + THEME_SCROLLBAR_THUMB + "; border-radius: 4px;",
            "}",
            "." + THEME_SCOPE_CLASS + " ::-webkit-scrollbar-thumb:hover {",
            "  background: rgba(255,255,255,0.5);",
            "}",
            ""
        ].join("\n");
        document.head.appendChild(style);
    }

    function removeThemeStyles() {
        var existing = document.getElementById(THEME_STYLE_TAG_ID);
        if (existing) existing.remove();
    }

    function applyThemeToUiRoots() {
        if (!isGlassTheme()) return;
        var panel = document.getElementById(PANEL_ID);
        if (panel && !panel.classList.contains(THEME_SCOPE_CLASS)) {
            panel.classList.add(THEME_SCOPE_CLASS);
        }
        var popups = document.querySelectorAll("[id^='clinsparkPopup_']");
        for (var i = 0; i < popups.length; i++) {
            if (!popups[i].classList.contains(THEME_SCOPE_CLASS)) {
                popups[i].classList.add(THEME_SCOPE_CLASS);
            }
        }
    }

    // Create a draggable, closeable popup styled like the panel

    // Create a draggable, closeable popup styled like the panel
    function createPopup(options) {
        var glass = isGlassTheme();
        if (glass) injectThemeStylesIfNeeded();
        options = options || {};
        var title = options.title || "Popup";
        var description = options.description || "";
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
        if (glass) {
            popup.classList.add(THEME_SCOPE_CLASS);
            popup.classList.add("ie-glass-panel");
        }
        popup.style.position = "fixed";
        popup.style.top = "50%";
        popup.style.left = "50%";
        popup.style.transform = "translate(-50%, -50%)";
        popup.style.zIndex = glass ? String(THEME_Z_OVERLAY) : "999998";
        if (!glass) {
            popup.style.background = "#111";
            popup.style.color = "#fff";
            popup.style.border = "1px solid #444";
            popup.style.borderRadius = "8";
            popup.style.padding = "10px";
            popup.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.5)";
        } else {
            popup.style.padding = "0";
            popup.style.overflow = "hidden";
        }
        popup.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        popup.style.fontSize = "14px";
        popup.style.width = width;
        popup.style.maxWidth = maxWidth;
        popup.style.height = height;
        popup.style.maxHeight = maxHeight;
        popup.style.boxSizing = "border-box";
        popup.style.display = "flex";
        popup.style.flexDirection = "column";

        var headerBar = document.createElement("div");
        if (glass) {
            headerBar.classList.add("ie-glass-panel-header");
        }
        headerBar.style.position = "relative";
        headerBar.style.display = "grid";
        headerBar.style.gridTemplateColumns = "1fr auto";
        headerBar.style.alignItems = "center";
        headerBar.style.gap = String(PANEL_HEADER_GAP_PX) + "px";
        headerBar.style.height = String(PANEL_HEADER_HEIGHT_PX) + "px";
        headerBar.style.boxSizing = "border-box";
        headerBar.style.padding = "0 12px";
        if (!glass) {
            headerBar.style.borderBottom = "1px solid #444";
        } else {
            headerBar.style.borderRadius = THEME_RADIUS + "px " + THEME_RADIUS + "px 0 0";
        }
        headerBar.style.cursor = "move";
        headerBar.style.userSelect = "none";

        var titleContainer = document.createElement("div");
        titleContainer.style.display = "flex";
        titleContainer.style.flexDirection = "column";
        titleContainer.style.justifyContent = "center";

        var titleEl = document.createElement("div");
        titleEl.textContent = title;
        titleEl.style.fontWeight = "600";
        titleEl.style.textAlign = "left";
        if (glass) titleEl.style.color = THEME_TEXT_PRIMARY;
        titleContainer.appendChild(titleEl);

        if (description) {
            var descEl = document.createElement("div");
            descEl.textContent = description;
            descEl.style.fontSize = "12px";
            descEl.style.color = glass ? THEME_TEXT_MUTED : "#aaa";
            descEl.style.textAlign = "left";
            descEl.style.marginTop = "2px";
            titleContainer.appendChild(descEl);
        }

        headerBar.appendChild(titleContainer);

        var closeBtn = document.createElement("button");
        closeBtn.textContent = "✕";
        closeBtn.style.background = "transparent";
        closeBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
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
            closeBtn.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#333";
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
        handle.style.background = isGlassTheme() ? THEME_SURFACE_BG_HEAVY : "#333";
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

            var minW = toInt(scale(PANEL_DEFAULT_WIDTH));
            if (newW < minW) {
                newW = minW;
            }
            if (newW > PANEL_MAX_WIDTH_PX) {
                newW = PANEL_MAX_WIDTH_PX;
            }

            var minH = PANEL_HEADER_HEIGHT_PX + 88;
            if (newH < minH) {
                newH = minH;
            }

            panel.style.width = String(newW) + "px";
            panel.style.height = String(newH) + "px";

            if (bodyContainer) {
                bodyContainer.style.display = "block";

                bodyContainer.style.height = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
                bodyContainer.style.maxHeight = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";

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
        var savedHotkey = getPanelHotkey();
        var match = false;

        if (n.key && n.key.toUpperCase() === savedHotkey.toUpperCase()) {
            match = true;
        } else if (n.code && n.code.toUpperCase() === savedHotkey.toUpperCase()) {
            match = true;
        } else if (typeof n.keyCode === "number") {
            // Fallback for F-keys using keyCode
            if (savedHotkey.toUpperCase() === "F2" && n.keyCode === 113) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F1" && n.keyCode === 112) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F3" && n.keyCode === 114) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F4" && n.keyCode === 115) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F5" && n.keyCode === 116) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F6" && n.keyCode === 117) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F7" && n.keyCode === 118) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F8" && n.keyCode === 119) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F9" && n.keyCode === 120) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F10" && n.keyCode === 121) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F11" && n.keyCode === 122) {
                match = true;
            } else if (savedHotkey.toUpperCase() === "F12" && n.keyCode === 123) {
                match = true;
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
        log("Hotkey: bound for " + String(getPanelHotkey()));
    }

    
    function setPanelHidden(flag) {
        try {
            localStorage.setItem(STORAGE_PANEL_HIDDEN, flag ? "1" : "0");
            log("Panel hidden state set to " + String(flag));
        } catch (e) {
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
    //=========================
    // SETTING FEATURE
    //=========================
    // This section contains functions for setting up the extension.
    // It includes functions for loading and saving settings, as well as
    // functions for updating the UI based on the current settings.
    //=========================
    var STORAGE_BUTTON_VISIBILITY = "activityPlanState.buttonVisibility";
    var SETTINGS_POPUP_REF = null;
    function getButtonVisibility() {
        try {
            var raw = localStorage.getItem(STORAGE_BUTTON_VISIBILITY);
            if (raw) {
                return JSON.parse(raw);
            }
        } catch (e) {}
        return null;
    }

    function setButtonVisibility(visibilityMap) {
        try {
            localStorage.setItem(STORAGE_BUTTON_VISIBILITY, JSON.stringify(visibilityMap));
        } catch (e) {}
    }

    function getPanelHotkey() {
        try {
            var saved = localStorage.getItem(STORAGE_PANEL_HOTKEY);
            if (saved) {
                return saved;
            }
        } catch (err) {
            log("Error reading hotkey from localStorage: " + String(err));
        }
        return "F2";
    }

    function setPanelHotkey(hotkey) {
        try {
            localStorage.setItem(STORAGE_PANEL_HOTKEY, hotkey);
            return true;
        } catch (err) {
            log("Error saving hotkey to localStorage: " + String(err));
            return false;
        }
    }

    function isButtonVisible(label) {
        var visibility = getButtonVisibility();
        if (!visibility) {
            return true;
        }
        if (visibility.hasOwnProperty(label)) {
            return visibility[label];
        }
        return true;
    }

    function openSettingsPopup() {
        var glass = isGlassTheme();
        var buttonLabels = [
            "Lock Activity Plans",
            "Lock Sample Paths",
            "Update Study Status",
            "Add Cohort Subjects",
            "Run ICF Consent",
            "Run Button (1-5)",
            "Import Cohort Subjects",
            "Pull Barcode",
            "Pull Lab Barcode",
            "Add Existing Subject",
            "Search Methods",
            "Scheduled Activities Builder",
            "Run Form (OOR) Below Range",
            "Run Form (OOR) Above Range",
            "Run Form (IR) In Range",
            "Collect All",
            "Import I/E",
            "Clear Mapping",
            "Archive/Update Forms",
            "Copy Activity Forms",
            "Item Method Forms",
            "Find Form",
            "Find Study Events",
            "Pause",
            "Clear Logs",
            "Hide Logs"
        ];

        var currentVisibility = getButtonVisibility() || {};

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "8px";
        container.style.minWidth = "280px";

        var description = document.createElement("div");
        description.textContent = "Select which buttons to display in the panel:";
        description.style.fontSize = "13px";
        description.style.color = glass ? THEME_TEXT_MUTED : "#aaa";
        description.style.marginBottom = "12px";
        container.appendChild(description);

        var checkboxContainer = document.createElement("div");
        checkboxContainer.style.display = "flex";
        checkboxContainer.style.flexDirection = "column";
        checkboxContainer.style.gap = "6px";
        checkboxContainer.style.maxHeight = "320px";
        checkboxContainer.style.overflowY = "auto";
        checkboxContainer.style.paddingRight = "8px";

        var checkboxes = [];

        for (var i = 0; i < buttonLabels.length; i++) {
            var label = buttonLabels[i];
            var isChecked = currentVisibility.hasOwnProperty(label) ? currentVisibility[label] : true;

            var row = document.createElement("label");
            row.style.display = "flex";
            row.style.alignItems = "center";
            row.style.gap = "10px";
            row.style.padding = "8px 12px";
            row.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
            row.style.borderRadius = "6px";
            row.style.cursor = "pointer";
            row.style.transition = "background 0.15s";
            row.onmouseenter = (function(r) { return function() { r.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#252525"; }; })(row);
            row.onmouseleave = (function(r) { return function() { r.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a"; }; })(row);

            var checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = isChecked;
            checkbox.style.width = "18px";
            checkbox.style.height = "18px";
            checkbox.style.cursor = "pointer";
            checkbox.style.accentColor = glass ? THEME_ACCENT : "#5b43c7";
            checkbox.dataset.label = label;

            var labelText = document.createElement("span");
            labelText.textContent = label;
            labelText.style.fontSize = "14px";
            labelText.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";

            row.appendChild(checkbox);
            row.appendChild(labelText);
            checkboxContainer.appendChild(row);
            checkboxes.push(checkbox);
        }

        container.appendChild(checkboxContainer);

        // Theme Toggle Section
        var themeSection = document.createElement("div");
        themeSection.style.marginTop = "20px";
        themeSection.style.paddingTop = "20px";
        themeSection.style.borderTop = glass ? ("1px solid " + THEME_SURFACE_BORDER) : "1px solid #333";

        var themeLabel = document.createElement("div");
        themeLabel.textContent = "UI Theme:";
        themeLabel.style.fontSize = "13px";
        themeLabel.style.color = glass ? THEME_TEXT_MUTED : "#aaa";
        themeLabel.style.marginBottom = "8px";
        themeSection.appendChild(themeLabel);

        var themeRow = document.createElement("div");
        themeRow.style.display = "flex";
        themeRow.style.gap = "10px";

        var themes = [
            { value: THEME_MODE_BLACK, label: "Black" },
            { value: THEME_MODE_GLASS, label: "Glassmorphism" }
        ];
        for (var ti = 0; ti < themes.length; ti++) {
            var themeOpt = themes[ti];
            var themeBtn = document.createElement("button");
            themeBtn.textContent = themeOpt.label;
            themeBtn.dataset.themeValue = themeOpt.value;
            var isActive = (glass && themeOpt.value === THEME_MODE_GLASS) || (!glass && themeOpt.value === THEME_MODE_BLACK);
            if (isActive) {
                themeBtn.style.background = glass ? THEME_ACCENT : "#5b43c7";
                themeBtn.style.color = glass ? THEME_TEXT_INVERSE : "#fff";
                themeBtn.style.fontWeight = "600";
            } else {
                themeBtn.style.background = glass ? THEME_SURFACE_BG : "#333";
                themeBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
                themeBtn.style.fontWeight = "400";
            }
            themeBtn.style.border = "none";
            themeBtn.style.borderRadius = "6px";
            themeBtn.style.padding = "8px 16px";
            themeBtn.style.cursor = "pointer";
            themeBtn.style.fontSize = "13px";
            themeBtn.style.flex = "1";
            themeBtn.addEventListener("click", (function(val, allBtns) {
                return function() {
                    try { localStorage.setItem(STORAGE_THEME_MODE, val); } catch(e) {}
                    // Update button appearances
                    var siblings = themeRow.querySelectorAll("button");
                    for (var si = 0; si < siblings.length; si++) {
                        if (siblings[si].dataset.themeValue === val) {
                            siblings[si].style.background = val === THEME_MODE_GLASS ? THEME_ACCENT : "#5b43c7";
                            siblings[si].style.color = val === THEME_MODE_GLASS ? THEME_TEXT_INVERSE : "#fff";
                            siblings[si].style.fontWeight = "600";
                        } else {
                            siblings[si].style.background = val === THEME_MODE_GLASS ? THEME_SURFACE_BG : "#333";
                            siblings[si].style.color = val === THEME_MODE_GLASS ? THEME_TEXT_PRIMARY : "#fff";
                            siblings[si].style.fontWeight = "400";
                        }
                    }
                };
            })(themeOpt.value));
            themeRow.appendChild(themeBtn);
        }
        themeSection.appendChild(themeRow);

        var themeHint = document.createElement("div");
        themeHint.textContent = "Theme change will apply after Save & Refresh";
        themeHint.style.fontSize = "11px";
        themeHint.style.color = glass ? THEME_TEXT_MUTED : "#666";
        themeHint.style.marginTop = "6px";
        themeHint.style.fontStyle = "italic";
        themeSection.appendChild(themeHint);

        container.appendChild(themeSection);

        // Hotkey Configuration Section
        var hotkeySection = document.createElement("div");
        hotkeySection.style.marginTop = "20px";
        hotkeySection.style.paddingTop = "20px";
        hotkeySection.style.borderTop = glass ? ("1px solid " + THEME_SURFACE_BORDER) : "1px solid #333";

        var hotkeyLabel = document.createElement("div");
        hotkeyLabel.textContent = "Panel Toggle Hotkey:";
        hotkeyLabel.style.fontSize = "13px";
        hotkeyLabel.style.color = glass ? THEME_TEXT_MUTED : "#aaa";
        hotkeyLabel.style.marginBottom = "8px";
        hotkeySection.appendChild(hotkeyLabel);

        var hotkeyInputRow = document.createElement("div");
        hotkeyInputRow.style.display = "flex";
        hotkeyInputRow.style.gap = "10px";
        hotkeyInputRow.style.alignItems = "center";

        var hotkeyInput = document.createElement("input");
        hotkeyInput.type = "text";
        hotkeyInput.value = getPanelHotkey();
        hotkeyInput.placeholder = "Press a key...";
        hotkeyInput.style.flex = "1";
        hotkeyInput.style.padding = "10px 12px";
        hotkeyInput.style.background = glass ? THEME_SURFACE_BG : "#2a2a2a";
        hotkeyInput.style.border = glass ? ("1px solid " + THEME_SURFACE_BORDER) : "1px solid #444";
        hotkeyInput.style.borderRadius = "6px";
        hotkeyInput.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        hotkeyInput.style.fontSize = "14px";
        hotkeyInput.style.outline = "none";
        hotkeyInput.style.fontFamily = "monospace";
        hotkeyInput.style.cursor = "pointer";

        hotkeyInput.addEventListener("focus", function() {
            this.style.borderColor = glass ? THEME_ACCENT : "#5b43c7";
        });

        hotkeyInput.addEventListener("blur", function() {
            this.style.borderColor = glass ? THEME_SURFACE_BORDER : "#333";
        });

        hotkeyInput.addEventListener("keydown", function(e) {
            e.preventDefault();
            e.stopPropagation();

            var key = e.key;
            var code = e.code;
            var displayKey = key;

            // Handle special keys
            if (key.length === 1 && key.match(/[a-z]/i)) {
                displayKey = key.toUpperCase();
            } else if (code && code.startsWith("Key")) {
                displayKey = code.substring(3);
            } else if (code && code.startsWith("Digit")) {
                displayKey = code.substring(5);
            } else if (key === " ") {
                displayKey = "Space";
            } else if (key.startsWith("F") && key.length <= 3) {
                displayKey = key.toUpperCase();
            } else if (key === "Escape") {
                displayKey = "Escape";
            } else if (key === "Enter") {
                displayKey = "Enter";
            } else if (key === "Tab") {
                displayKey = "Tab";
            } else if (key === "Backspace") {
                displayKey = "Backspace";
            } else if (code) {
                displayKey = code;
            }

            this.value = displayKey;
        });

        var hotkeyResetBtn = document.createElement("button");
        hotkeyResetBtn.textContent = "Reset";
        hotkeyResetBtn.style.background = glass ? THEME_SURFACE_BG : "#333";
        hotkeyResetBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        hotkeyResetBtn.style.border = "none";
        hotkeyResetBtn.style.borderRadius = "6px";
        hotkeyResetBtn.style.padding = "10px 16px";
        hotkeyResetBtn.style.cursor = "pointer";
        hotkeyResetBtn.style.fontSize = "13px";
        if (glass) hotkeyResetBtn.className = "ie-btn-secondary";
        hotkeyResetBtn.onmouseenter = function() { this.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#444"; };
        hotkeyResetBtn.onmouseleave = function() { this.style.background = glass ? THEME_SURFACE_BG : "#333"; };
        hotkeyResetBtn.addEventListener("click", function() {
            hotkeyInput.value = "F2";
        });

        hotkeyInputRow.appendChild(hotkeyInput);
        hotkeyInputRow.appendChild(hotkeyResetBtn);
        hotkeySection.appendChild(hotkeyInputRow);

        var hotkeyHint = document.createElement("div");
        hotkeyHint.textContent = "Click the input field and press any key to set a new hotkey";
        hotkeyHint.style.fontSize = "11px";
        hotkeyHint.style.color = glass ? THEME_TEXT_MUTED : "#666";
        hotkeyHint.style.marginTop = "6px";
        hotkeyHint.style.fontStyle = "italic";
        hotkeySection.appendChild(hotkeyHint);

        container.appendChild(hotkeySection);

        var buttonRow = document.createElement("div");

        buttonRow.style.display = "flex";
        buttonRow.style.gap = "10px";
        buttonRow.style.marginTop = "16px";
        buttonRow.style.justifyContent = "flex-end";

        var selectAllBtn = document.createElement("button");
        selectAllBtn.textContent = "Select All";
        selectAllBtn.style.background = glass ? THEME_SURFACE_BG : "#333";
        selectAllBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        selectAllBtn.style.border = "none";
        selectAllBtn.style.borderRadius = "6px";
        selectAllBtn.style.padding = "8px 16px";
        selectAllBtn.style.cursor = "pointer";
        selectAllBtn.style.fontSize = "13px";
        if (glass) selectAllBtn.className = "ie-btn-secondary";
        selectAllBtn.onmouseenter = function() { this.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#444"; };
        selectAllBtn.onmouseleave = function() { this.style.background = glass ? THEME_SURFACE_BG : "#333"; };
        selectAllBtn.addEventListener("click", function() {
            for (var j = 0; j < checkboxes.length; j++) {
                checkboxes[j].checked = true;
            }
        });

        var deselectAllBtn = document.createElement("button");
        deselectAllBtn.textContent = "Deselect All";
        deselectAllBtn.style.background = glass ? THEME_SURFACE_BG : "#333";
        deselectAllBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        deselectAllBtn.style.border = "none";
        deselectAllBtn.style.borderRadius = "6px";
        deselectAllBtn.style.padding = "8px 16px";
        deselectAllBtn.style.cursor = "pointer";
        deselectAllBtn.style.fontSize = "13px";
        if (glass) deselectAllBtn.className = "ie-btn-secondary";
        deselectAllBtn.onmouseenter = function() { this.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#444"; };
        deselectAllBtn.onmouseleave = function() { this.style.background = glass ? THEME_SURFACE_BG : "#333"; };
        deselectAllBtn.addEventListener("click", function() {
            for (var j = 0; j < checkboxes.length; j++) {
                checkboxes[j].checked = false;
            }
        });

        var saveBtn = document.createElement("button");
        saveBtn.textContent = "Save & Refresh";
        if (glass) saveBtn.className = "ie-btn-primary";
        saveBtn.style.background = glass ? "" : "#5b43c7";
        saveBtn.style.color = glass ? "" : "#fff";
        saveBtn.style.border = "none";
        saveBtn.style.borderRadius = "6px";
        saveBtn.style.padding = "10px 20px";
        saveBtn.style.cursor = "pointer";
        saveBtn.style.fontSize = "14px";
        saveBtn.style.fontWeight = "600";
        if (!glass) {
            saveBtn.onmouseenter = function() { this.style.background = "#4a35a6"; };
            saveBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
        }

        buttonRow.appendChild(selectAllBtn);
        buttonRow.appendChild(deselectAllBtn);
        buttonRow.appendChild(saveBtn);
        container.appendChild(buttonRow);

        var settingsPopup = createPopup({
            title: "Settings",
            content: container,
            width: "360px",
            height: "auto",
            maxHeight: "80%"
        });
        var originalClose = settingsPopup.close;
        settingsPopup.close = function() {
            SETTINGS_POPUP_REF = null;
            originalClose();
        };
        saveBtn.addEventListener("click", function() {
            var newVisibility = {};
            for (var j = 0; j < checkboxes.length; j++) {
                var cb = checkboxes[j];
                newVisibility[cb.dataset.label] = cb.checked;
            }
            setButtonVisibility(newVisibility);

            var newHotkey = hotkeyInput.value.trim();
            if (newHotkey) {
                setPanelHotkey(newHotkey);
                log("Settings: Hotkey saved as " + newHotkey);
            }

            log("Settings: Button visibility saved");
            settingsPopup.close();
            location.reload();
        });
        return settingsPopup;
    }

    // =========================
    // UI SCALING FEATURE
    // =========================


    function scale(value) {
        if (typeof value === 'string' && value.endsWith('px')) {
            return (parseFloat(value) * UI_SCALE) + 'px';
        }
        return (value * UI_SCALE) + 'px';
    }

    function getStoredUIScale() {
        try {
            var stored = localStorage.getItem(STORAGE_UI_SCALE);
            if (stored) {
                var scale = parseFloat(stored);
                if (!isNaN(scale) && scale >= 0.5 && scale <= 1.0) {
                    return scale;
                }
            }
        } catch (e) {}
        return 1.0;
    }

    /**
     * @param {number} scale
     */
    function setStoredUIScale(scale) {
        try {
            localStorage.setItem(STORAGE_UI_SCALE, String(scale));
        } catch (e) {}
    }

    /**
     * @param {number} newScale
     */
    function updateUIScale(newScale) {
        UI_SCALE = Math.max(0.5, Math.min(1.0, newScale));
        setStoredUIScale(UI_SCALE);
    }

    // Initialize UI scale from storage
    UI_SCALE = getStoredUIScale();

    //==========================
    // MAKE PANEL FUNCTIONS
    //==========================
    // This section contains functions used to create and manage the panel UI.
    // These functions are used to create the panel UI and manage its state.
    //==========================
    function makePanel() {
        var glass = isGlassTheme();
        if (glass) injectThemeStylesIfNeeded();
        var prior = document.getElementById(PANEL_ID);
        if (prior) {
            return prior;
        }
        var panel = document.createElement("div");
        panel.id = PANEL_ID;
        if (glass) {
            panel.classList.add(THEME_SCOPE_CLASS);
            panel.classList.add("ie-glass-panel");
        }
        var savedTop = getStoredPos("activityPlanState.panel.top", "20px");
        var savedRight = getStoredPos("activityPlanState.panel.right", "20px");
        var savedSize = getStoredPanelSize();
        panel.style.top = savedTop;
        panel.style.position = "fixed";
        panel.style.right = savedRight;
        panel.style.zIndex = glass ? String(THEME_Z_BASE) : "999999";
        if (!glass) {
            panel.style.background = "#111";
            panel.style.color = "#fff";
            panel.style.border = "1px solid #444";
        }
        panel.style.borderRadius = scale(PANEL_BORDER_RADIUS_PX);
        panel.style.padding = scale(PANEL_PADDING_PX);
        panel.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        panel.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        panel.style.minWidth = scale(PANEL_DEFAULT_WIDTH);
        panel.style.width = savedSize.width || scale(PANEL_DEFAULT_WIDTH);
        if (savedSize.height && savedSize.height !== PANEL_DEFAULT_HEIGHT) {
            panel.style.height = savedSize.height;
        } else {
            panel.style.height = PANEL_DEFAULT_HEIGHT;
        }
        panel.style.boxSizing = "border-box";
        panel.style.overflow = "hidden";

        var headerBar = document.createElement("div");
        if (glass) {
            headerBar.classList.add("ie-glass-panel-header");
            headerBar.style.borderRadius = THEME_RADIUS + "px " + THEME_RADIUS + "px 0 0";
        }
        headerBar.style.position = "relative";
        headerBar.style.display = "grid";
        headerBar.style.gridTemplateColumns = "auto 1fr auto";
        headerBar.style.alignItems = "center";
        headerBar.style.gap = String(scale(PANEL_HEADER_GAP_PX)) + "px";
        headerBar.style.height = String(scale(PANEL_HEADER_HEIGHT_PX)) + "px";
        headerBar.style.boxSizing = "border-box";
        headerBar.style.cursor = "grab";
        headerBar.style.userSelect = "none";
        var leftSpacer = document.createElement("div");
        leftSpacer.style.width = scale(HEADER_LEFT_SPACER_WIDTH_PX);
        var title = document.createElement("div");
        title.textContent = "Automator";
        title.style.fontWeight = HEADER_FONT_WEIGHT;
        title.style.textAlign = "center";
        title.style.justifySelf = "center";
        title.style.transform = "translateX(" + scale(HEADER_TITLE_OFFSET_PX) + ")";
        title.style.paddingBottom = scale(8);
        if (glass) title.style.color = THEME_TEXT_PRIMARY;
        headerBar.appendChild(title);
        headerBar.appendChild(leftSpacer);
        var rightControls = document.createElement("div");
        rightControls.style.display = "inline-flex";
        rightControls.style.alignItems = "center";
        rightControls.style.gap = scale(PANEL_HEADER_GAP_PX);
        var collapseBtn = document.createElement("button");
        collapseBtn.textContent = getPanelCollapsed() ? "Expand" : "Collapse";
        collapseBtn.style.background = "transparent";
        collapseBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        collapseBtn.style.border = "none";
        collapseBtn.style.cursor = "pointer";
        collapseBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        var closeBtn = document.createElement("button");
        closeBtn.textContent = CLOSE_BTN_TEXT;
        closeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        closeBtn.style.background = "transparent";
        closeBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        closeBtn.style.border = "none";
        closeBtn.style.cursor = "pointer";

        var settingsBtn = document.createElement("button");
        settingsBtn.textContent = "⚙";
        settingsBtn.title = "Settings";
        settingsBtn.style.background = "transparent";
        settingsBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        settingsBtn.style.border = "none";
        settingsBtn.style.cursor = "pointer";
        settingsBtn.style.fontSize = scale(16);
        settingsBtn.style.padding = "4px 6px";
        settingsBtn.style.borderRadius = "4px";
        settingsBtn.style.display = "flex";
        settingsBtn.style.alignItems = "center";
        settingsBtn.style.justifyContent = "center";
        settingsBtn.addEventListener("click", function() {
            log("Settings: Button clicked");
            if (SETTINGS_POPUP_REF) {
                log("Settings: Closing existing popup");
                SETTINGS_POPUP_REF.close();
                SETTINGS_POPUP_REF = null;
            } else {
                SETTINGS_POPUP_REF = openSettingsPopup();
            }
        });
        rightControls.appendChild(settingsBtn);
        rightControls.appendChild(collapseBtn);
        rightControls.appendChild(closeBtn);

        headerBar.appendChild(rightControls);
        panel.appendChild(headerBar);
        var bodyContainer = document.createElement("div");
        bodyContainer.style.display = "block";
        bodyContainer.style.height = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
        bodyContainer.style.maxHeight = "calc(100% - " + String(scale(PANEL_HEADER_HEIGHT_PX)) + "px)";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.boxSizing = "border-box";
        bodyContainer.style.padding = scale(12) + "px";
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
        runPlansBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runPlansBtn.style.padding = scale(BUTTON_PADDING_PX);
        runPlansBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runPlansBtn.style.cursor = "pointer";
        runPlansBtn.style.fontWeight = "500";
        runPlansBtn.style.transition = "background 0.2s";
        runPlansBtn.onmouseenter = () => { runPlansBtn.style.background = "#357abd"; };
        runPlansBtn.onmouseleave = () => { runPlansBtn.style.background = "#4a90e2"; };
        var runStudyBtn = document.createElement("button");
        runStudyBtn.textContent = "Update Study Status";
        runStudyBtn.style.background = "#4a90e2";
        runStudyBtn.style.color = "#fff";
        runStudyBtn.style.border = "none";
        runStudyBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runStudyBtn.style.padding = scale(BUTTON_PADDING_PX);
        runStudyBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runStudyBtn.style.cursor = "pointer";
        runStudyBtn.style.fontWeight = "500";
        runStudyBtn.style.transition = "background 0.2s";
        runStudyBtn.onmouseenter = () => { runStudyBtn.style.background = "#357abd"; };
        runStudyBtn.onmouseleave = () => { runStudyBtn.style.background = "#4a90e2"; };
        var runAddCohortBtn = document.createElement("button");
        runAddCohortBtn.textContent = "Add Cohort Subjects";
        runAddCohortBtn.style.background = "#4a90e2";
        runAddCohortBtn.style.color = "#fff";
        runAddCohortBtn.style.border = "none";
        runAddCohortBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runAddCohortBtn.style.padding = scale(BUTTON_PADDING_PX);
        runAddCohortBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runAddCohortBtn.style.cursor = "pointer";
        runAddCohortBtn.style.fontWeight = "500";
        runAddCohortBtn.style.transition = "background 0.2s";
        runAddCohortBtn.onmouseenter = () => { runAddCohortBtn.style.background = "#357abd"; };
        runAddCohortBtn.onmouseleave = () => { runAddCohortBtn.style.background = "#4a90e2"; };
        var runConsentBtn = document.createElement("button");
        runConsentBtn.textContent = "Run ICF Barcode";
        runConsentBtn.style.background = "#4a90e2";
        runConsentBtn.style.color = "#fff";
        runConsentBtn.style.border = "none";
        runConsentBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runConsentBtn.style.padding = scale(BUTTON_PADDING_PX);
        runConsentBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runConsentBtn.style.cursor = "pointer";
        runConsentBtn.style.fontWeight = "500";
        runConsentBtn.style.transition = "background 0.2s";
        runConsentBtn.onmouseenter = () => { runConsentBtn.style.background = "#357abd"; };
        runConsentBtn.onmouseleave = () => { runConsentBtn.style.background = "#4a90e2"; };
        var runAllBtn = document.createElement("button");
        runAllBtn.textContent = "Run Button (1-5)";
        runAllBtn.style.background = "#5cb85c";
        runAllBtn.style.color = "#fff";
        runAllBtn.style.border = "none";
        runAllBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runAllBtn.style.padding = scale(BUTTON_PADDING_PX);
        runAllBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runAllBtn.style.cursor = "pointer";
        runAllBtn.style.fontWeight = "600";
        runAllBtn.style.transition = "background 0.2s";
        runAllBtn.onmouseenter = () => { runAllBtn.style.background = "#449d44"; };
        runAllBtn.onmouseleave = () => { runAllBtn.style.background = "#5cb85c"; };
        var runNonScrnBtn = document.createElement("button");
        runNonScrnBtn.textContent = "Import Cohort Subject";
        runNonScrnBtn.style.background = "#5b43c7";
        runNonScrnBtn.style.color = "#fff";
        runNonScrnBtn.style.border = "none";
        runNonScrnBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runNonScrnBtn.style.padding = scale(BUTTON_PADDING_PX);
        runNonScrnBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runNonScrnBtn.style.cursor = "pointer";
        runNonScrnBtn.style.fontWeight = "500";
        runNonScrnBtn.style.transition = "background 0.2s";
        runNonScrnBtn.onmouseenter = () => { runNonScrnBtn.style.background = "#4a37a0"; };
        runNonScrnBtn.onmouseleave = () => { runNonScrnBtn.style.background = "#5b43c7"; };

        var addExistingSubjectBtn = document.createElement("button");
        addExistingSubjectBtn.textContent = "Add Existing Subject";
        addExistingSubjectBtn.style.background = "#5b43c7";
        addExistingSubjectBtn.style.color = "#fff";
        addExistingSubjectBtn.style.border = "none";
        addExistingSubjectBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        addExistingSubjectBtn.style.padding = scale(BUTTON_PADDING_PX);
        addExistingSubjectBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        addExistingSubjectBtn.style.cursor = "pointer";
        addExistingSubjectBtn.style.fontWeight = "500";
        addExistingSubjectBtn.style.transition = "background 0.2s";
        addExistingSubjectBtn.onmouseenter = () => { addExistingSubjectBtn.style.background = "#4a37a0"; };
        addExistingSubjectBtn.onmouseleave = () => { addExistingSubjectBtn.style.background = "#5b43c7"; };

        var saBuilderBtn = document.createElement("button");
        saBuilderBtn.textContent = "Scheduled Activities Builder";
        saBuilderBtn.style.background = "#5b43c7";
        saBuilderBtn.style.color = "#fff";
        saBuilderBtn.style.border = "none";
        saBuilderBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        saBuilderBtn.style.padding = scale(BUTTON_PADDING_PX);
        saBuilderBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        saBuilderBtn.style.cursor = "pointer";
        saBuilderBtn.style.fontWeight = "500";
        saBuilderBtn.style.transition = "background 0.2s";
        saBuilderBtn.onmouseenter = () => { saBuilderBtn.style.background = "#4a37a0"; };
        saBuilderBtn.onmouseleave = () => { saBuilderBtn.style.background = "#5b43c7"; };

        var runBarcodeBtn = document.createElement("button");
        runBarcodeBtn.textContent = "Pull Barcode";
        runBarcodeBtn.style.background = "#5b43c7";
        runBarcodeBtn.style.color = "#fff";
        runBarcodeBtn.style.border = "none";
        runBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        runBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runBarcodeBtn.style.cursor = "pointer";
        runBarcodeBtn.style.fontWeight = "500";
        runBarcodeBtn.style.transition = "background 0.2s";
        runBarcodeBtn.onmouseenter = () => { runBarcodeBtn.style.background = "#4a37a0"; };
        runBarcodeBtn.onmouseleave = () => { runBarcodeBtn.style.background = "#5b43c7"; };
        var runFormOORBtn = document.createElement("button");
        runFormOORBtn.textContent = "Run Form (OOR) Below Range";
        runFormOORBtn.style.background = "#f0ad4e";
        runFormOORBtn.style.color = "#fff";
        runFormOORBtn.style.border = "none";
        runFormOORBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormOORBtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormOORBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormOORBtn.style.cursor = "pointer";
        runFormOORBtn.style.fontWeight = "500";
        runFormOORBtn.style.transition = "background 0.2s";
        runFormOORBtn.onmouseenter = () => { runFormOORBtn.style.background = "#ec971f"; };
        runFormOORBtn.onmouseleave = () => { runFormOORBtn.style.background = "#f0ad4e"; };
        var runFormOORABtn = document.createElement("button");
        runFormOORABtn.textContent = "Run Form (OOR) Above Range";
        runFormOORABtn.style.background = "#f0ad4e";
        runFormOORABtn.style.color = "#fff";
        runFormOORABtn.style.border = "none";
        runFormOORABtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormOORABtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormOORABtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormOORABtn.style.cursor = "pointer";
        runFormOORABtn.style.fontWeight = "500";
        runFormOORABtn.style.transition = "background 0.2s";
        runFormOORABtn.onmouseenter = () => { runFormOORABtn.style.background = "#ec971f"; };
        runFormOORABtn.onmouseleave = () => { runFormOORABtn.style.background = "#f0ad4e"; };
        var runFormIRBtn = document.createElement("button");
        runFormIRBtn.textContent = "Run Form (In Range)";
        runFormIRBtn.style.background = "#f0ad4e";
        runFormIRBtn.style.color = "#fff";
        runFormIRBtn.style.border = "none";
        runFormIRBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runFormIRBtn.style.padding = scale(BUTTON_PADDING_PX);
        runFormIRBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runFormIRBtn.style.cursor = "pointer";
        runFormIRBtn.style.fontWeight = "500";
        runFormIRBtn.style.transition = "background 0.2s";
        runFormIRBtn.onmouseenter = () => { runFormIRBtn.style.background = "#ec971f"; };
        runFormIRBtn.onmouseleave = () => { runFormIRBtn.style.background = "#f0ad4e"; };
        var parseMethodBtn = document.createElement("button");
        parseMethodBtn.textContent = "Item Method Forms";
        parseMethodBtn.style.background = "#4a90e2";
        parseMethodBtn.style.color = "#fff";
        parseMethodBtn.style.border = "none";
        parseMethodBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        parseMethodBtn.style.padding = scale(BUTTON_PADDING_PX);
        parseMethodBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        parseMethodBtn.style.cursor = "pointer";
        parseMethodBtn.onmouseenter = () => { parseMethodBtn.style.background = "#58a1f5"; };
        parseMethodBtn.onmouseleave = () => { parseMethodBtn.style.background = "#4a90e2"; };

        var searchMethodsBtn = document.createElement("button");
        searchMethodsBtn.textContent = "Search Methods";
        searchMethodsBtn.style.background = "#5b43c7";
        searchMethodsBtn.style.color = "#fff";
        searchMethodsBtn.style.border = "none";
        searchMethodsBtn.style.borderRadius = "6px";
        searchMethodsBtn.style.padding = "8px";
        searchMethodsBtn.style.cursor = "pointer";
        searchMethodsBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        searchMethodsBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };
        var archiveUpdateFormsBtn = document.createElement("button");
        archiveUpdateFormsBtn.textContent = "Archive/Update Forms";
        archiveUpdateFormsBtn.style.background = "#38dae6";
        archiveUpdateFormsBtn.style.color = "#fff";
        archiveUpdateFormsBtn.style.border = "none";
        archiveUpdateFormsBtn.style.borderRadius = "6px";
        archiveUpdateFormsBtn.style.padding = "8px";
        archiveUpdateFormsBtn.style.cursor = "pointer";
        archiveUpdateFormsBtn.style.fontWeight = "500";
        archiveUpdateFormsBtn.style.transition = "background 0.2s";
        archiveUpdateFormsBtn.onmouseenter = function() { this.style.background = "#2bb9c4"; };
        archiveUpdateFormsBtn.onmouseleave = function() { this.style.background = "#38dae6"; };

        var copyFormsBtn = document.createElement("button");
        copyFormsBtn.textContent = "Copy Activity Forms";
        copyFormsBtn.style.background = "#38dae6";
        copyFormsBtn.style.color = "#fff";
        copyFormsBtn.style.border = "none";
        copyFormsBtn.style.borderRadius = "6px";
        copyFormsBtn.style.padding = "8px";
        copyFormsBtn.style.cursor = "pointer";
        copyFormsBtn.style.fontWeight = "500";
        copyFormsBtn.style.transition = "background 0.2s";
        copyFormsBtn.onmouseenter = function() { this.style.background = "#2bb9c4"; };
        copyFormsBtn.onmouseleave = function() { this.style.background = "#38dae6"; };

        var pauseBtn = document.createElement("button");
        pauseBtn.textContent = isPaused() ? "Resume" : "Pause";
        pauseBtn.style.background = "#6c757d";
        pauseBtn.style.color = "#fff";
        pauseBtn.style.border = "none";
        pauseBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        pauseBtn.style.padding = scale(BUTTON_PADDING_PX);
        pauseBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        pauseBtn.style.cursor = "pointer";
        pauseBtn.style.fontWeight = "500";
        pauseBtn.style.transition = "background 0.2s";
        pauseBtn.onmouseenter = () => { pauseBtn.style.background = "#5a6268"; };
        pauseBtn.onmouseleave = () => { pauseBtn.style.background = "#6c757d"; };

        var clearLogsBtn = document.createElement("button");
        clearLogsBtn.textContent = "Clear Logs";
        clearLogsBtn.style.background = "#6c757d";
        clearLogsBtn.style.color = "#fff";
        clearLogsBtn.style.border = "none";
        clearLogsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        clearLogsBtn.style.padding = scale(BUTTON_PADDING_PX);
        clearLogsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        clearLogsBtn.style.cursor = "pointer";
        clearLogsBtn.style.fontWeight = "500";
        clearLogsBtn.style.transition = "background 0.2s";
        clearLogsBtn.onmouseenter = () => { clearLogsBtn.style.background = "#5a6268"; };
        clearLogsBtn.onmouseleave = () => { clearLogsBtn.style.background = "#6c757d"; };

        var toggleLogsBtn = document.createElement("button");
        var logVisible = getLogVisible();
        toggleLogsBtn.textContent = logVisible ? "Hide Logs" : "Show Logs";
        toggleLogsBtn.style.background = "#6c757d";
        toggleLogsBtn.style.color = "#fff";
        toggleLogsBtn.style.border = "none";
        toggleLogsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        toggleLogsBtn.style.padding = scale(BUTTON_PADDING_PX);
        toggleLogsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        toggleLogsBtn.style.cursor = "pointer";
        toggleLogsBtn.style.fontWeight = "500";
        toggleLogsBtn.style.transition = "background 0.2s";
        toggleLogsBtn.onmouseenter = () => { toggleLogsBtn.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = () => { toggleLogsBtn.style.background = "#6c757d"; };

        var runLockSamplePathsBtn = document.createElement("button");
        runLockSamplePathsBtn.textContent = "Lock Sample Paths";
        runLockSamplePathsBtn.style.background = "#4a90e2";
        runLockSamplePathsBtn.style.color = "#fff";
        runLockSamplePathsBtn.style.border = "none";
        runLockSamplePathsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        runLockSamplePathsBtn.style.padding = scale(BUTTON_PADDING_PX);
        runLockSamplePathsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        runLockSamplePathsBtn.style.cursor = "pointer";
        runLockSamplePathsBtn.style.fontWeight = "500";
        runLockSamplePathsBtn.style.transition = "background 0.2s";
        runLockSamplePathsBtn.onmouseenter = () => { runLockSamplePathsBtn.style.background = "#357abd"; };
        runLockSamplePathsBtn.onmouseleave = () => { runLockSamplePathsBtn.style.background = "#4a90e2"; };

        var importEligBtn = document.createElement("button");
        importEligBtn.textContent = "Import I/E";
        importEligBtn.style.background = "#38dae6";
        importEligBtn.style.color = "#fff";
        importEligBtn.style.border = "none";
        importEligBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        importEligBtn.style.padding = scale(BUTTON_PADDING_PX);
        importEligBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        importEligBtn.style.cursor = "pointer";
        importEligBtn.style.fontWeight = "500";
        importEligBtn.style.transition = "background 0.2s";
        importEligBtn.onmouseenter = () => { importEligBtn.style.background = "#2bb9c4"; };
        importEligBtn.onmouseleave = () => { importEligBtn.style.background = "#38dae6"; };

        var findFormBtn = document.createElement("button");
        findFormBtn.textContent = "Find Form";
        findFormBtn.style.background = "#4a90e2";
        findFormBtn.style.color = "#fff";
        findFormBtn.style.border = "none";
        findFormBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findFormBtn.style.padding = scale(BUTTON_PADDING_PX);
        findFormBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findFormBtn.style.cursor = "pointer";
        findFormBtn.onmouseenter = () => { findFormBtn.style.background = "#58a1f5"; };
        findFormBtn.onmouseleave = () => { findFormBtn.style.background = "#4a90e2"; };

        var pullLabBarcodeBtn = document.createElement("button");
        pullLabBarcodeBtn.textContent = BARCODE_LABELS.featureButton;
        pullLabBarcodeBtn.style.background = "#5b43c7";
        pullLabBarcodeBtn.style.color = "#fff";
        pullLabBarcodeBtn.style.border = "none";
        pullLabBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        pullLabBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        pullLabBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        pullLabBarcodeBtn.style.cursor = "pointer";
        pullLabBarcodeBtn.style.fontWeight = "500";
        pullLabBarcodeBtn.style.transition = "background 0.2s";
        pullLabBarcodeBtn.setAttribute("aria-label", BARCODE_LABELS.featureButton);
        pullLabBarcodeBtn.onmouseenter = function () {
            this.style.background = "#4a37a0";
        };
        pullLabBarcodeBtn.onmouseleave = function () {
            this.style.background = "#5b43c7";
        };

        PULL_LAB_BARCODE_BUTTON_REF = pullLabBarcodeBtn;

        var findStudyEventsBtn = document.createElement("button");
        findStudyEventsBtn.textContent = "Find Study Events";
        findStudyEventsBtn.style.background = "#4a90e2";
        findStudyEventsBtn.style.color = "#fff";
        findStudyEventsBtn.style.border = "none";
        findStudyEventsBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        findStudyEventsBtn.style.padding = scale(BUTTON_PADDING_PX);
        findStudyEventsBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        findStudyEventsBtn.style.cursor = "pointer";
        findStudyEventsBtn.onmouseenter = () => { findStudyEventsBtn.style.background = "#58a1f5"; };
        findStudyEventsBtn.onmouseleave = () => { findStudyEventsBtn.style.background = "#4a90e2"; };

        var clearMappingBtn = document.createElement("button");
        clearMappingBtn.textContent = "Clear Mapping";
        clearMappingBtn.style.background = "#38dae6";
        clearMappingBtn.style.color = "#fff";
        clearMappingBtn.style.border = "none";
        clearMappingBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        clearMappingBtn.style.padding = scale(BUTTON_PADDING_PX);
        clearMappingBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        clearMappingBtn.style.cursor = "pointer";
        clearMappingBtn.style.fontWeight = "500";
        clearMappingBtn.style.transition = "background 0.2s";
        clearMappingBtn.onmouseenter = () => { clearMappingBtn.style.background = "#2bb9c4"; };
        clearMappingBtn.onmouseleave = () => { clearMappingBtn.style.background = "#38dae6"; };

        var collectAllBtn = document.createElement("button");
        collectAllBtn.textContent = "Collect All";
        collectAllBtn.style.background = "#f0ad4e";
        collectAllBtn.style.color = "#fff";
        collectAllBtn.style.border = "none";
        collectAllBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        collectAllBtn.style.padding = scale(BUTTON_PADDING_PX);
        collectAllBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        collectAllBtn.style.cursor = "pointer";
        collectAllBtn.style.fontWeight = "500";
        collectAllBtn.style.transition = "background 0.2s";
        collectAllBtn.onmouseenter = () => { collectAllBtn.style.background = "#ec971f"; };
        collectAllBtn.onmouseleave = () => { collectAllBtn.style.background = "#f0ad4e"; };

        // Apply glassmorphism theme to all panel buttons if glass theme is active
        if (glass) {
            var allPanelBtns = [runPlansBtn, runStudyBtn, runAddCohortBtn, runConsentBtn, runAllBtn, runNonScrnBtn, addExistingSubjectBtn, saBuilderBtn, runBarcodeBtn, runFormOORBtn, runFormOORABtn, runFormIRBtn, parseMethodBtn, searchMethodsBtn, archiveUpdateFormsBtn, copyFormsBtn, pauseBtn, clearLogsBtn, toggleLogsBtn, runLockSamplePathsBtn, importEligBtn, findFormBtn, pullLabBarcodeBtn, findStudyEventsBtn, clearMappingBtn, collectAllBtn];
            for (var gi = 0; gi < allPanelBtns.length; gi++) {
                var gb = allPanelBtns[gi];
                gb.className = "ie-btn-primary";
                gb.style.background = "";
                gb.style.color = "";
                gb.style.border = "";
                gb.style.borderRadius = "";
                gb.style.fontWeight = "";
                gb.style.transition = "";
                gb.onmouseenter = null;
                gb.onmouseleave = null;
                gb.style.padding = scale(BUTTON_PADDING_PX);
                gb.style.fontSize = scale(PANEL_FONT_SIZE_PX);
            }
        }

        var panelButtons = [
            { el: runPlansBtn, label: "Lock Activity Plans" },
            { el: runLockSamplePathsBtn, label: "Lock Sample Paths" },
            { el: runStudyBtn, label: "Update Study Status" },
            { el: runAddCohortBtn, label: "Add Cohort Subjects" },
            { el: runConsentBtn, label: "Run ICF Consent" },
            { el: runAllBtn, label: "Run Button (1-5)" },
            { el: runNonScrnBtn, label: "Import Cohort Subjects" },
            { el: runBarcodeBtn, label: "Pull Barcode" },
            { el: pullLabBarcodeBtn, label: "Pull Lab Barcode" },
            { el: addExistingSubjectBtn, label: "Add Existing Subject" },
            { el: searchMethodsBtn, label: "Search Methods" },
            { el: saBuilderBtn, label: "Scheduled Activities Builder" },
            { el: runFormOORBtn, label: "Run Form (OOR) Below Range" },
            { el: runFormOORABtn, label: "Run Form (OOR) Above Range" },
            { el: runFormIRBtn, label: "Run Form (IR) In Range" },
            { el: collectAllBtn, label: "Collect All" },
            { el: importEligBtn, label: "Import I/E" },
            { el: clearMappingBtn, label: "Clear Mapping" },
            { el: archiveUpdateFormsBtn, label: "Archive/Update Forms" },
            { el: copyFormsBtn, label: "Copy Activity Forms"},
            { el: parseMethodBtn, label: "Item Method Forms" },
            { el: findFormBtn, label: "Find Form" },
            { el: findStudyEventsBtn, label: "Find Study Events" },
            { el: pauseBtn, label: "Pause" },
            { el: clearLogsBtn, label: "Clear Logs" },
            { el: toggleLogsBtn, label: "Hide Logs" }
        ];

        for (var bi = 0; bi < panelButtons.length; bi++) {
            var btnItem = panelButtons[bi];
            if (isButtonVisible(btnItem.label)) {
                btnRow.appendChild(btnItem.el);
            }
        }

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
        status.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        status.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        status.style.border = glass ? ("1px solid " + THEME_SURFACE_INNER_BORDER) : "1px solid #333";
        status.style.borderRadius = scale(STATUS_BORDER_RADIUS_PX);
        status.style.padding = scale(STATUS_PADDING_PX);
        status.style.fontSize = scale(STATUS_FONT_SIZE_PX);
        status.style.whiteSpace = "pre-wrap";
        if (glass) status.style.color = THEME_TEXT_PRIMARY;
        status.textContent = "Ready";
        bodyContainer.appendChild(status);


        // UI Scale Control
        var scaleControl = document.createElement("div");
        scaleControl.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        scaleControl.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        scaleControl.style.border = glass ? ("1px solid " + THEME_SURFACE_INNER_BORDER) : "1px solid #333";
        scaleControl.style.borderRadius = scale(STATUS_BORDER_RADIUS_PX);
        scaleControl.style.padding = scale(STATUS_PADDING_PX);
        scaleControl.style.fontSize = scale(STATUS_FONT_SIZE_PX);

        var scaleLabel = document.createElement("div");
        scaleLabel.textContent = "UI Scale: " + Math.round(UI_SCALE * 100) + "%";
        scaleLabel.style.marginBottom = "4px";
        scaleLabel.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";

        var scaleSlider = document.createElement("input");
        scaleSlider.type = "range";
        scaleSlider.min = "50";
        scaleSlider.max = "100";
        scaleSlider.value = String(UI_SCALE * 100);
        scaleSlider.step = "10";
        scaleSlider.style.width = "100%";
        scaleSlider.style.cursor = "pointer";

        scaleSlider.addEventListener("input", function() {
            var newScale = parseFloat(this.value) / 100;
            updateUIScale(newScale);
            scaleLabel.textContent = "UI Scale: " + Math.round(newScale * 100) + "%";
            status.textContent = "UI Scale updated to " + Math.round(newScale * 100) + "% - Refresh to see changes";
        });

        scaleSlider.addEventListener("change", function() {
            log("UI Scale changed to " + Math.round(UI_SCALE * 100) + "%");
        });

        scaleControl.appendChild(scaleLabel);
        scaleControl.appendChild(scaleSlider);
        bodyContainer.appendChild(scaleControl);
        var logBox = document.createElement("div");
        logBox.id = LOG_ID;
        logBox.style.marginTop = scale(LOG_MARGIN_TOP_PX);
        logBox.style.height = scale(LOG_HEIGHT_PX);
        logBox.style.overflowY = "auto";
        logBox.style.background = glass ? THEME_SURFACE_BG : "#141414";
        logBox.style.border = glass ? ("1px solid " + THEME_SURFACE_INNER_BORDER) : "1px solid #333";
        logBox.style.borderRadius = scale(LOG_BORDER_RADIUS_PX);
        logBox.style.padding = scale(LOG_PADDING_PX);
        logBox.style.fontSize = scale(LOG_FONT_SIZE_PX);
        logBox.style.color = glass ? THEME_TEXT_MUTED : "#00000";
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
        
        pauseBtn.addEventListener("click", function () {
            var nowPaused = isPaused();
            if (nowPaused) {
                setPaused(false);
                pauseBtn.textContent = "Pause";
                status.textContent = "Resumed";
                SA_BUILDER_PAUSE = false;
                CLEAR_MAPPING_PAUSE = false;
                log("Resumed");
            } else {
                setPaused(true);
                pauseBtn.textContent = "Resume";
                status.textContent = "Paused";
                log("Paused");

                // Update all active popup statuses to show "Paused"
                if (RUN_ALL_POPUP_REF && RUN_ALL_POPUP_REF.element) {
                    try {
                        var runAllStatusDiv = RUN_ALL_POPUP_REF.element.querySelector("#runAllStatus");
                        if (runAllStatusDiv) {
                            runAllStatusDiv.textContent = "⏸ Paused";
                            runAllStatusDiv.style.color = "#ffa500";
                        }
                        var runAllLoading = RUN_ALL_POPUP_REF.element.querySelector("#runAllLoading");
                        if (runAllLoading) {
                            runAllLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                if (LOCK_SAMPLE_PATHS_POPUP_REF && LOCK_SAMPLE_PATHS_POPUP_REF.element) {
                    try {
                        var lockStatusDiv = LOCK_SAMPLE_PATHS_POPUP_REF.element.querySelector("#lockSamplePathsStatus");
                        if (lockStatusDiv) {
                            lockStatusDiv.textContent = "⏸ Paused";
                            lockStatusDiv.style.color = "#ffa500";
                        }
                        var lockLoading = LOCK_SAMPLE_PATHS_POPUP_REF.element.querySelector("#lockSamplePathsLoading");
                        if (lockLoading) {
                            lockLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                if (IMPORT_ELIG_POPUP_REF && IMPORT_ELIG_POPUP_REF.element) {
                    try {
                        var eligStatusDiv = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligStatus");
                        if (eligStatusDiv) {
                            eligStatusDiv.textContent = "⏸ Paused";
                            eligStatusDiv.style.color = "#ffa500";
                        }
                        var eligLoading = IMPORT_ELIG_POPUP_REF.element.querySelector("#importEligLoading");
                        if (eligLoading) {
                            eligLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                if (IMPORT_COHORT_POPUP_REF && IMPORT_COHORT_POPUP_REF.element) {
                    try {
                        var cohortStatusDiv = IMPORT_COHORT_POPUP_REF.element.querySelector("#importCohortStatus");
                        if (cohortStatusDiv) {
                            cohortStatusDiv.textContent = "⏸ Paused";
                            cohortStatusDiv.style.color = "#ffa500";
                        }
                        var cohortLoading = IMPORT_COHORT_POPUP_REF.element.querySelector("#importCohortLoading");
                        if (cohortLoading) {
                            cohortLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                if (COLLECT_ALL_POPUP_REF && COLLECT_ALL_POPUP_REF.element) {
                    try {
                        var collectStatusDiv = COLLECT_ALL_POPUP_REF.element.querySelector("#collectAllStatus");
                        if (collectStatusDiv) {
                            collectStatusDiv.textContent = "⏸ Paused";
                            collectStatusDiv.style.color = "#ffa500";
                        }
                        var collectLoading = COLLECT_ALL_POPUP_REF.element.querySelector("#collectAllLoading");
                        if (collectLoading) {
                            collectLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                if (CLEAR_MAPPING_POPUP_REF && CLEAR_MAPPING_POPUP_REF.element) {
                    try {
                        var clearStatusDiv = CLEAR_MAPPING_POPUP_REF.element.querySelector("#clearMappingStatus");
                        if (clearStatusDiv) {
                            clearStatusDiv.textContent = "⏸ Paused";
                            clearStatusDiv.style.color = "#ffa500";
                        }
                        var clearLoading = CLEAR_MAPPING_POPUP_REF.element.querySelector("#clearMappingLoading");
                        if (clearLoading) {
                            clearLoading.textContent = "";
                        }
                    } catch (e) {}
                }

                clearAllRunState();
                COLLECT_ALL_CANCELLED = true;
                SA_BUILDER_PAUSE = true;
                SA_BUILDER_CANCELLED = true;
                CLEAR_MAPPING_CANCELED = true;
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
        
        collapseBtn.addEventListener("click", function () {
            var collapsed = getPanelCollapsed();
            if (collapsed) {
                setPanelCollapsed(false);
            } else {
                setPanelCollapsed(true);
            }
            updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar);
        });
        
        saBuilderBtn.addEventListener("click", async function () {
            SA_BUILDER_CANCELLED = false;
            log("SA Builder: button clicked");
            if (SA_BUILDER_PAUSE) {
                log("SA Builder: Paused");
                return;
            }
            log("SA Builder: starting");
            await runSABuilder();
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

        // Check for Add Existing Subject mode first
        var runModeRaw = null;
        try {
            runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
        } catch (e) {
            runModeRaw = null;
        }
        
        var onShow = isShowPage();
        if (onShow) {
            processShowPageIfAuto();
            return;
        }
        var onStudyShow = isStudyShowPage();

        // Check for manual navigation away from Import Eligibility process
        var pendingPopup = null;
        try {
            pendingPopup = localStorage.getItem(STORAGE_ELIG_IMPORT_PENDING_POPUP);
        } catch (e) {
            pendingPopup = null;
        }

        // Re-use runModeRaw from earlier check or get it again if needed
        if (!runModeRaw) {
            try {
                runModeRaw = localStorage.getItem(STORAGE_RUN_MODE);
            } catch (e) {
                runModeRaw = null;
            }
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();