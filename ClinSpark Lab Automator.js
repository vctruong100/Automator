
// ==UserScript==
// @name        ClinSpark Lab Automator
// @namespace   vinh.activity.plan.state
// @version     1.0.0
// @description Automate various tasks in ClinSpark platform
// @match       https://cenexel.clinspark.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Lab%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/main/ClinSpark%20Lab%20Automator.js
// @run-at      document-idle
// @grant       GM.openInTab
// @grant       GM_openInTab
// @grant       GM.xmlHttpRequest
// ==/UserScript==

(function () {
    var STORAGE_PANEL_TOP = "activityPlanState.panel.top";
    var STORAGE_PANEL_RIGHT = "activityPlanState.panel.right";
    var PANEL_ID = "activityPlanStatePanel";
    var LOG_ID = "activityPlanStateLog";
    var STORAGE_PAUSED = "activityPlanState.paused";
    var btnRowRef = null;
    var PENDING_BUTTONS = [];
    var STORAGE_BARCODE_SUBJECT_TEXT = "activityPlanState.barcode.subjectText";
    var STORAGE_BARCODE_SUBJECT_ID = "activityPlanState.barcode.subjectId";
    var STORAGE_BARCODE_RESULT = "activityPlanState.barcode.result";
    var STORAGE_PANEL_COLLAPSED = "activityPlanState.panel.collapsed";
    var STORAGE_PANEL_WIDTH = "activityPlanState.panel.width";
    var STORAGE_PANEL_HEIGHT = "activityPlanState.panel.height";
    var STORAGE_LOG_VISIBLE = "activityPlanState.log.visible";

    // UI Scale Constants
    var UI_SCALE = 1.0; // Master scale factor (will be initialized after function definitions)
    var PANEL_DEFAULT_WIDTH = 340;
    var PANEL_DEFAULT_HEIGHT = "auto";
    var PANEL_HEADER_HEIGHT_PX = 30;
    var PANEL_HEADER_GAP_PX = 8;
    var PANEL_MAX_WIDTH_PX = 60;
    var PANEL_PADDING_PX = 12;
    var PANEL_BORDER_RADIUS_PX = 8;
    var PANEL_FONT_SIZE_PX = 14;
    var HEADER_FONT_WEIGHT = "600";
    var HEADER_TITLE_OFFSET_PX = 16;
    var HEADER_LEFT_SPACER_WIDTH_PX = 32;
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

    var BARCODE_START_TS = 0;
    var BARCODE_BG_TAB = null;

    const STORAGE_PANEL_HIDDEN = "activityPlanState.panel.hidden";
    const STORAGE_PANEL_HOTKEY = "activityPlanState.panel.hotkey";
    const PANEL_TOGGLE_KEY = "F2";

    // Auto-Refresh Feature
    var STORAGE_AUTO_REFRESH_ACTIVE = "activityPlanState.autoRefresh.active";
    var STORAGE_AUTO_REFRESH_INTERVAL = "activityPlanState.autoRefresh.intervalMinutes";
    var STORAGE_AUTO_REFRESH_COUNT = "activityPlanState.autoRefresh.refreshCount";
    var STORAGE_AUTO_REFRESH_NEXT_TIME = "activityPlanState.autoRefresh.nextRefreshTime";
    var AUTO_REFRESH_TIMER_ID = null;
    var AUTO_REFRESH_PANEL_ID = "autoRefreshStatusPanel";

    
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

    //==================================================
    // Pull Lab Barcode Feature
    //==================================================
    // 
    //==================================================
    function setAriaBusyOn() {
        log("[PullLabBarcode] setAriaBusyOn: setting aria-busy on body");
        try {
            var target = document.querySelector(BARCODE_ATTRS.ariaBusyTarget);
            if (target) {
                target.setAttribute(BARCODE_ATTRS.ariaBusyAttr, "true");
            }
        } catch (e) {
            log("[PullLabBarcode] setAriaBusyOn: error " + String(e));
        }
    }

    function setAriaBusyOff() {
        log("[PullLabBarcode] setAriaBusyOff: clearing aria-busy on body");
        try {
            var target = document.querySelector(BARCODE_ATTRS.ariaBusyTarget);
            if (target) {
                target.removeAttribute(BARCODE_ATTRS.ariaBusyAttr);
            }
        } catch (e) {
            log("[PullLabBarcode] setAriaBusyOff: error " + String(e));
        }
    }

    function stopPullLabBarcode() {
        log("[PullLabBarcode] stopPullLabBarcode: initiating full teardown");
        PULL_LAB_BARCODE_STOPPED = true;
        PULL_LAB_BARCODE_RUNNING = false;
        var i;
        i = 0;
        while (i < PULL_LAB_BARCODE_TIMEOUTS_LIST.length) {
            try {
                clearTimeout(PULL_LAB_BARCODE_TIMEOUTS_LIST[i]);
            } catch (e) {}
            i = i + 1;
        }
        PULL_LAB_BARCODE_TIMEOUTS_LIST = [];
        i = 0;
        while (i < PULL_LAB_BARCODE_INTERVALS_LIST.length) {
            try {
                clearInterval(PULL_LAB_BARCODE_INTERVALS_LIST[i]);
            } catch (e) {}
            i = i + 1;
        }
        PULL_LAB_BARCODE_INTERVALS_LIST = [];
        i = 0;
        while (i < PULL_LAB_BARCODE_OBSERVERS_LIST.length) {
            try {
                PULL_LAB_BARCODE_OBSERVERS_LIST[i].disconnect();
            } catch (e) {}
            i = i + 1;
        }
        PULL_LAB_BARCODE_OBSERVERS_LIST = [];
        i = 0;
        while (i < PULL_LAB_BARCODE_RAF_IDS.length) {
            try {
                cancelAnimationFrame(PULL_LAB_BARCODE_RAF_IDS[i]);
            } catch (e) {}
            i = i + 1;
        }
        PULL_LAB_BARCODE_RAF_IDS = [];
        i = 0;
        while (i < PULL_LAB_BARCODE_IDLE_IDS.length) {
            try {
                if (typeof cancelIdleCallback === "function") {
                    cancelIdleCallback(PULL_LAB_BARCODE_IDLE_IDS[i]);
                }
            } catch (e) {}
            i = i + 1;
        }
        PULL_LAB_BARCODE_IDLE_IDS = [];
        try {
            var openModal = document.querySelector(BARCODE_SELECTORS.bootboxVisibleModal);
            if (openModal) {
                var modalCloseBtn = openModal.querySelector(".bootbox-close-button, button[data-dismiss='modal']");
                if (modalCloseBtn) {
                    modalCloseBtn.click();
                    log("[PullLabBarcode] stopPullLabBarcode: closed open bootbox modal");
                }
            }
        } catch (e) {
            log("[PullLabBarcode] stopPullLabBarcode: error closing bootbox modal " + String(e));
        }
        if (PULL_LAB_BARCODE_POPUP_REF) {
            var popupToClose = PULL_LAB_BARCODE_POPUP_REF;
            PULL_LAB_BARCODE_POPUP_REF = null;
            try {
                popupToClose.close();
            } catch (e) {}
        }
        setAriaBusyOff();
        PULL_LAB_BARCODE_STATUS_MAP = {};
        PULL_LAB_BARCODE_LEFT_LIST_EL = null;
        PULL_LAB_BARCODE_RIGHT_LIST_EL = null;
        PULL_LAB_BARCODE_SUMMARY_EL = null;
        PULL_LAB_BARCODE_ARIA_LIVE_EL = null;
        BARCODE_COUNTERS.total = 0;
        BARCODE_COUNTERS.processed = 0;
        BARCODE_COUNTERS.verified = 0;
        BARCODE_COUNTERS.skipped = 0;
        BARCODE_COUNTERS.failures = 0;
        BARCODE_COUNTERS.pending = 0;
        if (PULL_LAB_BARCODE_BUTTON_REF) {
            try {
                PULL_LAB_BARCODE_BUTTON_REF.focus();
                log("[PullLabBarcode] stopPullLabBarcode: focus restored to feature button");
            } catch (e) {}
        }
        log("[PullLabBarcode] stopPullLabBarcode: teardown complete");
    }

    function extractBarcodeValue(iconEl) {
        log("[PullLabBarcode] extractBarcodeValue: checking icon element");
        if (!iconEl) {
            log("[PullLabBarcode] extractBarcodeValue: icon element is null");
            return null;
        }
        var className = iconEl.className || "";
        var match = BARCODE_REGEX.barcodeClassToken.exec(className);
        if (match && match[1]) {
            var val = match[1].trim();
            log("[PullLabBarcode] extractBarcodeValue: extracted '" + val + "' from class");
            return val;
        }
        var tooltip = iconEl.getAttribute(BARCODE_SELECTORS.tooltipAttr) || "";
        if (tooltip) {
            var fallbackStr = tooltip.replace(/^(Scan Required:|Verified:)\s*/i, "").trim();
            var tokens = fallbackStr.split(/\s+/);
            if (tokens.length > 0 && tokens[0].length > 0) {
                log("[PullLabBarcode] extractBarcodeValue: fallback extracted '" + tokens[0] + "' from tooltip");
                return tokens[0];
            }
        }
        log("[PullLabBarcode] extractBarcodeValue: no barcode found");
        return null;
    }

    function announceBarcodeAriaLive(message) {
        if (!PULL_LAB_BARCODE_ARIA_LIVE_EL) {
            return;
        }
        try {
            PULL_LAB_BARCODE_ARIA_LIVE_EL.textContent = message;
        } catch (e) {}
    }

    function updateBarcodeRightPanelStatus(barcodeKey, newStatus, detailOptional) {
        log("[PullLabBarcode] updateBarcodeRightPanelStatus: barcode='" + barcodeKey + "' status='" + newStatus + "'");
        PULL_LAB_BARCODE_STATUS_MAP[barcodeKey] = newStatus;
        if (!PULL_LAB_BARCODE_RIGHT_LIST_EL) {
            return;
        }
        var statusColors = {};
        statusColors[BARCODE_LABELS.statusPending] = "#888";
        statusColors[BARCODE_LABELS.statusProcessing] = "#f0ad4e";
        statusColors[BARCODE_LABELS.statusVerified] = "#5cb85c";
        statusColors[BARCODE_LABELS.statusSkippedVerified] = "#5bc0de";
        statusColors[BARCODE_LABELS.statusFailed] = "#d9534f";
        statusColors[BARCODE_LABELS.statusStopped] = "#999";
        var rafId = requestAnimationFrame(function () {
            try {
                var escapedKey = barcodeKey.replace(/['"\\]/g, "\\$&");
                var existingRow = PULL_LAB_BARCODE_RIGHT_LIST_EL.querySelector("[data-barcode-key='" + escapedKey + "']");
                if (existingRow) {
                    var statusSpan = existingRow.querySelector("[data-barcode-status]");
                    if (statusSpan) {
                        statusSpan.textContent = newStatus;
                        statusSpan.style.color = statusColors[newStatus] || "#fff";
                    }
                    if (detailOptional) {
                        var detailSpan = existingRow.querySelector("[data-barcode-detail]");
                        if (detailSpan) {
                            detailSpan.textContent = detailOptional;
                        } else {
                            detailSpan = document.createElement("span");
                            detailSpan.setAttribute("data-barcode-detail", "");
                            detailSpan.style.fontSize = "11px";
                            detailSpan.style.color = "#777";
                            detailSpan.style.marginLeft = "6px";
                            detailSpan.textContent = detailOptional;
                            existingRow.appendChild(detailSpan);
                        }
                    }
                } else {
                    var row = document.createElement("div");
                    row.setAttribute("data-barcode-right-item", "");
                    row.setAttribute("data-barcode-key", barcodeKey);
                    row.style.padding = "4px 10px";
                    row.style.display = "flex";
                    row.style.alignItems = "center";
                    row.style.gap = "8px";
                    row.style.fontSize = "12px";
                    row.style.borderBottom = "1px solid #2a2a2a";
                    var labelSpan = document.createElement("span");
                    labelSpan.style.flex = "1";
                    labelSpan.style.color = "#ddd";
                    labelSpan.style.overflow = "hidden";
                    labelSpan.style.textOverflow = "ellipsis";
                    labelSpan.style.whiteSpace = "nowrap";
                    labelSpan.textContent = barcodeKey;
                    var newStatusSpan = document.createElement("span");
                    newStatusSpan.setAttribute("data-barcode-status", "");
                    newStatusSpan.style.fontWeight = "600";
                    newStatusSpan.style.whiteSpace = "nowrap";
                    newStatusSpan.textContent = newStatus;
                    newStatusSpan.style.color = statusColors[newStatus] || "#fff";
                    row.appendChild(labelSpan);
                    row.appendChild(newStatusSpan);
                    if (detailOptional) {
                        var newDetailSpan = document.createElement("span");
                        newDetailSpan.setAttribute("data-barcode-detail", "");
                        newDetailSpan.style.fontSize = "11px";
                        newDetailSpan.style.color = "#777";
                        newDetailSpan.textContent = detailOptional;
                        row.appendChild(newDetailSpan);
                    }
                    PULL_LAB_BARCODE_RIGHT_LIST_EL.appendChild(row);
                }
            } catch (e) {
                log("[PullLabBarcode] updateBarcodeRightPanelStatus: error " + String(e));
            }
        });
        PULL_LAB_BARCODE_RAF_IDS.push(rafId);
        announceBarcodeAriaLive(barcodeKey + ": " + newStatus);
    }

    function updateBarcodeRightPanelSummary(counters) {
        log("[PullLabBarcode] updateBarcodeRightPanelSummary: updating summary");
        if (!PULL_LAB_BARCODE_SUMMARY_EL) {
            return;
        }
        var pct = 0;
        if (counters.total > 0) {
            pct = Math.round(((counters.processed + counters.skipped) / counters.total) * 100);
        }
        var rafId = requestAnimationFrame(function () {
            try {
                PULL_LAB_BARCODE_SUMMARY_EL.innerHTML = "";
                var items = [
                    { label: "Total", value: counters.total, color: "#fff" },
                    { label: "Processed", value: counters.processed, color: "#ccc" },
                    { label: "Verified", value: counters.verified, color: "#5cb85c" },
                    { label: "Skipped", value: counters.skipped, color: "#5bc0de" },
                    { label: "Failed", value: counters.failures, color: "#d9534f" },
                    { label: "Pending", value: counters.pending, color: "#888" }
                ];
                var si = 0;
                while (si < items.length) {
                    var span = document.createElement("span");
                    span.style.color = items[si].color;
                    span.textContent = items[si].label + ": " + items[si].value;
                    PULL_LAB_BARCODE_SUMMARY_EL.appendChild(span);
                    si = si + 1;
                }
                var pctSpan = document.createElement("span");
                pctSpan.style.fontWeight = "600";
                if (pct >= 100) {
                    pctSpan.style.color = "#5cb85c";
                } else {
                    pctSpan.style.color = "#f0ad4e";
                }
                pctSpan.textContent = pct + "% complete";
                PULL_LAB_BARCODE_SUMMARY_EL.appendChild(pctSpan);
            } catch (e) {
                log("[PullLabBarcode] updateBarcodeRightPanelSummary: error " + String(e));
            }
        });
        PULL_LAB_BARCODE_RAF_IDS.push(rafId);
    }

    function barcodeYieldToUI() {
        return new Promise(function (resolve) {
            if (typeof requestIdleCallback === "function") {
                var idleId = requestIdleCallback(function () {
                    resolve();
                });
                PULL_LAB_BARCODE_IDLE_IDS.push(idleId);
            } else {
                var tid = setTimeout(function () {
                    resolve();
                }, BARCODE_TIMEOUTS.idleBetweenItemsMs);
                PULL_LAB_BARCODE_TIMEOUTS_LIST.push(tid);
            }
        });
    }

    function openBarcodeProgressPanel() {
        log("[PullLabBarcode] openBarcodeProgressPanel: building progress panel");
        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "12px";
        container.style.height = "100%";
        container.style.boxSizing = "border-box";

        var ariaLive = document.createElement("div");
        ariaLive.setAttribute("aria-live", "polite");
        ariaLive.setAttribute("aria-atomic", "true");
        ariaLive.setAttribute("role", "status");
        ariaLive.style.position = "absolute";
        ariaLive.style.width = "1px";
        ariaLive.style.height = "1px";
        ariaLive.style.overflow = "hidden";
        ariaLive.style.clip = "rect(0,0,0,0)";
        ariaLive.style.whiteSpace = "nowrap";
        container.appendChild(ariaLive);
        PULL_LAB_BARCODE_ARIA_LIVE_EL = ariaLive;

        var splitContainer = document.createElement("div");
        splitContainer.style.display = "flex";
        splitContainer.style.gap = "12px";
        splitContainer.style.flex = "1";
        splitContainer.style.minHeight = "0";
        splitContainer.style.overflow = "hidden";

        var leftPanel = document.createElement("div");
        leftPanel.style.flex = "1";
        leftPanel.style.display = "flex";
        leftPanel.style.flexDirection = "column";
        leftPanel.style.border = "1px solid #333";
        leftPanel.style.borderRadius = "6px";
        leftPanel.style.background = "#1a1a1a";
        leftPanel.style.overflow = "hidden";

        var leftHeader = document.createElement("div");
        leftHeader.textContent = "Discovered Icons";
        leftHeader.style.padding = "8px 10px";
        leftHeader.style.fontWeight = "600";
        leftHeader.style.fontSize = "13px";
        leftHeader.style.borderBottom = "1px solid #333";
        leftHeader.style.color = "#ccc";
        leftPanel.appendChild(leftHeader);

        var leftSearch = document.createElement("input");
        leftSearch.type = "text";
        leftSearch.placeholder = "Search icons...";
        leftSearch.setAttribute("aria-label", "Search discovered icons");
        leftSearch.style.width = "100%";
        leftSearch.style.padding = "6px 10px";
        leftSearch.style.background = "#222";
        leftSearch.style.color = "#fff";
        leftSearch.style.border = "none";
        leftSearch.style.borderBottom = "1px solid #333";
        leftSearch.style.boxSizing = "border-box";
        leftSearch.style.fontSize = "12px";
        leftSearch.style.outline = "none";
        leftPanel.appendChild(leftSearch);

        var leftList = document.createElement("div");
        leftList.style.flex = "1";
        leftList.style.overflowY = "auto";
        leftList.style.padding = "4px 0";
        leftPanel.appendChild(leftList);
        PULL_LAB_BARCODE_LEFT_LIST_EL = leftList;

        leftSearch.addEventListener("input", function () {
            var term = leftSearch.value.toLowerCase();
            var items = leftList.querySelectorAll("[data-barcode-left-item]");
            var si = 0;
            while (si < items.length) {
                var txt = (items[si].textContent || "").toLowerCase();
                if (term === "" || txt.indexOf(term) !== -1) {
                    items[si].style.display = "";
                } else {
                    items[si].style.display = "none";
                }
                si = si + 1;
            }
        });

        var rightPanel = document.createElement("div");
        rightPanel.style.flex = "1";
        rightPanel.style.display = "flex";
        rightPanel.style.flexDirection = "column";
        rightPanel.style.border = "1px solid #333";
        rightPanel.style.borderRadius = "6px";
        rightPanel.style.background = "#1a1a1a";
        rightPanel.style.overflow = "hidden";

        var rightHeader = document.createElement("div");
        rightHeader.textContent = "Status";
        rightHeader.style.padding = "8px 10px";
        rightHeader.style.fontWeight = "600";
        rightHeader.style.fontSize = "13px";
        rightHeader.style.borderBottom = "1px solid #333";
        rightHeader.style.color = "#ccc";
        rightPanel.appendChild(rightHeader);

        var rightSearch = document.createElement("input");
        rightSearch.type = "text";
        rightSearch.placeholder = "Search status...";
        rightSearch.setAttribute("aria-label", "Search item statuses");
        rightSearch.style.width = "100%";
        rightSearch.style.padding = "6px 10px";
        rightSearch.style.background = "#222";
        rightSearch.style.color = "#fff";
        rightSearch.style.border = "none";
        rightSearch.style.borderBottom = "1px solid #333";
        rightSearch.style.boxSizing = "border-box";
        rightSearch.style.fontSize = "12px";
        rightSearch.style.outline = "none";
        rightPanel.appendChild(rightSearch);

        var rightList = document.createElement("div");
        rightList.style.flex = "1";
        rightList.style.overflowY = "auto";
        rightList.style.padding = "4px 0";
        rightPanel.appendChild(rightList);
        PULL_LAB_BARCODE_RIGHT_LIST_EL = rightList;

        rightSearch.addEventListener("input", function () {
            var term = rightSearch.value.toLowerCase();
            var items = rightList.querySelectorAll("[data-barcode-right-item]");
            var si = 0;
            while (si < items.length) {
                var txt = (items[si].textContent || "").toLowerCase();
                if (term === "" || txt.indexOf(term) !== -1) {
                    items[si].style.display = "";
                } else {
                    items[si].style.display = "none";
                }
                si = si + 1;
            }
        });

        splitContainer.appendChild(leftPanel);
        splitContainer.appendChild(rightPanel);
        container.appendChild(splitContainer);

        var summaryFooter = document.createElement("div");
        summaryFooter.style.padding = "8px 10px";
        summaryFooter.style.background = "#1a1a1a";
        summaryFooter.style.border = "1px solid #333";
        summaryFooter.style.borderRadius = "6px";
        summaryFooter.style.fontSize = "12px";
        summaryFooter.style.color = "#ccc";
        summaryFooter.style.display = "flex";
        summaryFooter.style.flexWrap = "wrap";
        summaryFooter.style.gap = "8px";
        summaryFooter.style.justifyContent = "space-between";
        summaryFooter.textContent = "Waiting to start...";
        container.appendChild(summaryFooter);
        PULL_LAB_BARCODE_SUMMARY_EL = summaryFooter;

        var popup = createPopup({
            title: BARCODE_LABELS.progressTitle,
            content: container,
            width: "700px",
            height: "500px",
            onClose: function () {
                log("[PullLabBarcode] openBarcodeProgressPanel: popup close triggered");
                stopPullLabBarcode();
            }
        });

        PULL_LAB_BARCODE_POPUP_REF = popup;
        try {
            leftSearch.focus();
        } catch (e) {}
        log("[PullLabBarcode] openBarcodeProgressPanel: panel created");
        return popup;
    }

    async function collectBarcodeIcons() {
        log("[PullLabBarcode] collectBarcodeIcons: starting icon collection");
        var attempts = 0;
        var maxAttempts = BARCODE_RETRY.collectRetries;
        var result = [];

        while (attempts < maxAttempts) {
            if (PULL_LAB_BARCODE_STOPPED) {
                log("[PullLabBarcode] collectBarcodeIcons: stopped during collection");
                return [];
            }

            var presenceCheck = document.querySelectorAll(BARCODE_SELECTORS.presenceCheckAnyIcon);
            log("[PullLabBarcode] collectBarcodeIcons: attempt " + String(attempts + 1) + ", found " + String(presenceCheck.length) + " fa-barcode icons total");

            if (presenceCheck.length === 0) {
                attempts = attempts + 1;
                if (attempts < maxAttempts) {
                    log("[PullLabBarcode] collectBarcodeIcons: no icons found, retrying");
                    await sleep(BARCODE_TIMEOUTS.waitCollectIconsMs / maxAttempts);
                    continue;
                }
                log("[PullLabBarcode] collectBarcodeIcons: no barcode icons found after all retries");
                return [];
            }

            var allIcons = document.querySelectorAll(BARCODE_SELECTORS.barcodeIcon);
            log("[PullLabBarcode] collectBarcodeIcons: found " + String(allIcons.length) + " icons matching barcodeIcon selector");

            result = [];

            var ci = 0;
            while (ci < allIcons.length) {
                if (PULL_LAB_BARCODE_STOPPED) {
                    log("[PullLabBarcode] collectBarcodeIcons: stopped during filtering");
                    return [];
                }

                var icon = allIcons[ci];
                var isVerifiedByClass = icon.classList.contains(BARCODE_SELECTORS.verifiedClass);
                var tooltip = icon.getAttribute(BARCODE_SELECTORS.tooltipAttr) || "";
                var isVerifiedByTooltip = BARCODE_REGEX.verifiedPrefix.test(tooltip);

                if (isVerifiedByClass || isVerifiedByTooltip) {
                    log("[PullLabBarcode] collectBarcodeIcons: icon " + String(ci) + " already verified, skipping");
                    ci = ci + 1;
                    continue;
                }

                var isRequired = BARCODE_REGEX.requiredPrefix.test(tooltip);
                if (!isRequired) {
                    log("[PullLabBarcode] collectBarcodeIcons: icon " + String(ci) + " tooltip does not start with 'Scan Required:', skipping");
                    ci = ci + 1;
                    continue;
                }

                var barcode = extractBarcodeValue(icon);
                if (!barcode) {
                    log("[PullLabBarcode] collectBarcodeIcons: icon " + String(ci) + " has no extractable barcode, skipping");
                    ci = ci + 1;
                    continue;
                }

                var parentRow = icon.closest(BARCODE_SELECTORS.itemRow);
                if (parentRow) {
                    var rowDisplay = parentRow.style.display;
                    if (rowDisplay === "none") {
                        log("[PullLabBarcode] collectBarcodeIcons: icon " + String(ci) + " parent row is hidden (display:none), skipping");
                        ci = ci + 1;
                        continue;
                    }
                }

                var label = tooltip.replace(/^Scan Required:?\s*/i, "").trim();
                if (!label || label.length === 0) {
                    label = barcode;
                }

                var itemName = "";
                if (parentRow) {
                    var textCell = parentRow.querySelector(BARCODE_SELECTORS.itemTextCell);
                    if (textCell) {
                        var textCellClone = textCell.cloneNode(true);
                        var helpTexts = textCellClone.querySelectorAll(BARCODE_SELECTORS.itemHelpText);
                        var hi = 0;
                        while (hi < helpTexts.length) {
                            helpTexts[hi].remove();
                            hi = hi + 1;
                        }
                        itemName = (textCellClone.textContent || "").trim();
                    }
                }
                if (!itemName || itemName.length === 0) {
                    itemName = label;
                }

                var statusKey = barcode + " (#" + String(result.length + 1) + ")";

                result.push({
                    el: icon,
                    barcode: barcode,
                    label: label,
                    itemName: itemName,
                    statusKey: statusKey
                });

                log("[PullLabBarcode] collectBarcodeIcons: added icon " + String(result.length) + " barcode='" + barcode + "' itemName='" + itemName + "'");

                if (PULL_LAB_BARCODE_LEFT_LIST_EL) {
                    var leftRow = document.createElement("div");
                    leftRow.setAttribute("data-barcode-left-item", "");
                    leftRow.style.padding = "6px 10px";
                    leftRow.style.color = "#ddd";
                    leftRow.style.borderBottom = "1px solid #2a2a2a";
                    leftRow.style.overflow = "hidden";
                    var indexSpan = document.createElement("span");
                    indexSpan.style.color = "#666";
                    indexSpan.style.marginRight = "8px";
                    indexSpan.style.fontSize = "12px";
                    indexSpan.textContent = String(result.length) + ".";
                    var nameSpan = document.createElement("span");
                    nameSpan.style.fontSize = "12px";
                    nameSpan.style.fontWeight = "500";
                    nameSpan.textContent = itemName;
                    var barcodeDiv = document.createElement("div");
                    barcodeDiv.style.fontSize = "10px";
                    barcodeDiv.style.color = "#888";
                    barcodeDiv.style.marginTop = "2px";
                    barcodeDiv.style.marginLeft = "20px";
                    barcodeDiv.style.overflow = "hidden";
                    barcodeDiv.style.textOverflow = "ellipsis";
                    barcodeDiv.style.whiteSpace = "nowrap";
                    barcodeDiv.textContent = barcode;
                    leftRow.appendChild(indexSpan);
                    leftRow.appendChild(nameSpan);
                    leftRow.appendChild(barcodeDiv);
                    PULL_LAB_BARCODE_LEFT_LIST_EL.appendChild(leftRow);
                }

                ci = ci + 1;
            }

            if (result.length > 0) {
                break;
            }

            attempts = attempts + 1;
            if (attempts < maxAttempts) {
                log("[PullLabBarcode] collectBarcodeIcons: 0 qualifying icons, retrying");
                await sleep(BARCODE_TIMEOUTS.waitCollectIconsMs / maxAttempts);
            }
        }

        log("[PullLabBarcode] collectBarcodeIcons: collected " + String(result.length) + " icons to process");
        return result;
    }

    async function openBarcodeModalAndFill(iconObj) {
        log("[PullLabBarcode] openBarcodeModalAndFill: starting for barcode='" + iconObj.barcode + "'");
        var attempts = 0;
        var maxAttempts = BARCODE_RETRY.fillConfirmRetries + 1;

        while (attempts < maxAttempts) {
            if (PULL_LAB_BARCODE_STOPPED) {
                log("[PullLabBarcode] openBarcodeModalAndFill: stopped");
                return false;
            }

            attempts = attempts + 1;
            log("[PullLabBarcode] openBarcodeModalAndFill: attempt " + String(attempts) + "/" + String(maxAttempts));

            var existingModal = document.querySelector(BARCODE_SELECTORS.bootboxVisibleModal);
            if (existingModal) {
                log("[PullLabBarcode] openBarcodeModalAndFill: existing modal detected, waiting for it to close");
                await waitUntilHidden(BARCODE_SELECTORS.bootboxVisibleModal, BARCODE_TIMEOUTS.waitModalCloseMs);
                await sleep(BARCODE_TIMEOUTS.waitSettleMs);
            }

            var currentIcon = iconObj.el;
            if (!currentIcon || !document.body.contains(currentIcon)) {
                log("[PullLabBarcode] openBarcodeModalAndFill: icon detached, re-querying by barcode");
                var reQueryIcons = document.querySelectorAll(BARCODE_SELECTORS.barcodeIcon);
                var reFound = false;
                var qi = 0;
                while (qi < reQueryIcons.length) {
                    var qVal = extractBarcodeValue(reQueryIcons[qi]);
                    if (qVal === iconObj.barcode) {
                        iconObj.el = reQueryIcons[qi];
                        currentIcon = reQueryIcons[qi];
                        reFound = true;
                        log("[PullLabBarcode] openBarcodeModalAndFill: re-found icon in DOM");
                        break;
                    }
                    qi = qi + 1;
                }
                if (!reFound) {
                    log("[PullLabBarcode] openBarcodeModalAndFill: icon not found in DOM after re-query");
                    return false;
                }
            }

            var preVerifiedClass = currentIcon.classList.contains(BARCODE_SELECTORS.verifiedClass);
            var preTooltip = currentIcon.getAttribute(BARCODE_SELECTORS.tooltipAttr) || "";
            if (preVerifiedClass || BARCODE_REGEX.verifiedPrefix.test(preTooltip)) {
                log("[PullLabBarcode] openBarcodeModalAndFill: icon already verified before clicking");
                return "already_verified";
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: clicking icon");
            try {
                currentIcon.click();
            } catch (clickErr) {
                log("[PullLabBarcode] openBarcodeModalAndFill: click error " + String(clickErr));
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: waiting for modal input");
            var modalInput = await waitForSelector(BARCODE_SELECTORS.modalInput, BARCODE_TIMEOUTS.waitModalOpenMs);
            if (!modalInput) {
                log("[PullLabBarcode] openBarcodeModalAndFill: modal input not found after waiting");
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            var visibleModal = document.querySelector(BARCODE_SELECTORS.bootboxVisibleModal);
            if (!visibleModal) {
                log("[PullLabBarcode] openBarcodeModalAndFill: modal container not visible");
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: filling input with '" + iconObj.barcode + "'");
            try {
                modalInput.focus();
                modalInput.value = iconObj.barcode;
                modalInput.dispatchEvent(new Event("input", { bubbles: true }));
                modalInput.dispatchEvent(new Event("change", { bubbles: true }));
            } catch (fillErr) {
                log("[PullLabBarcode] openBarcodeModalAndFill: fill error " + String(fillErr));
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            await sleep(BARCODE_TIMEOUTS.waitSettleMs);

            log("[PullLabBarcode] openBarcodeModalAndFill: looking for confirm button");
            var confirmBtn = document.querySelector(BARCODE_SELECTORS.modalConfirmPrimary);
            if (!confirmBtn) {
                log("[PullLabBarcode] openBarcodeModalAndFill: primary confirm not found, trying fallback");
                var allPrimaryBtns = document.querySelectorAll(BARCODE_SELECTORS.modalAnyPrimary);
                var bi = 0;
                while (bi < allPrimaryBtns.length) {
                    var btnText = (allPrimaryBtns[bi].textContent || "").trim();
                    if (BARCODE_REGEX.okText.test(btnText)) {
                        confirmBtn = allPrimaryBtns[bi];
                        log("[PullLabBarcode] openBarcodeModalAndFill: found fallback confirm with text '" + btnText + "'");
                        break;
                    }
                    bi = bi + 1;
                }
            }

            if (!confirmBtn) {
                log("[PullLabBarcode] openBarcodeModalAndFill: no confirm button found");
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: clicking confirm");
            try {
                confirmBtn.click();
            } catch (confirmErr) {
                log("[PullLabBarcode] openBarcodeModalAndFill: confirm click error " + String(confirmErr));
                if (attempts < maxAttempts) {
                    await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    continue;
                }
                return false;
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: waiting for modal to close");
            var closed = await waitUntilHidden(BARCODE_SELECTORS.bootboxVisibleModal, BARCODE_TIMEOUTS.waitModalCloseMs);
            if (!closed) {
                log("[PullLabBarcode] openBarcodeModalAndFill: modal did not close in time");
                if (attempts < maxAttempts) {
                    continue;
                }
                return false;
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: modal closed, checking for error modal");
            var errorModal = await waitForSelector(BARCODE_SELECTORS.bootboxVisibleModal, 1500);
            if (errorModal) {
                var errorBody = errorModal.querySelector(BARCODE_SELECTORS.bootboxBody);
                var errorText = errorBody ? (errorBody.textContent || "").trim() : "";
                log("[PullLabBarcode] openBarcodeModalAndFill: post-confirm modal detected, text='" + errorText + "'");
                if (errorText.toLowerCase().indexOf("invalid") !== -1) {
                    log("[PullLabBarcode] openBarcodeModalAndFill: invalid barcode error detected");
                    var okBtn = errorModal.querySelector(BARCODE_SELECTORS.bootboxOkBtn);
                    if (!okBtn) {
                        var fallbackBtns = errorModal.querySelectorAll(".modal-footer .btn.btn-primary");
                        var fbi = 0;
                        while (fbi < fallbackBtns.length) {
                            var fbText = (fallbackBtns[fbi].textContent || "").trim();
                            if (BARCODE_REGEX.okText.test(fbText)) {
                                okBtn = fallbackBtns[fbi];
                                break;
                            }
                            fbi = fbi + 1;
                        }
                    }
                    if (okBtn) {
                        log("[PullLabBarcode] openBarcodeModalAndFill: clicking OK to dismiss invalid barcode modal");
                        try {
                            okBtn.click();
                        } catch (okErr) {
                            log("[PullLabBarcode] openBarcodeModalAndFill: OK click error " + String(okErr));
                        }
                        await waitUntilHidden(BARCODE_SELECTORS.bootboxVisibleModal, BARCODE_TIMEOUTS.waitModalCloseMs);
                        await sleep(BARCODE_TIMEOUTS.waitSettleMs);
                    }
                    return "invalid_barcode";
                }
            }

            log("[PullLabBarcode] openBarcodeModalAndFill: modal closed successfully");
            return true;
        }

        log("[PullLabBarcode] openBarcodeModalAndFill: all attempts exhausted");
        return false;
    }

    async function verifyIconUpdated(iconObj) {
        log("[PullLabBarcode] verifyIconUpdated: checking verification for barcode='" + iconObj.barcode + "'");
        var attempts = 0;
        var maxAttempts = BARCODE_RETRY.verifyRetries + 1;

        while (attempts < maxAttempts) {
            if (PULL_LAB_BARCODE_STOPPED) {
                log("[PullLabBarcode] verifyIconUpdated: stopped");
                return false;
            }

            attempts = attempts + 1;
            await sleep(BARCODE_TIMEOUTS.waitVerifyIconMs);

            var currentIcon = iconObj.el;
            if (!currentIcon || !document.body.contains(currentIcon)) {
                log("[PullLabBarcode] verifyIconUpdated: icon detached, re-querying");
                var allIcons = document.querySelectorAll(BARCODE_SELECTORS.barcodeIcon);
                var found = false;
                var qi = 0;
                while (qi < allIcons.length) {
                    var qVal = extractBarcodeValue(allIcons[qi]);
                    if (qVal === iconObj.barcode) {
                        iconObj.el = allIcons[qi];
                        currentIcon = allIcons[qi];
                        found = true;
                        break;
                    }
                    qi = qi + 1;
                }
                if (!found) {
                    var broadIcons = document.querySelectorAll(BARCODE_SELECTORS.presenceCheckAnyIcon);
                    var bqi = 0;
                    while (bqi < broadIcons.length) {
                        var broadVal = extractBarcodeValue(broadIcons[bqi]);
                        if (broadVal === iconObj.barcode) {
                            iconObj.el = broadIcons[bqi];
                            currentIcon = broadIcons[bqi];
                            found = true;
                            break;
                        }
                        bqi = bqi + 1;
                    }
                }
                if (!found) {
                    log("[PullLabBarcode] verifyIconUpdated: icon not found in DOM");
                    if (attempts < maxAttempts) {
                        continue;
                    }
                    return false;
                }
            }

            var hasClass = currentIcon.classList.contains(BARCODE_SELECTORS.verifiedClass);
            var vTooltip = currentIcon.getAttribute(BARCODE_SELECTORS.tooltipAttr) || "";
            var tooltipVerified = BARCODE_REGEX.verifiedPrefix.test(vTooltip);

            if (hasClass || tooltipVerified) {
                log("[PullLabBarcode] verifyIconUpdated: icon verified (class=" + String(hasClass) + ", tooltip=" + String(tooltipVerified) + ")");
                return true;
            }

            log("[PullLabBarcode] verifyIconUpdated: not yet verified, attempt " + String(attempts) + "/" + String(maxAttempts));
        }

        log("[PullLabBarcode] verifyIconUpdated: verification failed after all retries");
        return false;
    }

    async function processIconsSequentially(iconList) {
        log("[PullLabBarcode] processIconsSequentially: starting with " + String(iconList.length) + " icons");

        setAriaBusyOn();
        announceBarcodeAriaLive("Starting barcode processing for " + String(iconList.length) + " items");

        BARCODE_COUNTERS.total = iconList.length;
        BARCODE_COUNTERS.processed = 0;
        BARCODE_COUNTERS.verified = 0;
        BARCODE_COUNTERS.skipped = 0;
        BARCODE_COUNTERS.failures = 0;
        BARCODE_COUNTERS.pending = iconList.length;
        updateBarcodeRightPanelSummary(BARCODE_COUNTERS);

        var startTime = Date.now();
        var idx = 0;

        while (idx < iconList.length) {
            if (PULL_LAB_BARCODE_STOPPED) {
                log("[PullLabBarcode] processIconsSequentially: stopped at index " + String(idx));
                updateBarcodeRightPanelStatus(iconList[idx].statusKey, BARCODE_LABELS.statusStopped);
                break;
            }

            if (Date.now() - startTime > BARCODE_TIMEOUTS.maxTotalDurationMs) {
                log("[PullLabBarcode] processIconsSequentially: max duration exceeded");
                updateBarcodeRightPanelStatus(iconList[idx].statusKey, BARCODE_LABELS.statusStopped, "timeout");
                break;
            }

            var iconObj = iconList[idx];
            log("[PullLabBarcode] processIconsSequentially: processing " + String(idx + 1) + "/" + String(iconList.length) + " barcode='" + iconObj.barcode + "'");

            updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusProcessing);

            var currentIcon = iconObj.el;
            if (currentIcon && document.body.contains(currentIcon)) {
                var alreadyVerifiedClass = currentIcon.classList.contains(BARCODE_SELECTORS.verifiedClass);
                var alreadyTooltip = currentIcon.getAttribute(BARCODE_SELECTORS.tooltipAttr) || "";
                var alreadyVerifiedTooltip = BARCODE_REGEX.verifiedPrefix.test(alreadyTooltip);
                if (alreadyVerifiedClass || alreadyVerifiedTooltip) {
                    log("[PullLabBarcode] processIconsSequentially: icon already verified, skipping");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusSkippedVerified);
                    BARCODE_COUNTERS.skipped = BARCODE_COUNTERS.skipped + 1;
                    BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                    updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
                    await barcodeYieldToUI();
                    idx = idx + 1;
                    continue;
                }
            } else {
                log("[PullLabBarcode] processIconsSequentially: icon detached, re-querying");
                var reIcons = document.querySelectorAll(BARCODE_SELECTORS.barcodeIcon);
                var reFound = false;
                var ri = 0;
                while (ri < reIcons.length) {
                    var rVal = extractBarcodeValue(reIcons[ri]);
                    if (rVal === iconObj.barcode) {
                        iconObj.el = reIcons[ri];
                        reFound = true;
                        log("[PullLabBarcode] processIconsSequentially: re-found icon");
                        break;
                    }
                    ri = ri + 1;
                }
                if (!reFound) {
                    log("[PullLabBarcode] processIconsSequentially: icon not found, marking failed");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusFailed, "detached");
                    BARCODE_COUNTERS.failures = BARCODE_COUNTERS.failures + 1;
                    BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                    updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
                    await barcodeYieldToUI();
                    idx = idx + 1;
                    continue;
                }
            }

            try {
                var fillResult = await openBarcodeModalAndFill(iconObj);

                if (PULL_LAB_BARCODE_STOPPED) {
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusStopped);
                    break;
                }

                if (fillResult === "already_verified") {
                    log("[PullLabBarcode] processIconsSequentially: icon was already verified");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusSkippedVerified);
                    BARCODE_COUNTERS.skipped = BARCODE_COUNTERS.skipped + 1;
                    BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                    updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
                    await barcodeYieldToUI();
                    idx = idx + 1;
                    continue;
                }

                if (fillResult === "invalid_barcode") {
                    log("[PullLabBarcode] processIconsSequentially: invalid barcode for '" + iconObj.barcode + "'");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusInvalidBarcode);
                    BARCODE_COUNTERS.failures = BARCODE_COUNTERS.failures + 1;
                    BARCODE_COUNTERS.processed = BARCODE_COUNTERS.processed + 1;
                    BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                    updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
                    await barcodeYieldToUI();
                    idx = idx + 1;
                    continue;
                }

                if (!fillResult) {
                    log("[PullLabBarcode] processIconsSequentially: fill/confirm failed for barcode='" + iconObj.barcode + "'");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusFailed);
                    BARCODE_COUNTERS.failures = BARCODE_COUNTERS.failures + 1;
                    BARCODE_COUNTERS.processed = BARCODE_COUNTERS.processed + 1;
                    BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                    updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
                    await barcodeYieldToUI();
                    idx = idx + 1;
                    continue;
                }

                var verified = await verifyIconUpdated(iconObj);

                if (PULL_LAB_BARCODE_STOPPED) {
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusStopped);
                    break;
                }

                if (verified) {
                    log("[PullLabBarcode] processIconsSequentially: verified barcode='" + iconObj.barcode + "'");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusVerified);
                    BARCODE_COUNTERS.verified = BARCODE_COUNTERS.verified + 1;
                } else {
                    log("[PullLabBarcode] processIconsSequentially: verification failed for barcode='" + iconObj.barcode + "'");
                    updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusFailed, "verify failed");
                    BARCODE_COUNTERS.failures = BARCODE_COUNTERS.failures + 1;
                }

                BARCODE_COUNTERS.processed = BARCODE_COUNTERS.processed + 1;
                BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
            } catch (procErr) {
                log("[PullLabBarcode] processIconsSequentially: error processing barcode='" + iconObj.barcode + "': " + String(procErr));
                updateBarcodeRightPanelStatus(iconObj.statusKey, BARCODE_LABELS.statusFailed, String(procErr));
                BARCODE_COUNTERS.failures = BARCODE_COUNTERS.failures + 1;
                BARCODE_COUNTERS.processed = BARCODE_COUNTERS.processed + 1;
                BARCODE_COUNTERS.pending = BARCODE_COUNTERS.pending - 1;
                updateBarcodeRightPanelSummary(BARCODE_COUNTERS);
            }

            await barcodeYieldToUI();
            idx = idx + 1;
        }

        setAriaBusyOff();
        PULL_LAB_BARCODE_RUNNING = false;

        var completionMsg = BARCODE_LABELS.done + ": " + String(BARCODE_COUNTERS.verified) + " verified, " + String(BARCODE_COUNTERS.skipped) + " skipped, " + String(BARCODE_COUNTERS.failures) + " failed out of " + String(BARCODE_COUNTERS.total);
        log("[PullLabBarcode] processIconsSequentially: " + completionMsg);
        announceBarcodeAriaLive(completionMsg);
        updateBarcodeRightPanelSummary(BARCODE_COUNTERS);

        if (PULL_LAB_BARCODE_BUTTON_REF) {
            try {
                PULL_LAB_BARCODE_BUTTON_REF.focus();
            } catch (focusErr) {}
        }
    }

    async function pullLabBarcodeInit() {
        log("[PullLabBarcode] pullLabBarcodeInit: starting feature");

        if (PULL_LAB_BARCODE_RUNNING) {
            log("[PullLabBarcode] pullLabBarcodeInit: already running, ignoring");
            return;
        }

        PULL_LAB_BARCODE_STOPPED = false;
        PULL_LAB_BARCODE_RUNNING = true;
        PULL_LAB_BARCODE_STATUS_MAP = {};

        BARCODE_COUNTERS.total = 0;
        BARCODE_COUNTERS.processed = 0;
        BARCODE_COUNTERS.verified = 0;
        BARCODE_COUNTERS.skipped = 0;
        BARCODE_COUNTERS.failures = 0;
        BARCODE_COUNTERS.pending = 0;

        openBarcodeProgressPanel();

        if (PULL_LAB_BARCODE_STOPPED) {
            log("[PullLabBarcode] pullLabBarcodeInit: stopped before collection");
            return;
        }

        announceBarcodeAriaLive(BARCODE_LABELS.scanning);

        var iconList = await collectBarcodeIcons();

        if (PULL_LAB_BARCODE_STOPPED) {
            log("[PullLabBarcode] pullLabBarcodeInit: stopped after collection");
            return;
        }

        if (iconList.length === 0) {
            log("[PullLabBarcode] pullLabBarcodeInit: no icons found");
            if (PULL_LAB_BARCODE_SUMMARY_EL) {
                PULL_LAB_BARCODE_SUMMARY_EL.innerHTML = "";
                var warningSpan = document.createElement("span");
                warningSpan.style.color = "#f0ad4e";
                warningSpan.style.fontWeight = "600";
                warningSpan.textContent = "No barcode icons requiring input were found on this page.";
                PULL_LAB_BARCODE_SUMMARY_EL.appendChild(warningSpan);
            }
            announceBarcodeAriaLive("No barcode icons found");
            PULL_LAB_BARCODE_RUNNING = false;
            setAriaBusyOff();
            return;
        }

        var ri = 0;
        while (ri < iconList.length) {
            updateBarcodeRightPanelStatus(iconList[ri].statusKey, BARCODE_LABELS.statusPending);
            ri = ri + 1;
        }

        await processIconsSequentially(iconList);

        log("[PullLabBarcode] pullLabBarcodeInit: feature completed");
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
            "Pull Barcode",
            "Pull Lab Barcode",
            "Auto-Refresh",
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
            return false;
        });

        // Prevent typing but allow key capture for hotkey capture
        hotkeyInput.addEventListener("keypress", function(e) {
            e.preventDefault();
            return false;
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

    //========================================
    // UI Scaling
    //========================================
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


    // ======================
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

    //==========================
    // RUN BARCODE FEATURE
    //==========================
    // This section contains all functions related to barcode lookup and data entry.
    // This feature automates finding subject barcodes by searching through epochs,
    // cohorts, and subjects. Functions handle subject identification, barcode retrieval,
    // and populating barcode input modals.
    //==========================


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

                    // Wait for epoch select to appear and be populated (Edge needs more time than Chrome)
                    var epochSel = null;
                    var retryCount = 0;
                    var maxRetries = 80;
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

        log("APS_RunBarcode: Fetching barcode in background\u2026");

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
            w = scale(PANEL_DEFAULT_WIDTH);
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

    function updatePanelCollapsedState(panel, bodyContainer, resizeHandle, collapseBtn, headerBar) {
        var collapsed = getPanelCollapsed();
        if (collapsed) {
            panel.style.width = scale(PANEL_DEFAULT_WIDTH);
            panel.style.height = scale(PANEL_HEADER_HEIGHT_PX + 20);
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
                collapseBtn.textContent = "\u2212";
            }
        }
    }

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
        headerBar.style.minHeight = String(PANEL_HEADER_HEIGHT_PX) + "px";
        headerBar.style.boxSizing = "border-box";
        headerBar.style.padding = "8px 12px";
        headerBar.style.flexShrink = "0";
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
        closeBtn.textContent = "\u2715";
        closeBtn.style.background = "transparent";
        closeBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        closeBtn.style.border = "none";
        closeBtn.style.cursor = "pointer";
        closeBtn.style.fontSize = "18px";
        closeBtn.style.lineHeight = "1";
        closeBtn.style.padding = "4px 8px";
        closeBtn.style.borderRadius = "4px";
        closeBtn.style.width = "28px";
        closeBtn.style.height = "28px";
        closeBtn.style.display = "flex";
        closeBtn.style.alignItems = "center";
        closeBtn.style.justifyContent = "center";
        closeBtn.style.flexShrink = "0";
        closeBtn.style.overflow = "hidden";
        closeBtn.style.boxSizing = "border-box";
        closeBtn.addEventListener("mouseenter", function() {
            closeBtn.style.background = glass ? THEME_SURFACE_BG_HEAVY : "#333";
        });
        closeBtn.addEventListener("mouseleave", function() {
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

        return {
            element: popup,
            close: function () {
                if (onClose) {
                    try {
                        onClose();
                    } catch (e) {
                        log("Popup onClose error: " + String(e));
                    }
                }
                document.removeEventListener("keydown", keyHandler, true);
                popup.remove();
            },
            setContent: function (newContent) {
                bodyContainer.innerHTML = "";
                if (typeof newContent === "string") {
                    bodyContainer.innerHTML = newContent;
                } else if (newContent && newContent.nodeType === 1) {
                    bodyContainer.appendChild(newContent);
                }
            }
        };
    }

    function showWrongPagePopup(featureName, requiredPage, currentPath) {
        // Create warning popup content
        var warningContainer = document.createElement("div");
        warningContainer.style.padding = "20px";
        warningContainer.style.textAlign = "center";

        var warningIcon = document.createElement("div");
        warningIcon.innerHTML = "\u26A0\uFE0F";
        warningIcon.style.fontSize = "48px";
        warningIcon.style.marginBottom = "15px";

        var warningTitle = document.createElement("div");
        warningTitle.textContent = "Wrong Page Detected";
        warningTitle.style.fontSize = "18px";
        warningTitle.style.fontWeight = "bold";
        warningTitle.style.marginBottom = "10px";
        warningTitle.style.color = "#ff6b6b";

        var warningMessage = document.createElement("div");
        warningMessage.textContent = "You must be on the " + featureName + " page to use this feature.";
        warningMessage.style.marginBottom = "8px";
        warningMessage.style.lineHeight = "1.4";

        var currentPage = document.createElement("div");
        currentPage.textContent = "Current page: " + currentPath;
        currentPage.style.fontSize = "12px";
        currentPage.style.color = "#999";
        currentPage.style.marginBottom = "8px";

        var correctPage = document.createElement("div");
        correctPage.textContent = "Required page: " + requiredPage;
        correctPage.style.fontSize = "12px";
        correctPage.style.color = "#999";
        correctPage.style.marginBottom = "20px";

        var okButton = document.createElement("button");
        okButton.textContent = "OK";
        okButton.style.background = "#5b43c7";
        okButton.style.color = "#fff";
        okButton.style.border = "none";
        okButton.style.borderRadius = "4px";
        okButton.style.padding = "8px 24px";
        okButton.style.cursor = "pointer";
        okButton.style.fontSize = "14px";
        okButton.onmouseenter = function() { this.style.background = "#4a35a6"; };
        okButton.onmouseleave = function() { this.style.background = "#5b43c7"; };

        warningContainer.appendChild(warningIcon);
        warningContainer.appendChild(warningTitle);
        warningContainer.appendChild(warningMessage);
        warningContainer.appendChild(currentPage);
        warningContainer.appendChild(correctPage);
        warningContainer.appendChild(okButton);

        var warningPopup = createPopup({
            title: featureName + " - Page Warning",
            content: warningContainer,
            width: "400px",
            height: "auto",
            onClose: function () {
                log(featureName + ": warning popup closed");
            }
        });

        okButton.addEventListener("click", function () {
            warningPopup.close();
            log(featureName + ": warning popup acknowledged by user");
        });

        return warningPopup;
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
        panel.style.width = savedSize.width;
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
        headerBar.style.gap = scale(PANEL_HEADER_GAP_PX);
        headerBar.style.height = scale(PANEL_HEADER_HEIGHT_PX);
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
        settingsBtn.textContent = "\u2699";
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
        bodyContainer.style.height = "calc(100% - " + scale(PANEL_HEADER_HEIGHT_PX) + ")";
        bodyContainer.style.maxHeight = "calc(100% - " + scale(PANEL_HEADER_HEIGHT_PX) + ")";
        bodyContainer.style.overflowY = "auto";
        bodyContainer.style.boxSizing = "border-box";
        var btnRow = document.createElement("div");
        btnRow.style.display = "grid";
        btnRow.style.gridTemplateColumns = "1fr 1fr";
        btnRow.style.gap = scale(BUTTON_GAP_PX);
        btnRowRef = btnRow;
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
        runBarcodeBtn.onmouseenter = function() { this.style.background = "#4a37a0"; };
        runBarcodeBtn.onmouseleave = function() { this.style.background = "#5b43c7"; };

        var autoRefreshBtn = document.createElement("button");
        autoRefreshBtn.textContent = "Auto-Refresh";
        autoRefreshBtn.style.background = "#4a90e2";
        autoRefreshBtn.style.color = "#fff";
        autoRefreshBtn.style.border = "none";
        autoRefreshBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        autoRefreshBtn.style.padding = scale(BUTTON_PADDING_PX);
        autoRefreshBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        autoRefreshBtn.style.cursor = "pointer";
        autoRefreshBtn.style.fontWeight = "500";
        autoRefreshBtn.style.transition = "background 0.2s";
        autoRefreshBtn.onmouseenter = function() { this.style.background = "#58a1f5"; };
        autoRefreshBtn.onmouseleave = function() { this.style.background = "#4a90e2"; };
        autoRefreshBtn.addEventListener("click", function() {
            log("[AutoRefresh] Button clicked");
            openAutoRefreshPopup();
        });

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
        toggleLogsBtn.onmouseenter = function() { this.style.background = "#5a6268"; };
        toggleLogsBtn.onmouseleave = function() { this.style.background = "#6c757d"; };

        var pullLabBarcodeBtn = document.createElement("button");
        pullLabBarcodeBtn.textContent = BARCODE_LABELS.featureButton;
        pullLabBarcodeBtn.style.background = "#5b43c7";
        pullLabBarcodeBtn.style.color = "#fff";
        pullLabBarcodeBtn.style.border = "none";
        pullLabBarcodeBtn.style.borderRadius = scale(BUTTON_BORDER_RADIUS_PX);
        pullLabBarcodeBtn.style.padding = scale(BUTTON_PADDING_PX);
        pullLabBarcodeBtn.style.fontSize = scale(PANEL_FONT_SIZE_PX);
        pullLabBarcodeBtn.style.cursor = "pointer";
        pullLabBarcodeBtn.setAttribute("aria-label", BARCODE_LABELS.featureButton);
        pullLabBarcodeBtn.onmouseenter = function () {
            this.style.background = "#4a37a0";
        };
        pullLabBarcodeBtn.onmouseleave = function () {
            this.style.background = "#5b43c7";
        };
        pullLabBarcodeBtn.addEventListener("click", function () {
            log("[PullLabBarcode] Button clicked");
            pullLabBarcodeInit();
        });
        PULL_LAB_BARCODE_BUTTON_REF = pullLabBarcodeBtn;

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

        // Apply glassmorphism theme to all panel buttons if glass theme is active
        if (glass) {
            var allPanelBtns = [runBarcodeBtn, pullLabBarcodeBtn, autoRefreshBtn, clearLogsBtn, toggleLogsBtn];
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
            { el: runBarcodeBtn, label: "Pull Barcode" },
            { el: pullLabBarcodeBtn, label: "Pull Lab Barcode" },
            { el: autoRefreshBtn, label: "Auto-Refresh" },
            { el: clearLogsBtn, label: "Clear Logs" },
            { el: toggleLogsBtn, label: "Hide Logs" }
        ];

        for (var bi = 0; bi < panelButtons.length; bi++) {
            var btnItem = panelButtons[bi];
            if (isButtonVisible(btnItem.label)) {
                btnRow.appendChild(btnItem.el);
            }
        }

        bodyContainer.appendChild(btnRow);
        var status = document.createElement("div");
        status.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        status.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        status.style.border = glass ? ("1px solid " + THEME_SURFACE_INNER_BORDER) : "1px solid #333";
        status.style.borderRadius = scale(STATUS_BORDER_RADIUS_PX);
        status.style.padding = scale(STATUS_PADDING_PX);
        status.style.fontSize = scale(STATUS_FONT_SIZE_PX);
        status.style.whiteSpace = "pre-wrap";
        status.textContent = "Ready";
        bodyContainer.appendChild(status);

        // UI Scale Control
        var scaleControl = document.createElement("div");
        scaleControl.style.marginTop = scale(STATUS_MARGIN_TOP_PX);
        scaleControl.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        scaleControl.style.border = "1px solid #333";
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
        logBox.style.border = "1px solid #333";
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
        clearLogsBtn.addEventListener("click", function () {
            clearLogs();
            status.textContent = "Logs cleared";
        });
        runBarcodeBtn.addEventListener("click", async function () {
            log("Run Barcode: button clicked");
            await APS_RunBarcode();
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
        var minRight = -pw / 2;
        var maxRight = vw - pw / 2;
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

    //==========================
    // AUTO-REFRESH FEATURE
    //==========================

    function openAutoRefreshPopup() {
        var glass = isGlassTheme();

        var container = document.createElement("div");
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "16px";

        var instruction = document.createElement("div");
        instruction.style.fontSize = "14px";
        instruction.style.color = glass ? THEME_TEXT_MUTED : "#ccc";
        instruction.textContent = "Enter the refresh interval in minutes (1 - 100):";
        container.appendChild(instruction);

        var inputRow = document.createElement("div");
        inputRow.style.display = "flex";
        inputRow.style.gap = "10px";
        inputRow.style.alignItems = "center";

        var input = document.createElement("input");
        input.type = "number";
        input.min = "1";
        input.max = "100";
        input.step = "1";
        input.placeholder = "e.g. 5";
        input.style.flex = "1";
        input.style.padding = "10px 12px";
        input.style.borderRadius = "6px";
        input.style.border = glass ? ("1px solid " + THEME_SURFACE_BORDER) : "1px solid #444";
        input.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        input.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        input.style.fontSize = "16px";
        input.style.outline = "none";
        input.style.fontFamily = "monospace";
        if (glass) input.className = "ie-input";

        var unitLabel = document.createElement("span");
        unitLabel.textContent = "minutes";
        unitLabel.style.fontSize = "14px";
        unitLabel.style.color = glass ? THEME_TEXT_MUTED : "#aaa";

        inputRow.appendChild(input);
        inputRow.appendChild(unitLabel);
        container.appendChild(inputRow);

        var errorText = document.createElement("div");
        errorText.style.color = glass ? THEME_DANGER : "#ef4444";
        errorText.style.fontSize = "12px";
        errorText.style.minHeight = "16px";
        container.appendChild(errorText);

        var btnRow = document.createElement("div");
        btnRow.style.display = "flex";
        btnRow.style.gap = "10px";
        btnRow.style.justifyContent = "flex-end";

        var cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Cancel";
        cancelBtn.style.padding = "10px 20px";
        cancelBtn.style.borderRadius = "6px";
        cancelBtn.style.border = "none";
        cancelBtn.style.background = glass ? THEME_SURFACE_BG : "#333";
        cancelBtn.style.color = glass ? THEME_TEXT_PRIMARY : "#fff";
        cancelBtn.style.cursor = "pointer";
        cancelBtn.style.fontSize = "14px";
        if (glass) cancelBtn.className = "ie-btn-secondary";

        var confirmBtn = document.createElement("button");
        confirmBtn.textContent = "Confirm";
        confirmBtn.disabled = true;
        confirmBtn.style.opacity = "0.5";
        confirmBtn.style.padding = "10px 20px";
        confirmBtn.style.borderRadius = "6px";
        confirmBtn.style.border = "none";
        confirmBtn.style.background = glass ? "" : "#5b43c7";
        confirmBtn.style.color = glass ? "" : "#fff";
        confirmBtn.style.cursor = "not-allowed";
        confirmBtn.style.fontSize = "14px";
        confirmBtn.style.fontWeight = "600";
        if (glass) confirmBtn.className = "ie-btn-primary";

        function validateInput() {
            var val = input.value.trim();
            var num = parseInt(val, 10);
            if (val.length === 0) {
                errorText.textContent = "";
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = "0.5";
                confirmBtn.style.cursor = "not-allowed";
                return;
            }
            if (isNaN(num) || num < 1 || num > 100 || String(num) !== val) {
                errorText.textContent = "Please enter a whole number between 1 and 100.";
                confirmBtn.disabled = true;
                confirmBtn.style.opacity = "0.5";
                confirmBtn.style.cursor = "not-allowed";
                return;
            }
            errorText.textContent = "";
            confirmBtn.disabled = false;
            confirmBtn.style.opacity = "1";
            confirmBtn.style.cursor = "pointer";
        }

        input.addEventListener("input", validateInput);
        input.addEventListener("keydown", function(e) {
            if (e.key === "Enter" && !confirmBtn.disabled) {
                confirmBtn.click();
            }
        });

        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(confirmBtn);
        container.appendChild(btnRow);

        var popup = createPopup({
            title: "Auto-Refresh",
            description: "Set automatic page refresh interval",
            content: container,
            width: "380px",
            height: "auto"
        });

        setTimeout(function() { try { input.focus(); } catch(e) {} }, 50);

        cancelBtn.addEventListener("click", function() {
            popup.close();
        });

        confirmBtn.addEventListener("click", function() {
            var minutes = parseInt(input.value.trim(), 10);
            popup.close();
            autoRefreshStart(minutes);
        });
    }

    function autoRefreshStart(intervalMinutes) {
        var now = Date.now();
        var nextTime = now + (intervalMinutes * 60 * 1000);
        try {
            localStorage.setItem(STORAGE_AUTO_REFRESH_ACTIVE, "1");
            localStorage.setItem(STORAGE_AUTO_REFRESH_INTERVAL, String(intervalMinutes));
            localStorage.setItem(STORAGE_AUTO_REFRESH_COUNT, "0");
            localStorage.setItem(STORAGE_AUTO_REFRESH_NEXT_TIME, String(nextTime));
        } catch (e) {}
        log("[AutoRefresh] Started: interval=" + intervalMinutes + "min, nextRefresh=" + new Date(nextTime).toISOString());
        showAutoRefreshStatusPanel();
    }

    function autoRefreshStop() {
        try {
            localStorage.removeItem(STORAGE_AUTO_REFRESH_ACTIVE);
            localStorage.removeItem(STORAGE_AUTO_REFRESH_INTERVAL);
            localStorage.removeItem(STORAGE_AUTO_REFRESH_COUNT);
            localStorage.removeItem(STORAGE_AUTO_REFRESH_NEXT_TIME);
        } catch (e) {}
        if (AUTO_REFRESH_TIMER_ID) {
            clearInterval(AUTO_REFRESH_TIMER_ID);
            AUTO_REFRESH_TIMER_ID = null;
        }
        var existing = document.getElementById(AUTO_REFRESH_PANEL_ID);
        if (existing) existing.remove();
        log("[AutoRefresh] Stopped by user");
    }

    function autoRefreshRestore() {
        var active = null;
        try { active = localStorage.getItem(STORAGE_AUTO_REFRESH_ACTIVE); } catch(e) {}
        if (active !== "1") return;

        var intervalStr = null;
        var countStr = null;
        var nextTimeStr = null;
        try {
            intervalStr = localStorage.getItem(STORAGE_AUTO_REFRESH_INTERVAL);
            countStr = localStorage.getItem(STORAGE_AUTO_REFRESH_COUNT);
            nextTimeStr = localStorage.getItem(STORAGE_AUTO_REFRESH_NEXT_TIME);
        } catch(e) {}

        var interval = parseInt(intervalStr, 10);
        var count = parseInt(countStr, 10);
        var nextTime = parseInt(nextTimeStr, 10);

        if (isNaN(interval) || isNaN(count) || isNaN(nextTime)) {
            log("[AutoRefresh] Restore: invalid stored data, clearing");
            autoRefreshStop();
            return;
        }

        var now = Date.now();
        if (now >= nextTime) {
            count = count + 1;
            var newNextTime = now + (interval * 60 * 1000);
            try {
                localStorage.setItem(STORAGE_AUTO_REFRESH_COUNT, String(count));
                localStorage.setItem(STORAGE_AUTO_REFRESH_NEXT_TIME, String(newNextTime));
            } catch(e) {}
            log("[AutoRefresh] Restore: refresh #" + count + " occurred, next in " + interval + " min");
        }

        log("[AutoRefresh] Restore: active, interval=" + interval + "min, count=" + count);
        showAutoRefreshStatusPanel();
    }

    function showAutoRefreshStatusPanel() {
        var glass = isGlassTheme();
        if (glass) injectThemeStylesIfNeeded();

        var existing = document.getElementById(AUTO_REFRESH_PANEL_ID);
        if (existing) existing.remove();

        if (AUTO_REFRESH_TIMER_ID) {
            clearInterval(AUTO_REFRESH_TIMER_ID);
            AUTO_REFRESH_TIMER_ID = null;
        }

        var intervalStr = null;
        var countStr = null;
        var nextTimeStr = null;
        try {
            intervalStr = localStorage.getItem(STORAGE_AUTO_REFRESH_INTERVAL);
            countStr = localStorage.getItem(STORAGE_AUTO_REFRESH_COUNT);
            nextTimeStr = localStorage.getItem(STORAGE_AUTO_REFRESH_NEXT_TIME);
        } catch(e) {}

        var interval = parseInt(intervalStr, 10);
        var count = parseInt(countStr, 10);
        var nextTime = parseInt(nextTimeStr, 10);

        if (isNaN(interval) || isNaN(nextTime)) return;
        if (isNaN(count)) count = 0;

        var panel = document.createElement("div");
        panel.id = AUTO_REFRESH_PANEL_ID;
        if (glass) {
            panel.classList.add(THEME_SCOPE_CLASS);
            panel.classList.add("ie-glass-panel");
        }
        panel.style.position = "fixed";
        panel.style.bottom = "20px";
        panel.style.left = "20px";
        panel.style.zIndex = glass ? String(THEME_Z_BASE) : "999998";
        if (!glass) {
            panel.style.background = "#111";
            panel.style.color = "#fff";
            panel.style.border = "1px solid #444";
            panel.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.5)";
        }
        panel.style.borderRadius = "10px";
        panel.style.padding = "14px 18px";
        panel.style.fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, Arial";
        panel.style.fontSize = "13px";
        panel.style.minWidth = "240px";
        panel.style.display = "flex";
        panel.style.flexDirection = "column";
        panel.style.gap = "10px";

        var titleRow = document.createElement("div");
        titleRow.style.display = "flex";
        titleRow.style.justifyContent = "space-between";
        titleRow.style.alignItems = "center";

        var titleText = document.createElement("div");
        titleText.textContent = "Auto-Refresh Active";
        titleText.style.fontWeight = "600";
        titleText.style.fontSize = "14px";
        if (glass) titleText.style.color = THEME_TEXT_PRIMARY;

        var indicator = document.createElement("div");
        indicator.style.width = "8px";
        indicator.style.height = "8px";
        indicator.style.borderRadius = "50%";
        indicator.style.background = "#10b981";
        indicator.style.animation = "ieThemeSpin 2s linear infinite";
        indicator.style.boxShadow = "0 0 6px rgba(16,185,129,0.6)";

        titleRow.appendChild(titleText);
        titleRow.appendChild(indicator);
        panel.appendChild(titleRow);

        var infoContainer = document.createElement("div");
        infoContainer.style.display = "flex";
        infoContainer.style.flexDirection = "column";
        infoContainer.style.gap = "6px";
        infoContainer.style.padding = "8px 10px";
        infoContainer.style.background = glass ? THEME_SURFACE_BG : "#1a1a1a";
        infoContainer.style.borderRadius = "6px";
        infoContainer.style.border = glass ? ("1px solid " + THEME_SURFACE_BORDER) : "1px solid #333";

        var intervalInfo = document.createElement("div");
        intervalInfo.textContent = "Interval: " + interval + " min";
        intervalInfo.style.color = glass ? THEME_TEXT_MUTED : "#aaa";
        intervalInfo.style.fontSize = "12px";

        var countInfo = document.createElement("div");
        countInfo.style.fontSize = "13px";
        countInfo.style.fontWeight = "500";
        if (glass) countInfo.style.color = THEME_TEXT_PRIMARY;

        var countdownInfo = document.createElement("div");
        countdownInfo.style.fontSize = "16px";
        countdownInfo.style.fontWeight = "600";
        countdownInfo.style.color = glass ? THEME_ACCENT : "#a78bfa";
        countdownInfo.style.fontFamily = "monospace";

        infoContainer.appendChild(intervalInfo);
        infoContainer.appendChild(countInfo);
        infoContainer.appendChild(countdownInfo);
        panel.appendChild(infoContainer);

        var stopBtn = document.createElement("button");
        stopBtn.textContent = "Stop Auto-Refresh";
        stopBtn.style.padding = "8px 16px";
        stopBtn.style.borderRadius = "6px";
        stopBtn.style.border = "none";
        stopBtn.style.cursor = "pointer";
        stopBtn.style.fontSize = "13px";
        stopBtn.style.fontWeight = "600";
        if (glass) {
            stopBtn.className = "ie-btn-danger";
        } else {
            stopBtn.style.background = "#ef4444";
            stopBtn.style.color = "#fff";
            stopBtn.onmouseenter = function() { this.style.background = "#dc2626"; };
            stopBtn.onmouseleave = function() { this.style.background = "#ef4444"; };
        }
        stopBtn.addEventListener("click", function() {
            autoRefreshStop();
        });
        panel.appendChild(stopBtn);
        document.body.appendChild(panel);

        function updateDisplay() {
            var now = Date.now();
            var currentCount = 0;
            var currentNext = nextTime;
            try {
                var c = localStorage.getItem(STORAGE_AUTO_REFRESH_COUNT);
                var n = localStorage.getItem(STORAGE_AUTO_REFRESH_NEXT_TIME);
                if (c) currentCount = parseInt(c, 10) || 0;
                if (n) currentNext = parseInt(n, 10) || nextTime;
            } catch(e) {}

            countInfo.textContent = "Refreshed: " + currentCount + " time" + (currentCount !== 1 ? "s" : "");

            var remaining = currentNext - now;
            if (remaining <= 0) {
                countdownInfo.textContent = "Refreshing now...";
                var newCount = currentCount + 1;
                var newNext = Date.now() + (interval * 60 * 1000);
                try {
                    localStorage.setItem(STORAGE_AUTO_REFRESH_COUNT, String(newCount));
                    localStorage.setItem(STORAGE_AUTO_REFRESH_NEXT_TIME, String(newNext));
                } catch(e) {}
                log("[AutoRefresh] Triggering refresh #" + newCount);
                setTimeout(function() {
                    location.reload();
                }, 500);
                return;
            }

            var totalSec = Math.ceil(remaining / 1000);
            var min = Math.floor(totalSec / 60);
            var sec = totalSec % 60;
            var secStr = sec < 10 ? "0" + sec : String(sec);
            countdownInfo.textContent = "Next refresh: " + min + ":" + secStr;
        }

        updateDisplay();
        AUTO_REFRESH_TIMER_ID = setInterval(updateDisplay, 1000);
    }

    //==========================
    // INITIALIZATION
    //==========================

    function init() {
        makePanel();
        window.APS_AddButton = function (label, handler) {
            addButtonToPanel(label, handler);
        };
        bindPanelHotkeyOnce();

        // Auto-Refresh: restore state after page reload
        autoRefreshRestore();

        // Barcode subjects page handling
        var onBarcodeSubjects = isBarcodeSubjectsPage();
        if (onBarcodeSubjects) {
            processBarcodeSubjectsPage();
            return;
        }

        log("Init complete");
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();
