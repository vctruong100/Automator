# ClinSpark Test Automator

Scripts to automate basic tasks in ClinSpark.

**File:** `ClinSpark Test Automator.js`  
**Version:** 3.8.7  
**Platform:** Tampermonkey / Greasemonkey userscript  
**Target:** `https://cenexeltest.clinspark.com/*`

---

## Table of Contents

1. [Initialization & Routing](#1-initialization--routing)
2. [Lock Activity Plans](#2-lock-activity-plans)
3. [Lock Sample Paths](#3-lock-sample-paths)
4. [Update Study Status](#4-update-study-status)
5. [Add Cohort Subjects](#5-add-cohort-subjects)
6. [Run ICF Barcode (Informed Consent)](#6-run-icf-barcode-informed-consent)
7. [Run Button 1-5 (Run ALL Pipeline)](#7-run-button-1-5-run-all-pipeline)
8. [Import Cohort Subject](#8-import-cohort-subject)
9. [Pull Barcode](#9-pull-barcode)
10. [Pull Lab Barcode](#10-pull-lab-barcode)
11. [Add Existing Subject](#11-add-existing-subject)
12. [Search Methods (Methods Library)](#12-search-methods-methods-library)
13. [PLAP Builder (Build Procedure Log)](#13-plap-builder-build-procedure-log)
14. [Populate Form](#14-populate-form)
15. [Collect All](#15-collect-all)
16. [Import I/E (Import Eligibility)](#16-import-ie-import-eligibility)
17. [Clear Mapping](#17-clear-mapping)
18. [Archive/Update Forms](#18-archiveupdate-forms)
19. [Copy Activity Forms](#19-copy-activity-forms)
20. [Item Method Forms (Parse Method)](#20-item-method-forms-parse-method)
21. [Find Form](#21-find-form)
22. [Find Study Events](#22-find-study-events)
23. [Pause / Resume](#23-pause--resume)
24. [Settings](#24-settings)
25. [Shared Utilities](#25-shared-utilities)

---

## 1. Initialization & Routing

**Entry:** `init()` → called on `DOMContentLoaded` or immediately if DOM is ready.

```
init()
├── makePanel()                          // Build the floating UI panel with all buttons
│   ├── injectThemeStylesIfNeeded()      // Glass theme CSS injection
│   ├── setupResizeHandle()              // Panel resize handle
│   └── (button click handlers registered here — see each feature below)
├── bindPanelHotkeyOnce()               // F2 toggle hotkey
│   └── handler()
│       ├── keyMatchesToggle()
│       └── togglePanelHiddenViaHotkey()
├── recreatePopupsIfNeeded()            // Restore popups across page navigations
│
├── [Page Routing] based on URL:
│   ├── runMode == "addExistingSubject" → processAESOnPageLoad()
│   ├── /secure/study/data/list         → processFindFormOnList() + processFindStudyEventsOnList()
│   ├── /secure/samples/configure/paths → processLockSamplePathsPage()
│   ├── .../paths/show/...              → processLockSamplePathDetailPage()
│   ├── .../paths/update/...            → processLockSamplePathUpdatePage()
│   ├── isListPage()                    → processListPage()
│   ├── isShowPage()                    → processShowPageIfAuto()
│   ├── isStudyShowPage()               → processStudyShowPage()
│   │                                     processStudyShowPageForNonScrn()
│   │                                     processStudyShowPageForAddCohort()
│   ├── isStudyEditBasicsPage()         → processStudyEditBasicsPageIfFlag()
│   ├── isEpochShowPage()               → processEpochShowPageForImport()
│   │                                     processEpochShowPageForAddCohort()
│   │                                     processEpochShowPage()
│   ├── isCohortShowPage()              → processCohortShowPageImportNonScrn()
│   │                                     processCohortShowPageAddCohort()
│   │                                     processCohortShowPage()
│   ├── isStudyMetadataPage()           → processStudyMetadataPageForEligibilityLock()
│   ├── isEligibilityFormPage()         → processEligibilityFormPageForLocking()
│   ├── isSubjectsListPage()            → processSubjectsListPageForConsent()
│   ├── isSubjectShowPage()             → processSubjectShowPageForConsent()
│   ├── isBarcodeSubjectsPage()         → processBarcodeSubjectsPage()
│   ├── runMode == RUNMODE_ELIG_IMPORT  → startImportEligibilityMapping()
│   │                                     executeEligibilityMappingAutomation()
│   └── runMode == RUNMODE_CLEAR_MAPPING→ executeClearMappingAutomation()
```

---

## 2. Lock Activity Plans

**Button:** "Lock Activity Plans"  
**Section:** `RunActivityPlansFunctions` (line 27034)

```
[Button Click]
├── localStorage.setItem(STORAGE_KEY, "1")
├── localStorage.setItem(STORAGE_RUN_MODE, "activity")
└── location.href = LIST_URL                    // Navigate to Activity Plans list

processListPage()                               // On list page load
├── getRunMode()
├── getPendingIds()
├── getPlanLinks()                              // Collect all plan links from table
├── extractPlanIdFromHref()
├── setPendingIds()
├── (For each plan link, opens in background)
│   ├── monitorCompletionThenAdvance()
│   │   ├── extractAutostateIdFromUrl()
│   │   ├── hasSuccessAlert()
│   │   ├── removePendingId()
│   │   └── updateRunAllPopupStatus()
├── clearRunFlag()
└── clearAllRunState()                          // When all plans processed
```

---

## 3. Lock Sample Paths

**Button:** "Lock Sample Paths"  
**Section:** `RunSamplePathsFunctions` (line 22008)

```
[Button Click]
├── localStorage.setItem(STORAGE_RUN_LOCK_SAMPLE_PATHS, "1")
└── location.href = ".../samples/configure/paths"

processLockSamplePathsPage()                    // On sample paths page
├── (Scans table for all sample path rows)
├── (For each path, fetches detail + update pages in background)
│   ├── fetchPage()
│   ├── parseHtml()
│   └── submitForm()
├── processLockSamplePathDetailPage()           // (no-op, handled in background)
└── processLockSamplePathUpdatePage()           // (no-op, handled in background)
```

---

## 4. Update Study Status

**Button:** "Update Study Status"  
**Section:** `StudyUpdateFunctions` (line 26749)

```
[Button Click]
├── localStorage.setItem(STORAGE_RUN_MODE, "study")
└── location.href = STUDY_SHOW_URL + "?autoupdate=1"

processStudyShowPage()                          // On study show page
├── getQueryParam("autoupdate")
├── getRunMode()
├── getPrimaryIcBarcodeFromStudyShow()
├── setIcBarcode()
├── (Clicks "Edit Basics" link)
└── location.href → edit basics page

processStudyEditBasicsPageIfFlag()              // On edit basics page
├── (Sets study status to Active)
├── (Fills in reason textarea)
├── (Clicks Save)
├── setContinueEpoch()
└── location.href → study show page (for epoch routing)

processStudyShowRouting()                       // Route to epoch
├── waitForSelector('tbody#epochTableBody')
├── isScreeningLabel()
└── location.href → epoch show page
```

---

## 5. Add Cohort Subjects

**Button:** "Add Cohort Subjects"  
**Section:** `AddCohortSubjectsFunctions` (line 25823)

```
[Button Click]
├── localStorage.setItem(STORAGE_RUN_MODE, "epochAddCohort")
├── createPopup() → Add Cohort progress popup
└── location.href = STUDY_SHOW_URL + "?autoaddcohort=1"

processStudyShowPageForAddCohort()              // On study show page
├── getQueryParam("autoaddcohort")
├── (Finds epoch links)
└── location.href → epoch show page

processEpochShowPageForAddCohort()              // On epoch show page
├── (Finds cohort links)
└── location.href → cohort show page

processCohortShowPageAddCohort()                // On cohort show page
├── hasDuplicateVolunteerError()
├── getCohortEditDoneMap()
├── isCohortEditDone()
├── (Clicks "Add Subjects" or edit links)
├── setCheckboxStateById()                      // Toggle volunteer checkboxes
├── setCohortEditDone()
├── getModalCancelButton()
├── hasValidationError()
├── hasNoAssignmentsInImportModal()
└── (Navigates to next cohort or completes)
```

---

## 6. Run ICF Barcode (Informed Consent)

**Button:** "Run ICF Barcode"  
**Section:** `InformedConsentFunctions` (line 24485)

```
[Button Click]
├── isOnSupportedSubjectShowPage()
│
├── [In-Place Flow] (if on subject show page):
│   ├── createPopup() → ICF progress popup
│   └── runIcfBarcodeInPlace()
│       ├── backgroundFetchIcfBarcodeViaIframe()
│       │   ├── backgroundFetchViaIframe()
│       │   └── parseIcBarcodeFromHtml()
│       ├── findEligibleCollectButton()
│       ├── getDataCollectionModalElement()
│       ├── isModalOpen()
│       ├── randomDelay()
│       └── finish()
│
├── [Navigation Flow] (if not on subject page):
│   ├── localStorage.setItem(STORAGE_RUN_MODE, "consent")
│   └── location.href = STUDY_SHOW_URL + "?autoconsent=1"
│
│   processStudyShowPage() → processStudyShowRouting()
│   processEpochShowPage() → (routes to cohort)
│   processCohortShowPage()
│   ├── findCohortRowByVolunteerId()
│   ├── getRowActionButton()
│   ├── getMenuLinkActivatePlan()
│   ├── getMenuLinkActivateVolunteer()
│   └── location.href → subjects list
│
│   processSubjectsListPageForConsent()
│   ├── getSubjectsListTbody()
│   ├── rowHasNoInformedConsentInFourthCell()
│   ├── extractSubjectShowIdFromRow()
│   ├── setConsentScanIndex()
│   ├── appendSelectedVolunteerId()
│   └── location.href → subject show page
│
│   processSubjectShowPageForConsent()
│   ├── subjectShowHasConsentFilled()
│   ├── getIcBarcode()
│   ├── getSelectedVolunteerIds()
│   ├── getConsentScanIndex()
│   └── (Fills barcode input, clicks Collect)
```

---

## 7. Run Button 1-5 (Run ALL Pipeline)

**Button:** "Run Button (1-5)"

Orchestrates all five steps in sequence:

```
[Button Click]
├── localStorage.setItem(STORAGE_RUN_MODE, "all")
├── localStorage.setItem(STORAGE_CONTINUE_EPOCH, "1")
├── localStorage.setItem(STORAGE_CHECK_ELIG_LOCK, "1")
├── createPopup() → Run All progress popup
└── location.href = LIST_URL

Step 1: Lock Activity Plans
├── processListPage() → (same as Feature #2)
│   └── updateRunAllPopupStatus("Running Lock Activity Plans")

Step 2: Update Study Status
├── processStudyShowPage() → processStudyEditBasicsPageIfFlag()
│   └── updateRunAllPopupStatus("Running Study Update")

Step 3: Lock Eligibility
├── processStudyMetadataPageForEligibilityLock()
│   └── processEligibilityFormPageForLocking()

Step 4: Process Epoch → Cohort
├── processEpochShowPage()
│   └── processCohortShowPage()

Step 5: Run ICF Barcode / Consent
├── processSubjectsListPageForConsent()
│   └── processSubjectShowPageForConsent()

(Each step checks getRunMode() == "all" and getContinueEpoch()
 to determine whether to chain to the next step)
```

---

## 8. Import Cohort Subject

**Button:** "Import Cohort Subjects"  
**Section:** `ImportCohortSubjectsFunctions` (line 22361)

```
[Button Click]
├── localStorage.setItem(STORAGE_RUN_MODE, "nonscrn")
├── createPopup() → Import Cohort progress popup
└── location.href = STUDY_SHOW_URL + "?autononscrn=1"

processStudyShowPageForNonScrn()                // On study show page
├── getQueryParam("autononscrn")
├── getNonScrnEpochIndex()
├── setNonScrnEpochIndex()
└── location.href → epoch show page

processEpochShowPageForImport()                 // On epoch show page
├── (Finds cohort links)
└── location.href → cohort show + "?autocohortimport=1"

processCohortShowPageImportNonScrn()            // On cohort show page
├── getImportDoneMap()
├── isImportDone()
├── getImportSubjectIds()
├── setImportSubjectIds()
├── clearImportSubjectIds()
├── pickRandomUnusedLetter()
├── extractVolunteerIdFromChosenText()
├── extractVolunteerIdFromFreeText()
├── setUsedLetters() / getUsedLetters()
├── clearUsedLetters()
├── findCohortRowByVolunteerId()
├── parseCohortIdFromUpdateHref()
├── setCheckboxStateById()
├── setImportDone()
├── markImportDoneIfSuccessOnLoad()
├── setCohortGuard() / getCohortGuard()
└── (Navigates to next cohort or back to study show)
```

---

## 9. Pull Barcode

**Button:** "Pull Barcode"  
**Section:** `BarcodeFunctions` (line 27542)

```
[Button Click]
└── APS_RunBarcode()
    ├── getSubjectFromBreadcrumbOrTooltip()
    ├── setBarcodeSubjectText()
    ├── setBarcodeSubjectId()
    ├── setBarcodeResult()
    ├── (If on barcode subjects page):
    │   └── processBarcodeSubjectsPage()
    │       ├── getBarcodeSubjectText()
    │       ├── getBarcodeSubjectId()
    │       ├── normalizeSubjectString()
    │       └── (Clicks matching row)
    ├── (If needs navigation):
    │   └── location.href → barcode subjects page
    ├── clearBarcodeSubjectId()
    ├── clearBarcodeSubjectText()
    └── clearBarcodeResult()
```

---

## 10. Pull Lab Barcode

**Button:** "Pull Lab Barcode"

```
[Button Click]
└── pullLabBarcodeInit()
    ├── openBarcodeProgressPanel()
    │   ├── createPopup() → dual-panel progress popup
    │   ├── setAriaBusyOn()
    │   └── (Builds left panel: scanned icons, right panel: status)
    ├── (Scans page for barcode icons)
    ├── processIconsSequentially(iconList)
    │   ├── extractBarcodeValue()
    │   ├── updateBarcodeRightPanelStatus()
    │   ├── updateBarcodeRightPanelSummary()
    │   ├── announceBarcodeAriaLive()
    │   └── barcodeYieldToUI()
    └── stopPullLabBarcode()                    // Teardown
        ├── (Clears all timeouts, intervals, observers, RAF IDs, idle callbacks)
        ├── setAriaBusyOff()
        └── (Closes popup, resets state)
```

---

## 11. Add Existing Subject

**Button:** "Add Existing Subject"  
**Section:** Lines 5720–7079

```
[Button Click]
└── startAddExistingSubject()
    ├── (Fetches study page, parses epochs)
    ├── showAESProgressPopup()
    ├── setAESStep() / getAESStep()
    ├── setAESProgress() / getAESProgress()
    ├── localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_ADD_EXISTING_SUBJECT)
    └── location.href → study show page

processAESOnPageLoad()                          // On each page load
├── getRunMode()
├── getAESStep()
├── [Step: epoch]     → processEpochPage()
│   └── location.href → cohort page
├── [Step: cohort]    → processCohortPage()
│   ├── (Clicks "Edit" on cohort)
│   ├── (Sets checkbox states)
│   └── (Clicks Save)
├── [Step: afterEdit] → processAfterCohortEdit()
│   └── addSubjectToCohort()
│       ├── formatDateForClinSpark()
│       └── (Fills modal form, clicks Add)
├── [Step: afterAdd]  → processAfterAddSubject()
│   └── processActivatePlan()
│       └── processActivateVolunteer()
├── showAESError()
├── showAESComplete()
└── clearAddExistingSubjectData()
```

---

## 12. Search Methods (Methods Library)

**Button:** "Search Methods"

```
[Button Click]
└── openMethodsLibraryModal()
    ├── fetchMethodsIndex(forceRefresh)
    │   ├── getMethodsCache()
    │   └── setMethodsCache()
    ├── buildTagList()
    ├── populateTagSelect()
    ├── doSearch()
    │   ├── filterAndSortMethods()
    │   │   └── scoreMethod()
    │   ├── getFavorites() / getPins() / getRecents()
    │   └── renderList()
    ├── selectMethod()
    │   ├── fetchMethodBody()
    │   ├── addRecent()
    │   └── saveLastMethod()
    ├── toggleFavorite() / togglePin()
    ├── saveLastSearch() / saveLastTag() / saveSortOrder()
    └── closeModal()
```

---

## 13. PLAP Builder (Build Procedure Log)

**Button:** "PLAP Builder"  
**Section:** `BuildProcedureLogFunctions` (line 7080), `SABuilderFunctions` (line 8853)

```
[Button Click]
└── runBuildProcedureLog()
    ├── isOnBPLPage()
    ├── scanExistingBPLTable()
    ├── scanExistingBPLTableEnhanced()
    │   ├── bplDetectTimepointColumn()
    │   ├── bplParseTimepointText()
    │   ├── bplParseStatusCell()
    │   └── bplParseClinSparkDateTime()
    ├── bplFetchSegmentsPageHtml()
    ├── bplMergeSaTableWithSession()
    ├── restoreBPLSessionState() / saveBPLSessionState()
    ├── createBPLSelectionGUI()
    │   ├── renderStudyEventsList()
    │   ├── renderFormsDragList()
    │   ├── renderTimePanel()
    │   │   ├── createBPLTimeInput()
    │   │   ├── createBPLCheckbox()
    │   │   └── createBPLTextInput()
    │   ├── renderCenterPanel()
    │   │   ├── loadFormDataToPanel()
    │   │   ├── saveFormDataFromPanel()
    │   │   └── getSegmentRefDateTime()
    │   ├── bplFormatTimePoint()
    │   ├── bplBuildStatusIcons()
    │   ├── bplBuildFormLabel()
    │   ├── bplValidateConfiguration()
    │   ├── runAutoValidation()
    │   └── createResizeHandle() / applyPanelWidths()
    ├── createBPLMappedItemList()
    ├── createBPLProgressPopup()
    │   ├── renderItems()
    │   └── updateSummary()
    ├── bplFormatClinSparkDateTime()
    ├── bplApplyOffsetToDateTime()
    ├── bplComputeExampleTime()
    ├── bplParseOffsetText()
    └── clearBPLSessionState() / clearBPLStorage()
```

---

## 14. Populate Form

**Button:** "Populate Form"  
**Section:** `FormAutomationFunctions` (line 20566)

```
[Button Click]
├── showRunFormOptionPanel()                    // Choose mode (IR, Random, etc.)
│   ├── doSelect()
│   ├── doCancel()
│   └── cleanup()
├── setFormValueMode()
└── runFormAutomationV2()
    ├── getFormValueMode()
    ├── isDataCollectionSubjectPage()
    ├── getFormDataRows()
    ├── findNextVisibleUnprocessedGroupV2()
    │   ├── groupHasActionableItems()
    │   └── hasActionableControl()
    ├── processItemGroupDirectV2(groupDiv)
    │   ├── parseItemGroupId()
    │   ├── getItemRowId()
    │   ├── getItemTextCellFromRow()
    │   ├── findRangeSpecForRow()
    │   │   ├── getRangeTextFromHelp()
    │   │   ├── getRangeTextFromItemText()
    │   │   └── parseRangeSpecFromText()
    │   ├── pickIntegerForSpec() / pickDecimalForSpec()
    │   ├── getDecimalPlacesFromMeta()
    │   ├── randomIntInclusive()
    │   ├── randomIntInInclusiveRange()
    │   ├── getText()
    │   └── ensureAuditDefault()
    ├── isBarcodeVerifyModalVisible()
    ├── isFormOrderModalVisible()
    └── sleep()
```

---

## 15. Collect All

**Button:** "Collect All"

```
[Button Click]
└── runCollectAll()
    ├── clearCollectAllData()
    ├── isDataCollectionSubjectPage()
    ├── getFormDataRows()
    ├── getFormNameFromRow()
    ├── getFormStatusFromRow()
    ├── getFormDataRowId()
    ├── findCollectButtonInRow()
    ├── APS_RunBarcode()                        // Run barcode if needed
    ├── runFormAutomationV2()                   // Run form population
    ├── addFormToList()
    ├── isBarcodeVerifyModalVisible()
    ├── isFormOrderModalVisible()
    ├── finishWithResult()
    └── cleanup()
```

---

## 16. Import I/E (Import Eligibility)

**Button:** "Import I/E"  
**Section:** `SubjectEligibilityFunctions` (line 14996)

```
[Button Click]
└── startImportEligibilityMapping()
    ├── isEligibilityListPage()
    ├── buildImportEligPopup()
    │   ├── buildImportIEReviewPanel()
    │   │   ├── buildHierarchy()
    │   │   ├── deduplicateMappings()
    │   │   ├── renderLeftList()
    │   │   ├── renderRightList()
    │   │   │   ├── toggleItem() / toggleSA() / togglePlan()
    │   │   │   ├── setDescendants() / setSADescendants()
    │   │   │   └── updateSelectedCountAndConfirmState()
    │   │   ├── renderPoolList()
    │   │   │   ├── createDropbox()
    │   │   │   ├── removeFromAvailablePool()
    │   │   │   ├── restoreToAvailablePool()
    │   │   │   ├── sortAvailablePool()
    │   │   │   └── findPoolEntryByValue()
    │   │   ├── renderRightPanelFlatCodeView()
    │   │   │   ├── buildFlatCodeOrderedDataset()
    │   │   │   ├── sortFlatDataset()
    │   │   │   ├── filterFlatRowsToDuplicates()
    │   │   │   ├── applySearchFilterToFlatView()
    │   │   │   └── toggleDuplicatesFilterOnCodeView()
    │   │   ├── captureSelectionState() / restoreSelectionToHierarchy() / restoreSelectionToFlat()
    │   │   ├── syncSelectionStateBetweenViews()
    │   │   ├── ensureGenderStateForKey() / renderRowGenderControl()
    │   │   ├── parsePoolCodeParts() / formatPoolCodeDisplay()
    │   │   ├── getMappingKeyForPool() / getMappingKey()
    │   │   └── truncateToWords()
    │   ├── parseStoredKeywords()
    │   ├── getFormExclusion() / setFormExclusion()
    │   ├── getFormPriority() / setFormPriority()
    │   ├── getFormPriorityOnly() / setFormPriorityOnly()
    │   ├── getPlanPriority() / setPlanPriority()
    │   ├── getIgnoreKeywords() / setIgnoreKeywords() / addToIgnoreKeywords()
    │   └── showWarningPopup()
    ├── buildProgressPopup()
    │   └── reportProgress()
    └── executeEligibilityMappingAutomation()
        ├── isEligibilityListPage()
        ├── isValidEligibilityPage()
        ├── getImportedItemsSet() / persistImportedItemsSet()
        ├── parseItemCodeFromEligibilityLabel()
        ├── parseItemCodeFromEligibilityOptionText()
        ├── parseSexFromEligibilityText()
        ├── getDesiredSexForExecution()
        ├── firstNonEmptyOptionIndex()
        ├── isExcluded() / isPriority() / isPlanPriority()
        ├── getCheckItemCache() / getCachedCheckItem() / cacheCheckItem() / persistCheckItemCache()
        ├── getItemRefOptionsSignature()
        ├── buildOrGetSelectOptionMap()
        ├── select2TriggerChange()
        ├── isCurrentSelectValue()
        ├── extractIECode() / extractIECodeStrict()
        ├── normalizeTextForCompare()
        ├── importIERandomDelay()
        ├── formatElapsedTime()
        ├── addToImportEligCompletedList()
        ├── addToImportEligFailedList()
        ├── addToImportEligExcludedList()
        ├── getBaseUrl()
        ├── setLastMatchSelection() / getLastMatchSelection() / clearLastMatchSelection()
        └── clearEligibilityWorkingState()
```

---

## 17. Clear Mapping

**Button:** "Clear Mapping"  
**Section:** `ClearEligibilityFunctions` (line 14801)

```
[Button Click]
└── startClearMapping()
    ├── localStorage.setItem(STORAGE_RUN_MODE, RUNMODE_CLEAR_MAPPING)
    └── executeClearMappingAutomation()
        ├── isEligibilityListPage()
        ├── (Iterates through eligibility items)
        ├── (Clears select2 mappings)
        ├── select2TriggerChange()
        └── clearEligibilityWorkingState()
```

---

## 18. Archive/Update Forms

**Button:** "Archive/Update Forms"  
**Section:** `ArchiveUpdateFormsFunctions` (line 2563)

```
[Button Click]
└── runArchiveUpdateForms()
    ├── isOnArchiveUpdateFormsPage()
    ├── scanSATableForArchiveUpdate()
    ├── buildUniqueFormsMap()
    ├── (Opens Add modal to collect target form options)
    ├── createArchiveUpdateFormsGUI()
    │   ├── renderSourcePanel()
    │   ├── renderTargetPanel()
    │   └── validateSelection()
    ├── createPopup() → selection popup
    ├── (On Confirm):
    │   ├── createArchiveUpdateProgressPopup()
    │   ├── checkTargetFormExists()
    │   ├── findRowInDOM()
    │   ├── normalizeVisibilityText()
    │   ├── (For each occurrence):
    │   │   ├── Archive old form (click archive link, fill reason)
    │   │   ├── Add new form (click Add, select form, set visibility)
    │   │   └── Update progress popup
    │   └── normalizeSAText()
    └── createPopup() → completion summary
```

---

## 19. Copy Activity Forms

**Button:** "Copy Activity Forms"

```
[Button Click]
└── runCopyFormsToStudyEvents()
    ├── isOnCopyFormsPage()
    ├── (Scans SA table for existing forms)
    │   ├── scanSATableForArchiveUpdate()       // Reuses same scanner
    │   ├── buildHierarchicalFormMap()
    │   └── buildExistingFormsMap()
    ├── (Collects study events from Add modal)
    ├── createCopyFormsGUI()
    │   ├── renderLeftPanel()                   // Hierarchical form selection
    │   ├── renderRightPanel()                  // Study event selection
    │   ├── updateSelectionInfo()
    │   └── validateSelection()
    ├── createPopup() → selection popup
    ├── (On Continue):
    │   ├── rebuildScheduledActivityForTarget()
    │   ├── formExistsInMap()                   // Duplicate detection
    │   ├── createCopyFormsProgressPopup()
    │   │   ├── updateStatus() / updateProgress()
    │   │   ├── setItemStatus()
    │   │   └── showSummary()
    │   └── (For each form × study event):
    │       ├── Click Add button
    │       ├── Select segment, study event, form
    │       └── Click Save
    └── normalizeSAText()
```

---

## 20. Item Method Forms (Parse Method)

**Button:** "Item Method Forms"  
**Section:** `ParseMethodFunctions` (line 13694)

```
[Button Click]
└── openParseMethod()
    ├── showParseMethodPopup()
    │   ├── doConfirm()
    │   │   └── populateFormsFromParseMethod()
    │   │       ├── doFindFormWithForms()
    │   │       │   ├── previewFormsByKeyword()
    │   │       │   │   └── formMatchContainsAllTokens()
    │   │       │   ├── showFormNoMatchPopup()
    │   │       │   └── addParseMethodResult()
    │   │       └── showParseMethodCompletedPopup()
    │   └── keyHandler()
    ├── clearParseMethodState()
    ├── stopParseMethodAutomation()
    ├── clearParseMethodStoredResults()
    ├── showParseMethodResults()
    └── checkAndRestoreParseMethodPopup()
```

---

## 21. Find Form

**Button:** "Find Form"

```
[Button Click]
└── openFindForm()
    ├── showFindFormPopup()
    │   ├── previewFormsByKeyword()
    │   │   ├── formNormalize()
    │   │   └── formMatchContainsAllTokens()
    │   ├── showFormNoMatchPopup()
    │   ├── gatherStatusSelections()
    │   ├── findSubjectIdentifierForFindForm()
    │   ├── getSubjectIdentifierForAE()
    │   ├── doContinue()
    │   └── keyHandler()
    └── processFindFormOnList()                 // On data list page after navigation
        ├── applyStatusValuesOnList()
        ├── resetStatusValuesOnList()
        ├── clearSelect2ChoicesByContainerId()
        └── aeNormalize()
```

---

## 22. Find Study Events

**Button:** "Find Study Events"

```
[Button Click]
└── openFindStudyEvents()
    ├── showFindStudyEventsPopup()
    │   ├── previewStudyEventsByKeyword()
    │   │   └── studyEventMatchContainsAllTokens()
    │   ├── showStudyEventNoMatchPopup()
    │   ├── gatherStatusSelections()
    │   ├── doContinue()
    │   └── keyHandler()
    └── processFindStudyEventsOnList()          // On data list page after navigation
        ├── applyStatusValuesOnList()
        ├── resetStudyEventsOnList()
        ├── deselectAllOptionsBySelect()
        └── resetStatusValuesOnList()
```

---

## 23. Pause / Resume

**Button:** "Pause" / "Resume"

```
[Button Click]
├── isPaused()
├── setPaused(true/false)
└── (Updates all active popup statuses to show Paused/Running)
    ├── RUN_ALL_POPUP_REF status update
    ├── LOCK_SAMPLE_PATHS_POPUP_REF status update
    └── IMPORT_ELIG_POPUP_REF status update

(All process* functions check isPaused() at entry and exit early if paused)
```

---

## 24. Settings

**Button:** ⚙ (gear icon in panel header)

```
[Button Click]
└── openSettingsPopup()
    ├── getButtonVisibility() / setButtonVisibility()
    ├── isButtonVisible()
    ├── getPanelHotkey() / setPanelHotkey()
    ├── getThemeMode() / setThemeMode()
    ├── isGlassTheme()
    ├── applyThemeToUiRoots()
    ├── injectThemeStylesIfNeeded() / removeThemeStyles()
    └── createPopup() → settings popup with toggles per button
```

---

## 25. Shared Utilities

**Section:** `SharedUtilityFunctions` (line 27852), `SharedPanelFunctions` (line 29087)

### State Management
```
clearAllRunState()
├── clearRunMode()
├── clearPendingIds()
├── clearSelectedVolunteerIds()
├── clearContinueEpoch()
├── clearCohortGuard()
├── clearAfterRefresh()
├── clearConsentScanIndex()
├── clearLastVolunteerId()
├── clearBarcodeResult()
└── clearIcBarcode()
```

### Page Detection
```
isListPage()                    isStudyShowPage()
isShowPage()                    isStudyEditBasicsPage()
isEpochShowPage()               isCohortShowPage()
isSubjectsListPage()            isSubjectShowPage()
isStudyMetadataPage()           isEligibilityFormPage()
isBarcodeSubjectsPage()         isEligibilityListPage()
isDataCollectionSubjectPage()
```

### UI & Panel
```
makePanel()                     createPopup()
addButtonToPanel()              clampPopupPosition()
setupResizeHandle()             clampPanelPosition()
scale()                         getStoredUIScale() / updateUIScale()
log()                           clearLogs()
getStoredPos()                  setStoredPanelSize() / getStoredPanelSize()
setPanelCollapsed()             updatePanelCollapsedState()
getPanelHidden()                applyPanelHiddenState()
bindPanelHotkeyOnce()          normalizeKeyForMatch() / keyMatchesToggle()
```

### Common Helpers
```
normalizeText()                 normalizeSAText()
isScreeningLabel()              getQueryParam()
getCurrentPlanId()              hasSuccessAlert()
getRunMode() / setRunMode()     getPendingIds() / setPendingIds()
openInTab()                     closeTabWithFallback()
sleep()                         recreatePopupsIfNeeded()
```

### Theme System
```
getThemeMode() / setThemeMode()
isGlassTheme()
injectThemeStylesIfNeeded()
removeThemeStyles()
applyThemeToUiRoots()
```

---

## Architecture Notes

- **Navigation-based automation:** Many features work by setting `localStorage` flags, navigating to a new URL, and continuing automation on page load via `init()` routing.
- **Popup persistence:** `recreatePopupsIfNeeded()` restores progress popups after page navigations.
- **Pause system:** Global `isPaused()` check at the start of every `process*` function.
- **Run modes:** `STORAGE_RUN_MODE` tracks the current automation mode (`activity`, `study`, `consent`, `all`, `nonscrn`, `epochImport`, `epochAddCohort`, `eligibilityImport`, `clearMapping`, `addExistingSubject`).
- **Glass theme:** Optional glassmorphism UI theme toggled in Settings, applied via CSS class injection.
