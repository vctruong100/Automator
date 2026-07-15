/* jshint strict: false */
// ---------------------------------------------------------------------------
// CONVERSION & LOGGING HELPERS
// ---------------------------------------------------------------------------
// To convert a number to an integer string:
value.toFixed(0);
// For a float with specific significant figures:
value.toFixed(sigfig);
// To log the current item JSON for debugging:
logger(JSON.stringify(itemJson, null, 2));

var sigfig = itemJson.item.significantDigits;



// For detailed function description, scroll below. The following functions are for copy-paste purposes.
function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
        }
    }
    return false;
}
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function checkFormMostRecent(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[completedForm.length - 1];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false &&
            (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') ||
                formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        }
    }
    return completedForms;
}
// ---------------------------------------------------------------------------
// STRING & NORMALIZATION HELPERS
// ---------------------------------------------------------------------------

// Checks if the input string contains the specified keyword (case-insensitive).
// Returns false for null or undefined input to prevent runtime errors.
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

// Strictly normalizes a string using Unicode NFKC decomposition, removing
// zero-width characters and standardizing whitespace to a single space.
// This ensures consistent comparison of user-entered text values.
function normalizeNameStrict(inputString) {
    if (inputString === undefined || inputString === null) {
        return "";
    }
    var normalizedString = String(inputString);
    normalizedString = normalizedString.normalize ? normalizedString.normalize('NFKC') : normalizedString;
    normalizedString = normalizedString.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ');
    normalizedString = normalizedString.replace(/\s+/g, ' ');
    normalizedString = normalizedString.trim().toLowerCase();
    return normalizedString;
}

// ---------------------------------------------------------------------------
// FORM DATA RETRIEVAL HELPERS
// ---------------------------------------------------------------------------

// Searches a form's item groups in first-to-last order for an item whose
// name matches the target list and returns its value.
// Skips canceled groups and items with null or empty values.
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

// Searches a form's item groups in last-to-first order for an item whose
// name matches the target list and returns its value.
// Use this variant when the most recently entered value should take precedence.
function pullItemFromFormLastToFirst(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

// Returns the item group ID that contains the current item.
// Useful when you need to scope a search to the same item group as the active item.
function getItemGroupID(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var currentItem = items[j];
            if (currentItem.id === item.id) {
                return group.id;
            }
        }
    }
    return null;
}

// Iterates through a list of study events and form names to find the first
// matching form that passes the completion check.
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
}

// ---------------------------------------------------------------------------
// FORM COMPLETION & FILTERING HELPERS
// ---------------------------------------------------------------------------

// Retrieves the first completed (or nonconformant) form instance from a study event.
// Iterates newest-to-oldest and returns the first match from the beginning of the filtered array.
function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

// Retrieves the most recently completed (or nonconformant) form instance from a study event.
// Iterates newest-to-oldest and returns the last element of the filtered array.
function checkFormMostRecent(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[completedForm.length - 1];
}

// Filters an array of form data to return only entries with a valid completion status.
// Accepted statuses: Complete, Nonconformant (if INCLUDE_NONCONFORMANT_DATA is true), or Incomplete.
// Canceled forms or item groups are excluded. Iterates newest-to-oldest.
function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false &&
            (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') ||
                formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        }
    }
    return completedForms;
}

// Filters an array of form data to return only entries with a valid completion status
// that also match the current form's timepoint. This variant is used when multiple
// timepoint versions of the same form may exist within a single study event.
function collectCompletedByTimepoint(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    var filteredForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (formData.form.dataCollectionStatus == 'Complete' ||
            (INCLUDE_NONCONFORMANT_DATA == true &&
            formData.form.dataCollectionStatus == 'Nonconformant') ||
            formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        }
    }
    for (var j = 0; j < completedForms.length; j++) {
        var candidateForm = completedForms[j];
        if (candidateForm.form.timepoint == timepoint) filteredForms.push(candidateForm);
    }
    return filteredForms;
}

// ---------------------------------------------------------------------------
// EXAMPLE: RELATED ITEM DATA CONTEXT
// ---------------------------------------------------------------------------
// The block below demonstrates how to retrieve related item data contexts.
// Uncomment and adjust the item name as needed for your specific automation.
//
var item = itemJson.item;

var relatedItemDataContext = JSON.parse(getRelatedItemDataContext('💧GenX Screening Urinalysis + UDS'));
var itemDataContext = JSON.parse(getRelatedItemDataContext());
var transferBarcodes = relatedItemDataContext.transferBarcodes;
logger(JSON.stringify(relatedItemDataContext, null, 2));

var rawgroupName = getItemDataContextByItemDataId(item.id);
logger(JSON.stringify(JSON.parse(rawgroupName), null, 2))
var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;

var itemDataContext = JSON.parse(getItemDataContext());
logger(JSON.stringify(itemDataContext, null, 2));
/*
getRelatedItemDataContext()
{
"collectedBarcodes": "F6516739-C7",
"transferBarcodes": "F6516739-C7",
"foundItemDataId": 6516739,
"foundItemName": "💧GenX Screening Urinalysis + UDS",
"foundItemGroupName": "GenX Screening UA/UDS"
}
getItemDataContextByItemDataId()
{
"collectedBarcodes": "F006516738",
"transferBarcodes": "",
"foundItemDataId": 6516738,
"foundItemName": "LAB_Fasting Check for lab orders",
"foundItemGroupName": "GenX Screening UA/UDS"
}
getItemDataContext()
{
"itemDataId": 6516738,
"itemGroupDataId": 1408074,
"formDataId": 524212,
"siteName": "CenExel ACT",
"investigatorName": "Dr. Peter Winkle (Labs can repeat ONCE only for SCRN & Day-1 for eligibility)",
"investigatorId": ""
} */