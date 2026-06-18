<<<<<<< Updated upstream
/* jshint strict: false */ 
To convert to Int:
value.toFixed(0);
=======
/* jshint strict: false */
>>>>>>> Stashed changes

// Version: v1
// Purpose: Provides reusable helper functions for form data retrieval, item extraction, and data normalization across ClinSpark automations.

// ---------------------------------------------------------------------------
// CONVERSION & LOGGING HELPERS
// ---------------------------------------------------------------------------
// To convert a number to an integer string:
//   value.toFixed(0);
// For a float with specific significant figures:
//   value.toFixed(sigfig);
// To log the current item JSON for debugging:
//   logger(JSON.stringify(itemJson, null, 2));

var sigfig = itemJson.item.significantDigits;

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

<<<<<<< Updated upstream
const sigfig = itemJson.item.significantDigits;
=======
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
>>>>>>> Stashed changes

// ---------------------------------------------------------------------------
// FORM DATA RETRIEVAL HELPERS
// ---------------------------------------------------------------------------

// Searches a form's item groups in first-to-last order for an item whose
// name matches the target list and returns its value.
// Skips canceled groups and items with null or empty values.
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
<<<<<<< Updated upstream
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
=======
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

>>>>>>> Stashed changes
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

<<<<<<< Updated upstream
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' || 
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            keepers.push(formData);
        } else {

        }
    }
    return keepers;
}

// By last to first
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
=======
// Searches a form's item groups in last-to-first order for an item whose
// name matches the target list and returns its value.
// Use this variant when the most recently entered value should take precedence.
function pullItemFromFormLastToFirst(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

>>>>>>> Stashed changes
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") {
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
<<<<<<< Updated upstream
        if (formData.form.canceled == false && 
            formData.form.itemGroups[0].canceled == false && 
            (formData.form.dataCollectionStatus == 'Complete' || 
            (INCLUDE_NONCONFORMANT_DATA == true && 
            formData.form.dataCollectionStatus == 'Nonconformant') || 
            formData.form.dataCollectionStatus == "Incomplete")) 
        {
            keepers.push(formData);
        } else {}
=======
        if (formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (formData.form.dataCollectionStatus == 'Complete' ||
            (INCLUDE_NONCONFORMANT_DATA == true &&
            formData.form.dataCollectionStatus == 'Nonconformant') ||
            formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        }
>>>>>>> Stashed changes
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
// var relatedItemDataContext = JSON.parse(getRelatedItemDataContext('BE_Date of sample 2'));
// var itemDataContext = JSON.parse(getRelatedItemDataContext());
// var transferBarcodes = relatedItemDataContext.transferBarcodes;
// logger(JSON.stringify(relatedItemDataContext, null, 2));
// return transferBarcodes;
