/* jshint strict: false */

// Version: v1
// Purpose: Protocol Check Any QTcF increase ≥60 msec from Baseline OR average QTcF 500 msec (REPEAT).

// Add item names
var baselineForms = [
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)"
];
var baselineFormStudyEvents = [
    "Visit 2 Week 1 Day 0"
];

var baselineItem = [
    "Baseline QTcF"
];
var qtcfItems = [
    "QTcF #1",
    "QTcF #2",
    "QTcF #3"
];

// ======== Don't modify ========
var item = itemJson.item;

try {
    var form = pullForm(baselineFormStudyEvents, baselineForms);
    if (!form) return true;
    var baseline = pullItemFromForm(form, baselineItem);
    var qtcf = pullItemFromForm(formJson, qtcfItems);

    if (!baseline || baseline == null || !qtcf || qtcf == null) return null;
    var diff = qtcf - baseline;
    if (diff >= 60 || qtcf > 500) return "Y";
    else if (diff < 60 || qtcf <= 500) return "N";
    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Iterates through study events and form names to find the first matching completed form.
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
}

// Searches a form's item groups for an item matching the target name and returns its value or the item object.
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
    return null;
}

// Retrieves the first completed (or nonconformant) form instance from a study event.
function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

// Filters an array of form data to return only entries with valid completion status (Complete, Nonconformant, or Incomplete).
function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        } else {

        }
    }
    return completedForms;
}
