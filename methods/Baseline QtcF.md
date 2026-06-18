// Version: v1
// Purpose: Computes baseline QTcF from screening or predose ECG forms.

// ======== Add item names ========
const baselineStudyEvent = [
    "Visit 2 Week 1 Day 0",
];
const baselineFormName = [
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)"
];

const baselineItem = [
    "Baseline QTcF"
];

// ======== Don't modify ========
const sigfig = itemJson.item.significantDigits;

try {
    var form = pullForm(baselineStudyEvent, baselineFormName);
    if (!form) return null;

    return (pullItemFromForm(form, baselineItem)).toFixed(sigfig);
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return parseFloat(item.value);
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
<<<<<<< Updated upstream
    return keepers;
}
=======
    return completedForms;
}
>>>>>>> Stashed changes
