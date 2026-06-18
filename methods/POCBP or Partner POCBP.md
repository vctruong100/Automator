// Version: v1
// Purpose: Handles point-of-care blood pressure data capture.

const studyEvent = [
    "SCREENING", 
    "Screening"
];
const formName = [
    "REM_Study Reminders (SCRN)",
];
const itemName = [
    "POCBP or Partner POCBP"
];
const attachedItemCodeList = [
    "YES",
    "NO",
]
const pulledItemCodeList = [
    "YES",
    "NO",
]

const gender = formJson.form.subject.volunteer.sexMale;

try {
    if (gender) return itemJson.item.codeListItems[1].codedValue; // return no

    var form = pullForm(studyEvent, formName);
    if (!form) return null;

    var POCBP = pullItemFromForm(form, itemName);
    if (!POCBP || POCBP.value == null) return null;

    log();

    if (POCBP.value == POCBP.codeListItems[0].codedValue) return itemJson.item.codeListItems[0].codedValue; // return yes
    else if (POCBP.value == POCBP.codeListItems[1].codedValue) return itemJson.item.codeListItems[1].codedValue; // return no
    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Logs the current values of sys, dia, and hr variables for debugging purposes.
function log() {
    logger("POCBP: " + POCBP);
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
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item;
        }
    }
    return null;
}

// Retrieves the first completed (or nonconformant) form instance from a study event.
function checkForm(studyevent, form) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms, true);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm[0];
    }
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
