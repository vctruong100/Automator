// Version: v1
// Purpose: Evaluates pregnancy test results and sets flags.

const formName = formJson.form.name;

const studyEvent = [
    "SCREENING",
    "Screening"
];
const formNames = [
    "REP_Reproductive Status and Contraception – Female",
    "REP_Female Reproductive Status and Contraceptive Use",
    "♀️ REP_CONTRACEPTIVE & REPRODUCTIVE STATUS (Female)"
];
const childbearingItem = [
    "Is the female subject of childbearing potential?",
    "Childbearing potential?"
];
const statusItem = [
    "Female Reproductive Status",
    "Female subject is considered non-childbearing potential due to"
];

const gender = formJson.form.subject.volunteer.sexMale;
const age = formJson.form.subject.volunteer.age;

try {
    logger(age);
    logger(gender);
    if (gender) return itemJson.item.codeListItems[3].codedValue; // return None

    var form = pullForm(studyEvent, formNames);
    if (!form) return null;

    var childbearing = pullItemFromForm(form, childbearingItem);
    if (childbearing && childbearing.value !== null && childbearing.value == childbearing.codeListItems[0].codedValue) return itemJson.item.codeListItems[0].codedValue; // if childbearing = Yes, return

    var status = pullItemFromForm(form, statusItem);
    if (status && status.value !== null && (status.value == status.codeListItems[2].codedValue)) return itemJson.item.codeListItems[3].codedValue; // return None
    else if (status && status.value !== null) return itemJson.item.codeListItems[0]; // return pregnancy for other status

    return itemJson.item.codeListItems[3].codedValue; // return none
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Logs the current values of sys, dia, and hr variables for debugging purposes.
function log() {
    logger("Childbearing: " + childbearing);
    logger("Status: " + status);
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
    return completedForms;
}

// Checks if the input string contains the specified keyword (case-insensitive). Returns false for null/undefined input.
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

<<<<<<< Updated upstream
    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}
=======
    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}
>>>>>>> Stashed changes
