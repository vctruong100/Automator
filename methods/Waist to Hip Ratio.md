// Version: v1
// Purpose: Computes waist-to-hip ratio.

const waistItemList = [
    "Waist Circumference #1", 
    "Waist Circumference #2"
];
const hipItemList = [
    "Hip Circumference #1", 
    "Hip Circumference #2"
];

const heightItemList = [
    "BM_HT",
    "BMI_Height",
];

const sigfig = itemJson.item.significantDigits;
const maxCount = 2; 

try {
    var waist = calculateAverage(populateList(formJson, waistItemList), sigfig);
    var hip = calculateAverage(populateList(formJson, hipItemList), sigfig);

    log();

    if (hip && waist && hip !== null && waist !== null) return (waist / hip).toFixed(sigfig);

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Logs the current values of sys, dia, and hr variables for debugging purposes.
function log() {
    logger("Max count: " + maxCount);
    logger("Waist: " + waist);
    logger("Hip: " + hip);
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

// Collects numeric values from form items matching target names, stopping when an attached item is encountered (first-to-last order).
function populateList(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                    if (list.length >= maxCount) return list;
                }
            }
        }
    }
    return list;
}

// Calculates the arithmetic mean of an array of numeric values, ignoring non-numeric entries.
function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;
    for (var i = 0; i < values.length; i++) {
        logger("Values: " + values[i])
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);
    return Math.round(avg * factor) / factor;
}

// Retrieves the first completed (or nonconformant) form instance from a study event.
function checkForm(studyevent, form) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms, true);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm[completedForm.length - 1];
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