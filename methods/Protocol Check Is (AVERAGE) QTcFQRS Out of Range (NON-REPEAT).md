// Version: v1
// Purpose: Protocol check for averaged QTcF/QRS out-of-range (non-repeat).

// Add Item names
const QTcFitems = [
    '- QTcF (≤ 450 msec "Males") (≤ 470 msec "Females")', 
    "QTcF", 
    "QTcF_Protocol",
    "QTcF #1",
    "QTcF #2", 
    "QTcF #3"
];

const QRSitems = [
    "- QRS duration ( ≤ 120 ms)", 
    "QRS", 
    "QRS_Protocol"
];

// Inclusive (Edit)
var QTcF_max_range = 450;
var QRS_max_range = 120;

// ======== Don't modify ========
var QTcFmaxCount = 3; 
var QTcFlist = [];
var QTcFavg = 0;

var QRSmaxCount = 3; 
var QRSlist = [];
var QRSavg = 0;

try {
    QTcFlist = populateList(formJson, QTcFitems, QTcFlist, QTcFmaxCount);
    QTcFavg = calculateAverage(QTcFlist);

    QRSlist = populateList(formJson, QRSitems, QRSlist, QRSmaxCount);
    QRSavg = calculateAverage(QRSlist);

    log();

    if (QTcFlist.length !== QTcFmaxCount || QRSlist.length !== QRSmaxCount) return itemJson.item.codeListItems[0].codedValue; // return pending result

    if (QTcFavg > QTcF_max_range || QRSavg > QRS_max_range) return itemJson.item.codeListItems[2].codedValue; // return Out of protocol range
    else if (QTcFavg <= QTcF_max_range || QRSavg > QRS_max_range) return itemJson.item.codeListItems[1].codedValue; // return within protocol range

    return attachedItemCodeList[0];
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Logs the current values of sys, dia, and hr variables for debugging purposes.
function log() {
    logger("List: " + QTcFlist);
    logger("List length: " + QTcFlist.length);
    logger("Max count: " + QTcFmaxCount);
    logger("Average: " + QTcFavg);
    
    logger("List: " + QRSlist);
    logger("List length: " + QRSlist.length);
    logger("Max count: " + QRSmaxCount);
    logger("Average: " + QRSavg);
}

// Collects numeric values from form items matching target names, stopping when an attached item is encountered (first-to-last order).
function populateList(form, targetItem, list, maxCount) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
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

// Iterates through study events and form names to find the first matching completed form.
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
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

// Calculates the arithmetic mean of an array of numeric values, ignoring non-numeric entries.
function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    return avg;
}