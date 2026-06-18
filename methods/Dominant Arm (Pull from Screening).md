// Version: v1
// Purpose: Pulls the subject's dominant arm from Screening vitals.

// Add Item Names
const itemList = [
    "Hip Circumference #1", 
    "Hip Circumference #2"
];
const attachedItem = [
    "Average Hip Circumference"
];

// ======== Don't modify ========
const sigfig = itemJson.item.significantDigits;
var maxCount = 0; 
var list = [];
var avg = 0;

try {
    list = populateList(formJson, itemList, list);

    avg = calculateAverage(list, sigfig);
    logger("List: " + list);
    logger("List length: " + list.length)
    logger("Max count: " + maxCount);
    logger("Average: " + avg)
    if (list.length === maxCount) return (avg).toFixed(sigfig);
    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Collects numeric values from form items matching target names, stopping when an attached item is encountered (first-to-last order).
function populateList(form, targetItem, ilist) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
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
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);
    return Math.round(avg * factor) / factor;
}const formName = [
    "❤️ SCREEN VITALS (SUPINE) (BP,HR,RR,TEMP) V1.0",
]
const studyevent = [
    "SCREENING",
    "Screening"
]

const itemName = [
    "Confirm_Arm.",
]

const attachedItemCodeList = [
    "R",
    "L",
]

const pulledItemCodeList = [
    "R",
    "L",
]

try {
    var form = pullForm(studyevent, formName);

    if (!form) return null;

    var arm = pullItemFromForm(form, itemName);
    if (arm == pulledItemCodeList[0]) return attachedItemCodeList[0];
    else if (arm == pulledItemCodeList[1]) return attachedItemCodeList[1];

    return arm;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
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
