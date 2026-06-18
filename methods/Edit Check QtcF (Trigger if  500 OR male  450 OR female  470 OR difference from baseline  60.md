// Version: v1
// Purpose: QTcF safety edit check with sex-specific thresholds.

// Add item names
const baselineForm = [
    "⚡DAY-1 ECG SINGLE 12 LEAD V1",
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)",

]
const baselineEvent = [
    "Day -1",
    "Visit 2 Week 1 Day 0",
]
const baselineItem = [
    "Fridericia QTc Interval.",
    "Fridericia QTc Interval",
    "Baseline QTcF",
    "QTcF.",
];

const maxRange = 500;
const maleRange = 450;
const femaleRange = 470;
const differenceRange = 60;

// ======== Don't modify ========
const item = itemJson.item;
const sexMale = formJson.form.subject.volunteer.sexMale;
var diff = 0;

try {
    logger("Is it male: " + sexMale);
    logger("Qtcf value: " + item.value);
    if (item.value > maxRange) return log("500", item.value);
    else if (sexMale && item.value > maleRange) return log("male", item.value);
    else if (!sexMale && item.value > femaleRange) return log("female", item.value);

    var form = pullForm(baselineEvent, baselineForm);
    if (!form) return null;
    var baseline = pullItemFromForm(form, baselineItem);
    logger("Baseline: " + baseline);

    if (checkBaseline(itemJson.item.value)) return log("base", diff);

    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Logs the current values of sys, dia, and hr variables for debugging purposes.
function log(type, val)  {
    if (type == "500") {
        customErrorMessage("OOR > " + maxRange + ". REPEAT 2 TIMES: " + val);
        return false;
    }
    else if (type == "male") {
        customErrorMessage("Male QTcF OOR > " + maleRange + ": " + val);
        return false;
    }
    else if (type == "female") {
        customErrorMessage("Female QTcF OOR > " + femaleRange + ": " + val);
        return false;
    }
    else if (type == "base") {
        customErrorMessage("Difference From Baseline OOR > " + differenceRange + ". REPEAT 2 TIMES: " + val);
        return false;
    }
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

// Evaluates whether: checkBaseline.
function checkBaseline(itemValue) {
    if (baseline !== null && baseline !== undefined) {
        diff = (itemValue - baseline)
        logger("Difference: " + diff)
        if (diff >= differenceRange) {
            return true;
        }
        return false;
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
