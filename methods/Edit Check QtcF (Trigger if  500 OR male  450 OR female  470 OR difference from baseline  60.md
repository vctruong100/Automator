/* jshint strict: false */

// Version: v1
// Purpose: QTcF safety edit check with sex-specific thresholds.

// Add item names
var baselineForm = [
    "⚡DAY-1 ECG SINGLE 12 LEAD V1",
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)",

]
var baselineEvent = [
    "Day -1",
    "Visit 2 Week 1 Day 0",
]
var baselineItem = [
    "Fridericia QTc Interval.",
    "Fridericia QTc Interval",
    "Baseline QTcF",
    "QTcF.",
];

var maxRange = 500;
var maleRange = 450;
var femaleRange = 470;
var differenceRange = 60;

var item = itemJson.item;
var sexMale = formJson.form.subject.volunteer.sexMale;
var diff = 0;

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

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
