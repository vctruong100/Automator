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

var item = itemJson.item;

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
