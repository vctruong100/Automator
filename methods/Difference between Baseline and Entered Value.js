/* jshint strict: false */

// Version: v1
// Purpose: Calculates delta between baseline and currently entered value.

// Add Item Names
var baselineStudyEvent = ["Period 1 Day 1","Visit 2 Week 1 Day 0",];
var baselineFormName = ["ECG_Single 12 - Lead ECG + Triplicate"];

var qtcfItem = ["Baseline QTcF"];

var sigfig = itemJson.item.significantDigits;

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function pullItemFromForm(form, targetItem, isBaseline) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    if (!isBaseline) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled) return parseInt(item.value);
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled) return parseInt(item.value);
            }
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
    var form = pullForm(baselineStudyEvent, baselineFormName);
    if (!form) return null;

    var baseline = pullItemFromForm(form, qtcfItem, true);
    var qtcfValue = pullItemFromForm(formJson, qtcfItem, false);
    if (!qtcfValue || qtcfValue == null || !baseline || baseline == null) return null;
    logger("qtcfValue: " + qtcfValue);
    logger("Baseline: " + baseline);

    var difference = Math.round(qtcfValue - baseline).toFixed(sigfig);
    logger("Difference: " + difference);
    return difference;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
