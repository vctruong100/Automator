/* jshint strict: false */

// Version: v2
// Description: Runs a QTcF protocol check using keyword-based QTcF/Fridericia detection instead of exact item names. Pulls baseline QTcF from configured baseline forms and compares it against the latest matching QTcF value on the current form, flagging when the increase is at least 60 msec or the current QTcF is above 500 msec.

var baselineForms = [
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)"
];
var baselineFormStudyEvents = [
    "Visit 2 Week 1 Day 0"
];

var qtcfMaxRange = 500;
var differenceRange = 60;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function isQtcfItem(itemName) {
    var name = normalizeName(itemName);
    return containsValue(name, "QTCF") ||
        (containsValue(name, "FRIDERICIA") && containsValue(name, "QTC")) ||
        containsValue(name, "QT CORRECTED BY FRIDERICIA");
}

function addNumericValue(list, sourceItem) {
    if (!sourceItem || sourceItem.canceled || sourceItem.value === null || sourceItem.value === undefined || sourceItem.value === "") return;

    var numericValue = Number(sourceItem.value);
    if (!isNaN(numericValue)) {
        list.push({
            name: sourceItem.name,
            value: numericValue
        });
    }
}

function pullLatestQtcfFromForm(formJsonValue) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;
    var matches = [];

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;
        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (isQtcfItem(groupItem.name)) addNumericValue(matches, groupItem);
        }
    }

    if (matches.length < 1) return null;

    var latest = matches[matches.length - 1];
    logger("QTcF keyword matches: " + matches.map(function(match) { return match.name + "=" + match.value; }).join(", "));
    logger("Latest QTcF value used: " + latest.value + " from " + latest.name);
    return latest.value;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
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
        }
    }
    return keepers;
}

try {
    var baselineForm = pullForm(baselineFormStudyEvents, baselineForms);
    if (!baselineForm) return true;

    var baseline = pullLatestQtcfFromForm(baselineForm);
    var qtcf = pullLatestQtcfFromForm(formJson);

    logger("Baseline QTcF: " + baseline);
    logger("Current QTcF: " + qtcf);

    if (baseline === null || baseline === undefined || isNaN(baseline) || qtcf === null || qtcf === undefined || isNaN(qtcf)) return null;

    var diff = qtcf - baseline;
    logger("QTcF difference from baseline: " + diff);

    if (diff >= differenceRange || qtcf > qtcfMaxRange) return "Y";
    return "N";
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
