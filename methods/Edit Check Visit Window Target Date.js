/* jshint strict: false */

// Version: v1
// Purpose: Validates visit dates fall within protocol-defined windows.

// Add item names
var studyEventNames = [
    "Day 1 (pre)",
    "DAY 1 (Pre)",
    "V1 (Pre-Randomization)",
];
var formNames = [
    "RANDOMIZATION",
    "DISPOSITION RANDOMIZATION_V2.0",
    "DISPOSITION RANDOMIZATION",
    "RANDOMIZATION IMPALA (IRT)",
    "RANDOMIZATION IMPALA (IRT) V2.0",
];
var startDateItem = [
    "Date..",
    "Date of Randomization"
];

var map = {
    "V3 (D28 to D35)": 31,
    "V4 (D84 to D98)": 91,
    "V5 (D175 to D189)": 182,
    "V6 (Within 4 weeks)": 28
};

var range = {
    "V3 (D28 to D35)": 4,
    "V4 (D84 to D98)": 7,
    "V5 (D175 to D189)": 7,
    "V6 (Within 4 weeks)": 28
};

var studyEvent = formJson.form.studyEventName;
var item = itemJson.item;

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
        }
    }
    return false;
}
function collectCompleted(formDataArray) {
    if (formDataArray == null) return [];
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (
            formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (
                formData.form.dataCollectionStatus == 'Complete' ||
                formData.form.dataCollectionStatus == 'Incomplete' ||
                formData.form.dataCollectionStatus == 'Nonconformant'
            )
        ) {
            keepers.push(formData);
        }
    }
    return keepers;
}

function checkForm(form, studyevent) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm[completedForm.length - 1];
    }
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null) return item;
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function msToDate(ms) {
    var d = new Date(Number(ms));
    var day = d.getDate();
    if (day < 10) day = "0" + day;
    var months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    var month = months[d.getMonth()];
    var year = d.getFullYear();
    return day + " " + month + " " + year;
}

function isoToLocalMidnight(isoStr) {
    if (!isoStr) return null;
    var d = isoStr.split("T")[0];
    var p = d.split("-");
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getTime();
}

try {
    var form = pullForm(formNames, studyEventNames);
    var day1Date = pullItemFromForm(form, startDateItem);

    if (!day1Date || !item || item.value == null || item.dateValueMs == null) return true;

    var day1DateMs = isoToLocalMidnight(day1Date.value);
    var day1Dateformat = msToDate(day1DateMs);

    var addDays = map[studyEvent];
    var allowedRange = range[studyEvent];
    if (addDays == null || allowedRange == null) return true;

    var targetMs = day1DateMs + addDays * 86400000;
    var allowedRangeMs = allowedRange * 86400000;

    var collectedMs = isoToLocalMidnight(itemJson.item.value);
    var diffDays = Math.abs((collectedMs - targetMs) / 86400000);

    var collectedDateformat = msToDate(collectedMs);
    var targetDateformat = msToDate(targetMs);
    var minTargeDateformat = msToDate(targetMs - allowedRangeMs);
    var maxTargetDateformat = msToDate(targetMs + allowedRangeMs);

    if (diffDays > allowedRange) {
        customErrorMessage(
            "Target Date: " + targetDateformat +
            ", Allowed Range: ±" + allowedRange + " day(s)"
        );
        return false;
    }
    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
