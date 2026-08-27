/* jshint strict: false */

// Version: v3
var baselineStudyEvent = [
    "P1 D1 (PRE)",
    "Part 2 D-1"
];
var baselineFormName = [
    "⚡ ECG 12-LEAD (TRIPLICATE)",
    "⚡ ECG 12-LEAD (TRIPLICATE) SCRN/BASELINE",
    "⚡ ECG 12-LEAD (TRIPLICATE) SCRN/BASELINE V2",
    "⚡ ECG 12-LEAD (TRIPLICATE) SCRN/BASELINE V3",
    "⚡ ECG 12-LEAD (TRIPLICATE) SCRN/BASELINE V4"
];

var attachedItem = itemJson.item;
var item = itemJson.item;

var rawgroupName = getItemDataContextByItemDataId(item.id);
var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
logger("Group name: " + parsedGroupName);

var form = pullForm(baselineStudyEvent, baselineFormName);

if (!form) return null;
logger(form.form.name)

var qtcfBaseline = pullItemOnKeyword(form, "QTCF");
var PRbaseline = pullItemOnKeyword(form, "PR");
var qrsBaseline = pullItemOnKeyword(form, "QRS");

logger("QTCF baseline: " + qtcfBaseline);
logger("PR Baseline: " + PRbaseline);
logger("QRS baseline: " + qrsBaseline)

var qtcfNew = pullItemFromSameItemGroup(formJson, "QTCF", parsedGroupName);
var prNew = pullItemFromSameItemGroup(formJson, "PR", parsedGroupName);
var qrsNew = pullItemFromSameItemGroup(formJson, "QRS", parsedGroupName);

logger("QTCF New: " + qtcfNew);
logger("PR New: " + prNew);
logger("QRS New: " + qrsNew);

var qtcfDiff = qtcfNew - qtcfBaseline;
var prDiff = prNew - PRbaseline;
var qrsDiff = qrsNew - qrsBaseline;

logger("QTCF Difference: " + qtcfDiff);
logger("PR Difference: " + prDiff);
logger("QRS Differnce: " + qrsDiff);

return "QTCF: " + qtcfDiff + ", PR: " + prDiff + ", QRS: "  + qrsDiff;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function containsStandaloneKeyword(input, keyword) {
    var value = normalizeName(input);
    var target = normalizeName(keyword);
    var startIndex = 0;
    var index;
    var before;
    var after;

    while (startIndex < value.length) {
        index = value.indexOf(target, startIndex);
        if (index === -1) return false;

        before = index === 0 ? "" : value.charAt(index - 1);
        after = index + target.length >= value.length ? "" : value.charAt(index + target.length);

        if ((before === "" || !/[A-Z0-9]/.test(before)) && (after === "" || !/[A-Z0-9]/.test(after))) return true;

        startIndex = index + target.length;
    }

    return false;
}

function matchesMetric(itemName, metric) {
    var name = normalizeName(itemName);
    if (metric === "PR") return name.indexOf("PR") !== -1 || containsStandaloneKeyword(name, "PR");
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTCF");
    if (metric === "QRS") return name.indexOf("QRS") !== -1 || containsStandaloneKeyword(name, "QRS");

    return false;
}

function pullItemFromSameItemGroup(formJsonValue, metric, groupname) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;

        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (matchesMetric(groupItem.name, metric) && groupItem.value !== null) {
                logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                return groupItem.value;
            }
        }
    }

    return null;
}

function pullItemOnKeyword(formJsonValue, metric) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;

        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (matchesMetric(groupItem.name, metric) && groupItem.value !== null) {
                logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                return groupItem.value;
            }
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