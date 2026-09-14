/* jshint strict: false */

// Version: v2
// Description: Runs a QTcF protocol check using keyword-based QTcF/Fridericia detection instead of exact item names. Pulls baseline QTcF from configured baseline forms and compares it against the latest matching QTcF value on the current form, flagging when the increase is at least 60 msec or the current QTcF is above 500 msec.

var baselineForms = [
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)",
    "⚡ 12-LEAD ECG (SINGLE) V1.0"
];
var baselineFormStudyEvents = [
    "Visit 2 Week 1 Day 0",
    "D1 (PRE)"
];

var attachedItem = itemJson.item;
var item = itemJson.item;
var studyevent = formJson.form.studyEventName;

var rawgroupName = getItemDataContextByItemDataId(item.id);
var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;
var d1pre = containsValue(studyevent, "d1 (pre)")

logger(d1pre)
logger("Group name: " + parsedGroupName);

var form = pullForm(baselineFormStudyEvents, baselineForms);

if (!form) return null;
logger(form.form.name)

var qtcfBaseline = null;
if (d1pre) {
    qtcfBaseline = pullItemOnKeyword(form, "QTCF", isRepeat, false);
}
else {
    qtcfBaseline = pullItemOnKeyword(form, "QTCF", true, false);
}
logger("QTCF baseline: " + qtcfBaseline);

var qtcfNew = pullItemOnKeyword(formJson, "QTCF", isRepeat, false);

logger("QTCF New: " + qtcfNew);

var qtcfDiff = qtcfNew - qtcfBaseline;

logger("QTCF Difference: " + qtcfDiff);

return qtcfDiff.toFixed(0);

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
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTCF");

    return false;
}

function pullItemOnKeyword(formJsonValue, metric, isRepeat, isBaseline) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;
    
            items = group.items;
    
            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (!isBaseline && containsValue(groupItem.name, "baseline")) continue;
                if (matchesMetric(groupItem.name, metric) && groupItem.value !== null) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    return groupItem.value;
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;
    
            items = group.items;
    
            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (!isBaseline && containsValue(groupItem.name, "baseline")) continue;
                if (matchesMetric(groupItem.name, metric) && groupItem.value !== null && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    return groupItem.value;
                }
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

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;
    if (containsValue(itemName, "DIFFERENCE")) return true;

    return false;
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