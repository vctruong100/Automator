/* jshint strict: false */

// Version: v4
// Purpose: Calculates the median SYS, DIA, and HR, then checks protocol ranges.
//          Combines Protocol Check Vital Signs Range v2 and Median SYS, DIA, and HR.

// Inclusive
var sys_min_range = 91;
var sys_max_range = 150;

var dia_min_range = 51;
var dia_max_range = 100;

var hr_min_range = 61;
var hr_max_range = 100;

var form = formJson.form;
var attachedItem = itemJson.item;
var sigfig = attachedItem.significantDigits != null ? attachedItem.significantDigits : 0;

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

    if (metric === "SYS") return containsStandaloneKeyword(name, "SYS") || containsStandaloneKeyword(name, "SBP") || containsValue(name, "SYSTOLIC");
    if (metric === "DIA") return containsStandaloneKeyword(name, "DIA") || containsStandaloneKeyword(name, "DBP") || containsValue(name, "DIASTOLIC");
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "PR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE");

    return false;
}

function getMetricFromAverageItem(itemName) {
    if (matchesMetric(itemName, "SYS")) return "SYS";
    if (matchesMetric(itemName, "DIA")) return "DIA";
    if (matchesMetric(itemName, "HR")) return "HR";

    return null;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;
    if (containsValue(itemName, "MEDIAN")) return true;

    return false;
}

function addNumericValue(list, value) {
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) list.push(numericValue);
}

function populateList(formJsonValue, metric, groupName, attachedItemName, isRepeat) {
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return list;

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = items.length - 1; j >= 0; j--) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (isAverageItem(groupItem.name) && list.length > 1) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    } else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (isAverageItem(groupItem.name)) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    }

    return list;
}

function pullItemFromForm(formJsonValue, metric, groupName, isRepeat) {
    logger("Target metric: " + metric);
    var itemGroups = formJsonValue.form.itemGroups;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;

        for (var j = group.items.length - 1; j >= 0; j--) {
            var item = group.items[j];
            if (!item) continue;

            if (matchesMetric(item.name, metric) && item.value !== null && item.value !== "") return item.value;
        }
    }

    return null;
}

function checkIfItemGroupHasAverageOrMedian(formJsonValue, groupName) {
    var itemGroups = formJsonValue.form.itemGroups;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;

        for (var j = group.items.length - 1; j >= 0; j--) {
            var item = group.items[j];
            if (!item) continue;

            if (isAverageItem(item.name) && item.value !== null && item.value !== "") return true;
        }
    }

    return false;
}

function calculateMedian(values, sigfig) {
    if (values.length === 0) return null;
    values.sort(function(a, b) { return a - b; });

    var mid = Math.floor(values.length / 2);
    var median;

    if (values.length % 2 === 0) {
        median = (values[mid - 1] + values[mid]) / 2;
    } else {
        median = values[mid];
    }

    var factor = Math.pow(10, sigfig);
    return Math.round(median * factor) / factor;
}

var sys, dia, hr;

var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

logger("Group name: " + parsedGroupName);
logger("Is repeat: " + isRepeat);

var doesItemGroupContainMedianAverage = checkIfItemGroupHasAverageOrMedian(formJson, parsedGroupName);
logger("Dose Item Group Contains Median Or Average: " + doesItemGroupContainMedianAverage)
if (doesItemGroupContainMedianAverage) {
    var sysList = populateList(formJson, "SYS", parsedGroupName, attachedItem.name, isRepeat);
    var diaList = populateList(formJson, "DIA", parsedGroupName, attachedItem.name, isRepeat);
    var hrList = populateList(formJson, "HR", parsedGroupName, attachedItem.name, isRepeat);
    
    logger("sysList: " + sysList);
    logger("diaList: " + diaList);
    logger("hrList: " + hrList);

    sys = calculateMedian(sysList, sigfig);
    dia = calculateMedian(diaList, sigfig);
    hr = calculateMedian(hrList, sigfig);
    
    logger("Systolic median: " + sys);
    logger("Diastolic median: " + dia);
    logger("Heart rate median: " + hr);
} else {
    sys = pullItemFromForm(formJson, "SYS", parsedGroupName, isRepeat);
    dia = pullItemFromForm(formJson, "DIA", parsedGroupName, isRepeat);
    hr = pullItemFromForm(formJson, "HR", parsedGroupName, isRepeat);
    
    logger("Systolic: " + sys);
    logger("Diastolic: " + dia);
    logger("Heart rate: " + hr);
}

if (!sys || sys === null || !dia || dia == null || !hr || hr == null) return attachedItem.codeListItems[4].codedValue;

// OOR
if (
    sys > sys_max_range ||
    sys < sys_min_range ||
    dia > dia_max_range ||
    dia < dia_min_range ||
    hr > hr_max_range ||
    hr < hr_min_range
) return attachedItem.codeListItems[1].codedValue; // Out of Protocol Range
else if ( // IR
    sys <= sys_max_range &&
    sys >= sys_min_range &&
    dia <= dia_max_range &&
    dia >= dia_min_range &&
    hr <= hr_max_range &&
    hr >= hr_min_range
) return attachedItem.codeListItems[0].codedValue; // Within Normal Range

return attachedItem.codeListItems[4].codedValue;

