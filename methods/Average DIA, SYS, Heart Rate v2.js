/* jshint strict: false */

// Version: v1
// Purpose: Calculates average diastolic, systolic BP, and heart rate. Does not require any item names.

var form = formJson.form;
var attachedItem = itemJson.item;

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
    var name = normalizeName(itemName);

    if (containsValue(name, "SYS") || containsValue(name, "SBP") || containsValue(name, "SYSTOLIC")) return "SYS";
    if (containsValue(name, "DIA") || containsValue(name, "DBP") || containsValue(name, "DIASTOLIC")) return "DIA";
    if (containsValue(name, "HR") || containsValue(name, "PR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE")) return "HR";

    return null;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;

    return false;
}

function isAttachedItem(itemName, attachedItem) {
    return containsValue(attachedItem, itemName);
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

function addNumericValue(list, value) {
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) list.push(numericValue);
}

function populateList(formJsonValue, metric, attachedItem, isRepeat) {
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return list;

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItem) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    } else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = items.length - 1; j >= 0; j--) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItem && list.length > 1) return list;

                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    }

    return list;
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var metric = getMetricFromAverageItem(attachedItem.name);
    logger("Attached Item name: " + attachedItem.name)
    logger("Metric:" + metric)
    var sysList = populateList(formJson, "SYS", attachedItem.name, isRepeat);
    var diaList = populateList(formJson, "DIA", attachedItem.name, isRepeat);
    var hrList = populateList(formJson, "HR", attachedItem.name, isRepeat);

    logger("Is it a repeat? " + isRepeat);
    logger("sysList: " + sysList);
    logger("diaList: " + diaList);
    logger("hrList: " + hrList);

    var avgSys = calculateAverage(sysList);
    var avgDia = calculateAverage(diaList);
    var avgHR = calculateAverage(hrList);

    logger("Average PR: " + avgSys);
    logger("AVerage QRS: " + avgDia);
    logger("Average QT: " + avgHR);

    if (metric === "SYS") return avgSys;
    if (metric === "DIA") return avgDia;
    if (metric === "HR") return avgHR;

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
