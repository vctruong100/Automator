/* jshint strict: false */

// Version: v2
// Purpose: Protocol range check for averaged SBP and DBP. Does not require any item names.

// Inclusive
var sysAvg_maxRange = 170;
var diaAvg_maxRange = 100;

var form = formJson.form;
var attachedItem = itemJson.item;
var sigfig = attachedItem.significantDigits;

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

    return false;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;

    return false;
}

function addNumericValue(list, value) {
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) list.push(numericValue);
}
function populateList(formJsonValue, metric, attachedItemName, isRepeat) {
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
                if ((groupItem.name == attachedItemName || isAverageItem(groupItem.name)) && list.length > 1) return list;
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
                if (groupItem.name == attachedItemName || isAverageItem(groupItem.name)) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    }

    return list;
}
function calculateAverage(values, sigfig) {
    if (!values || values.length === 0) return null;

    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (!isNaN(values[i])) {
            sum += values[i];
            count++;
        }
    }

    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);

    return Math.round(avg * factor) / factor;
}

function pullItemFromForm(formJsonValue, targetItem) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && containsValue(item.name, targetItem) && item.value !== null && item.value !== "" && !item.canceled) {
                return item.value;
            }
        }
    }
    return null;
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = false;
    if (containsValue(parsedGroupName, "inclusion") || containsValue(parsedGroupName, "exclusion")) isRepeat = true;
    else isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    var sysList = populateList(formJson, "SYS", attachedItem.name, isRepeat);
    var diaList = populateList(formJson, "DIA", attachedItem.name, isRepeat);

    var sysAvg = calculateAverage(sysList, sigfig);
    var diaAvg = calculateAverage(diaList, sigfig);

    logger("Sys Avg: " + sysAvg);
    logger("Dia Avg: " + diaAvg);

    if (
        sysAvg !== null &&
        diaAvg !== null &&
        (sysAvg >= sysAvg_maxRange || diaAvg >= diaAvg_maxRange)
    ) {
        return attachedItem.codeListItems[1].codedValue; // Yes
    }

    return attachedItem.codeListItems[2].codedValue; // No
}
catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
