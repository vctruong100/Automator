/* jshint strict: false */

// Version: v2
// Purpose: Protocol eligibility check using median SYS, DIA, and HR. Does not require any item names.

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
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "RATE") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE");

    return false;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;
    if (containsValue(itemName, "MEDIAN")) return true;

    return false;
}

function isRepeatRequiredItem(itemName) {
    return containsValue(itemName, "repeat") && containsValue(itemName, "required");
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
                if (groupItem.name == attachedItemName && list.length > 1) return list;
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
                if (groupItem.name == attachedItemName) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    }

    return list;
}

function calculateMedian(values) {
    if (values.length === 0) return null;
    values.sort(function(a, b) { return a - b; });

    var mid = Math.floor(values.length / 2);
    var median;

    if (values.length % 2 === 0) {
        median = (values[mid - 1] + values[mid]) / 2;
    } else {
        median = values[mid];
    }

    return Math.round(median);
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
            if (containsValue(item.name, targetItem) && item.value !== null && !item.canceled && item.value !== "") {
                return item;
            }
        }
    }
    return null;
}

try {
    var isRepeat = false;
    var list = [];

    var repeatItem = pullItemFromForm(formJson, "repeat required");
    if (!repeatItem || repeatItem.value == null) return null;

    if (repeatItem.value == repeatItem.codeListItems[1].codedValue) isRepeat = true;

    logger("Attached item: " + attachedItem.name);
    var sysList = populateList(formJson, "SYS", attachedItem.name, isRepeat);
    var diaList = populateList(formJson, "DIA", attachedItem.name, isRepeat);
    var hrList = populateList(formJson, "HR", attachedItem.name, isRepeat);

    logger("Sys list: " + sysList);
    logger("Dia list: " + diaList);
    logger("Hr list: " + hrList);

    var sysMedian = calculateMedian(sysList);
    var diaMedian = calculateMedian(diaList);
    var hrMedian = calculateMedian(hrList);

    logger("Sys Median: " + sysMedian);
    logger("Dia Median: " + diaMedian);
    logger("Hr Median: " + hrMedian);

    if (containsValue(attachedItem.name, "exc004")) {
        if (sysMedian > 150 || diaMedian >= 100) return "YES, SF";
        return "NO";
    }
    if (containsValue(attachedItem.name, "exc005")) {
        if (hrMedian > 100) return "YES, SF";
        return "NO";
    }

    if (sysMedian > 90 && hrMedian > 60 && diaMedian > 50) return "YES";
    else if (sysMedian <= 90 || hrMedian <= 60 || diaMedian <= 50) return "NO, SF";
    return null;

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
