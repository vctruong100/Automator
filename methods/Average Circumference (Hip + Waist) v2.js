/* jshint strict: false */

// Version: v2
// Purpose: Computes sum of average hip and waist circumference values. Does not require any item names.

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

    if (metric === "HIP") return containsStandaloneKeyword(name, "HIP");
    if (metric === "WAIST") return containsStandaloneKeyword(name, "WAIST");

    return false;
}

function getMetricFromAverageItem(itemName) {
    var name = normalizeName(itemName);

    if (containsValue(name, "HIP")) return "HIP";
    if (containsValue(name, "WAIST")) return "WAIST";

    return null;
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

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItemName) return list;
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
                if (groupItem.name == attachedItemName && list.length > 1) return list;

                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    }

    return list;
}

function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    var avg = sum / values.length;
    var factor = Math.pow(10, sigfig);
    return (Math.round(avg * factor) / factor).toFixed(sigfig);
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var metric = getMetricFromAverageItem(attachedItem.name);
    logger("Attached Item name: " + attachedItem.name);
    logger("Metric: " + metric);

    var hipList = populateList(formJson, "HIP", attachedItem.name, isRepeat);
    var waistList = populateList(formJson, "WAIST", attachedItem.name, isRepeat);

    logger("Hip List: " + hipList);
    logger("Waist List: " + waistList);

    logger("Hip List length: " + hipList.length);
    logger("Waist List length: " + waistList.length);

    var hipAvg = calculateAverage(hipList, attachedItem.significantDigits);
    var waistAvg = calculateAverage(waistList, attachedItem.significantDigits);

    logger("Hip Average: " + hipAvg);
    logger("Waist Average: " + waistAvg);

    if (metric === "HIP") return hipAvg;
    if (metric === "WAIST") return waistAvg;

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
