/* jshint strict: false */

// Version: v2
// Purpose: Computes median values for full ECG/vitals panel (non-repeat or repeat). Does not require any item names.

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

    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "RATE") || containsValue(name, "HEART RATE");
    if (metric === "RR") return containsStandaloneKeyword(name, "RR") || containsValue(name, "RESPIRATORY RATE");
    if (metric === "PR") return containsStandaloneKeyword(name, "PR");
    if (metric === "QRS") return containsStandaloneKeyword(name, "QRS");
    if (metric === "QT") return name.indexOf("QTC") === -1 && containsStandaloneKeyword(name, "QT");
    if (metric === "QTC") return name.indexOf("QTCF") === -1 && containsStandaloneKeyword(name, "QTC");
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTCF");

    return false;
}

function getMetricFromAverageItem(itemName) {
    var name = normalizeName(itemName);

    if (containsValue(name, "QTCF")) return "QTCF";
    if (containsValue(name, "QTC") && !containsValue(name, "QTCF")) return "QTC";
    if (containsValue(name, "QT") && !containsValue(name, "QTC")) return "QT";
    if (containsValue(name, "QRS")) return "QRS";
    if (containsValue(name, "PR")) return "PR";
    if (containsValue(name, "RR")) return "RR";
    if (containsValue(name, "HR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE")) return "HR";

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
                if ((groupItem.name == attachedItemName || isAverageItem(groupItem.name))) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    }

    return list;
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

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var metric = getMetricFromAverageItem(attachedItem.name);
    logger("Attached Item name: " + attachedItem.name);
    logger("Metric: " + metric);

    var list = populateList(formJson, metric, attachedItem.name, isRepeat);

    logger(metric + " list: " + list);

    var median = calculateMedian(list, sigfig);
    logger(metric + " median: " + median);

    return median.toFixed(sigfig);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
