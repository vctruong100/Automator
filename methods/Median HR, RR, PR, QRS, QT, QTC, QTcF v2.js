/* jshint strict: false */

// Version: v3
// Description: Calculates median ECG or vital values from the attached median item, auto-detecting HR, RR, PR, QRS, QT, QTC, QTcF, SYS, or DIA source values across repeat and non-repeat item groups and logging collected values.

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

    if (matchesMetric(name, "QTCF")) return "QTCF";
    if (matchesMetric(name, "QTC")) return "QTC";
    if (matchesMetric(name, "QT")) return "QT";
    if (matchesMetric(name, "QRS")) return "QRS";
    if (matchesMetric(name, "PR")) return "PR";
    if (matchesMetric(name, "RR")) return "RR";
    if (matchesMetric(name, "SYS")) return "SYS";
    if (matchesMetric(name, "DIA")) return "DIA";
    if (matchesMetric(name, "HR") || containsValue(name, "PULSE")) return "HR";

    return null;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;
    if (containsValue(itemName, "MEDIAN")) return true;

    return false;
}

function isSummaryItemForMetric(itemName, metric) {
    return isAverageItem(itemName) && matchesMetric(itemName, metric);
}

function addNumericValue(list, details, groupItem) {
    var value = groupItem ? groupItem.value : null;
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) {
        list.push(numericValue);
        details.push(groupItem.name + "=" + numericValue);
    }
}

function populateList(formJsonValue, metric, attachedItemName, isRepeat) {
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var details = [];
    var group, items, groupItem, i, j;

    if (!metric || !itemGroups || itemGroups.length < 1) return { values: list, details: details };

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = items.length - 1; j >= 0; j--) {
                groupItem = items[j];
                if (!groupItem) continue;
                if ((groupItem.name == attachedItemName || isSummaryItemForMetric(groupItem.name, metric)) && list.length > 1) return { values: list, details: details };
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, details, groupItem);
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
                if ((groupItem.name == attachedItemName || isSummaryItemForMetric(groupItem.name, metric))) return { values: list, details: details };
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, details, groupItem);
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                }
            }
        }
    }

    return { values: list, details: details };
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

    if (sigfig === null || sigfig === undefined || sigfig === "") sigfig = 0;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var metric = getMetricFromAverageItem(attachedItem.name);
    logger("Attached Item name: " + attachedItem.name);
    logger("Metric: " + metric);

    if (!metric) return null;

    var collected = populateList(formJson, metric, attachedItem.name, isRepeat);
    var list = collected.values;

    logger(metric + " collected values: " + list.join(", "));
    logger(metric + " collected item/value list: " + collected.details.join(" | "));

    var median = calculateMedian(list, sigfig);
    logger(metric + " median: " + median);

    if (median === null) return null;
    return median.toFixed(sigfig);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
