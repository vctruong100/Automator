/* jshint strict: false */

// Version: v2
// Purpose: Protocol eligibility check using median HR/PR and average QTcF. Does not require any item names.

var form = formJson.form;
var attachedItem = itemJson.item;
var sigfig = attachedItem.significantDigits;
var isMale = formJson.form.subject.volunteer.sexMale;

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
    if (metric === "PR") return containsStandaloneKeyword(name, "PR");
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTC");

    return false;
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

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

try {
    logger("Attached item: " + attachedItem.name);

    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    var hrList = populateList(formJson, "HR", attachedItem.name, isRepeat);
    var prList = populateList(formJson, "PR", attachedItem.name, isRepeat);
    var qtcfList = populateList(formJson, "QTCF", attachedItem.name, isRepeat);

    var hrMedian = calculateMedian(hrList, sigfig);
    var prMedian = calculateMedian(prList, sigfig);
    var qtcfAverage = calculateAverage(qtcfList);

    logger("HR list: " + hrList);
    logger("PR list: " + prList);
    logger("QTcF list: " + qtcfList);
    logger("HR median: " + hrMedian);
    logger("PR median: " + prMedian);
    logger("QTcF average: " + qtcfAverage);

    if (containsValue(attachedItem.name, "inc001")) {
        if (prMedian < 220 && hrMedian > 60) return "YES";
        return "NO, SF";
    }
    else if (containsValue(attachedItem.name, "exc022")) {
        if ((isMale && qtcfAverage > 470) || (!isMale && qtcfAverage > 480)) return "YES, SF";
        else return "NO";
    }
    return null;

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
