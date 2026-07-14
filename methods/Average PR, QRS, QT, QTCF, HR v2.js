/* jshint strict: false */
// Purpose: Computes average ECG values. Uses keyword in item name (instead of full name) for pulling values. Useful if range changes often

var PRattach = "PR AVERAGE";
var QRSattach = "QRS AVERAGE";
var QTcFattach = "QTc AVERAGE";
var QTattach = "QT AVERAGE";
var HRattach = "HR AVERAGE";

var currentItem = itemJson.item;

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

    if (metric === "PR") return containsStandaloneKeyword(name, "PR");
    if (metric === "QRS") return containsStandaloneKeyword(name, "QRS");
    if (metric === "QT") return name.indexOf("QTC") === -1 && containsStandaloneKeyword(name, "QT");
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTC");
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "RATE");

    return false;
}

function isAttachedItem(itemName, attachedItem) {
    return containsValue(itemName, attachedItem);
}

function calculateAverage(values) {
    var sum = 0;
    var i;

    if (!values || values.length === 0) return null;

    for (i = 0; i < values.length; i++) sum += values[i];

    return Math.round(sum / values.length).toString();
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
                if (isAttachedItem(groupItem.name, attachedItem)) return list;

                if (matchesMetric(groupItem.name, metric) && !containsValue(groupItem.name, "average")) {
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
                if (isAttachedItem(groupItem.name, attachedItem) && list.length > 1) return list;

                if (matchesMetric(groupItem.name, metric) && !containsValue(groupItem.name, "average")) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    }

    return list;
}

try {
    var rawGroupName = getItemDataContextByItemDataId(currentItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var PRlist = populateList(formJson, "PR", PRattach, isRepeat);
    var QRSlist = populateList(formJson, "QRS", QRSattach, isRepeat);
    var QTlist = populateList(formJson, "QT", QTattach, isRepeat);
    var QTcFlist = populateList(formJson, "QTCF", QTcFattach, isRepeat);
    var HRlist = populateList(formJson, "HR", HRattach, isRepeat);

    logger("PR list: " + PRlist);
    logger("QRS list: " + QRSlist);
    logger("QT list: " + QTlist);
    logger("QTcF list: " + QTcFlist);
    logger("HR list: " + HRlist);

    var avgPR = calculateAverage(PRlist);
    var avgQRS = calculateAverage(QRSlist);
    var avgQT = calculateAverage(QTlist);
    var avgQTcF = calculateAverage(QTcFlist);
    var avgHR = calculateAverage(HRlist);

    logger("Average PR: " + avgPR);
    logger("Average QRS: " + avgQRS);
    logger("Average QT: " + avgQT);
    logger("Average QTcF: " + avgQTcF);
    logger("Average HR: " + avgHR);

    if (isAttachedItem(currentItem.name, PRattach)) return avgPR;
    if (isAttachedItem(currentItem.name, QRSattach)) return avgQRS;
    if (isAttachedItem(currentItem.name, QTcFattach)) return avgQTcF;
    if (isAttachedItem(currentItem.name, QTattach)) return avgQT;
    if (isAttachedItem(currentItem.name, HRattach)) return avgHR;

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}