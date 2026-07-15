/* jshint strict: false */

// Version: v2
// Purpose: Protocol check for averaged QTcF/QRS out-of-range. Does not require any item names.

// Inclusive (Edit)
var QTcF_max_range = 450;
var QRS_max_range = 120;

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

    if (metric === "QRS") return containsStandaloneKeyword(name, "QRS");
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTC");

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
                    addNumericValue(list, groupItem.value);
                    if (list.length >= 3) return list;
                }
            }
        }
    }
    else {
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
                    if (list.length >= 3) return list;
                }
            }
        }
    }

    return list;
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    return avg;
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    var QTcFlist = populateList(formJson, "QTCF", attachedItem.name, isRepeat);
    var QRSlist = populateList(formJson, "QRS", attachedItem.name, isRepeat);

    var QTcFavg = calculateAverage(QTcFlist);
    var QRSavg = calculateAverage(QRSlist);

    logger("QTcF list: " + QTcFlist);
    logger("QRS list: " + QRSlist);
    logger("QTcF average: " + QTcFavg);
    logger("QRS average: " + QRSavg);

    if (QTcFlist.length < 3 || QRSlist.length < 3) return attachedItem.codeListItems[0].codedValue; // return pending result

    if (QTcFavg > QTcF_max_range || QRSavg > QRS_max_range) return attachedItem.codeListItems[2].codedValue; // return Out of protocol range
    else if (QTcFavg <= QTcF_max_range || QRSavg > QRS_max_range) return attachedItem.codeListItems[1].codedValue; // return within protocol range

    return attachedItem.codeListItems[0].codedValue;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
