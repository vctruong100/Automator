/* jshint strict: false */

// Version: v1
// Purpose: Computes median values for full ECG/vitals panel (non-repeat).
var HRItem = [
    "HR (P: 40-100)",
]

var RRItem = [
    "RR (I: 600-1000 ms)",
]

var PRItem = [
    "PR (P: <230)",
]

var QRSItem = [
    "QRS (P: <130)",
]

var QTItem = [
    "QT (I: ≤500 ms)",
]

var QTCItem = [
    "QTc (P: Male < 470 ms / Female < 480 ms)",
]
var QTcFitem = [
    "QTcF (P: Male < 470 ms / Female < 480 ms)",
]
var HRAttachedItem = [
    "HR (P: 40-100) - Median"
]
var RRAttachedItem = [
    "RR (I: 600-1000) - Median",
]
var PRAttachedItem = [
    "PR (P: <230) - Median",
]
var QTAttachedItem = [
    "QT (I: ≤500 ms) - Median",
]
var QRSAttachedItem = [
    "QRS (P: <130) - Median",
]
var QTCAttachedItem = [
    "QTC (P: Male < 470 ms / Female < 480 ms) - Median",
]
var QTcFAttachedItem = [
    "QTcF (P: Male < 470 ms / Female < 480 ms) - Median",
]

var item = itemJson.item;
var sigfig = item.significantDigits;

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
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

function populateList(form, targetItem, attachedItem, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    var list = [];
	if (!itemGroups || itemGroups.length < 1) return null;

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name.trim())) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        logger("Item name: " + item.name + ", value: " + item.value);
                    }
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name.trim())) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        logger("Item name: " + item.name + ", value: " + item.value);
                    }
                }
            }
        }
    }
    return list;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

try {
    var isRepeat = false;
    var list = [];

    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
    logger("Group name: " + parsedGroupName);
    if (containsValue(parsedGroupName, "repeat")) isRepeat = true;

    logger("Attached item: " + item.name);
    if (containsItemName(HRAttachedItem, item.name.trim())) list = populateList(formJson, HRItem, HRAttachedItem, isRepeat);
    if (containsItemName(RRAttachedItem, item.name.trim())) list = populateList(formJson, RRItem, RRAttachedItem, isRepeat);
    if (containsItemName(PRAttachedItem, item.name.trim())) list = populateList(formJson, PRItem, PRAttachedItem, isRepeat);
    if (containsItemName(QRSAttachedItem, item.name.trim())) list = populateList(formJson, QRSItem, QRSAttachedItem, isRepeat);
    if (containsItemName(QTAttachedItem, item.name.trim())) list = populateList(formJson, QTItem, QTAttachedItem, isRepeat);
    if (containsItemName(QTCAttachedItem, item.name.trim())) list = populateList(formJson, QTCItem, QTCAttachedItem, isRepeat);
    if (containsItemName(QTcFAttachedItem, item.name.trim())) list = populateList(formJson, QTcFitem, QTcFAttachedItem, isRepeat)

    var median = calculateMedian(list, sigfig);
    return (median).toFixed(sigfig);

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
