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

var list = [];
var maxCount = 0;
var median = 0;

var item = itemJson.item;
var sigfig = item.significantDigits;

try {
    logger("Attached item: " + item.name);
    if (HRAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, HRItem, HRAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    if (RRAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, RRItem, RRAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    if (PRAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, PRItem, PRAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    logger(QRSAttachedItem.indexOf(item.name.trim()) !== -1)
    if (QRSAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, QRSItem, QRSAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    if (QTAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, QTItem, QTAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    if (QTCAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, QTCItem, QTCAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    if (QTcFAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, QTcFitem, QTcFAttachedItem)) {
        median = calculateMedian(list, sigfig);
        if (list.length == maxCount) return (median).toFixed(sigfig);
    }

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function log() {
    logger("List: " + list);
    logger("List length: " + list.length)
    logger("Max count: " + maxCount);
    logger("Median: " + median)
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

function populateList(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    list = [];
    maxCount = 0;
	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name.trim()) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                    logger("Item name: " + item.name + ", value: " + item.value);
                }
            }
        }
    }
    if (list.length == 0) return false;
    return list;
}
