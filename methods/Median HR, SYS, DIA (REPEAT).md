const sysItem = [
    "Systolic BP", 
    "SCRN_Systolic BP", 
    "ET Threshold SBP",
    "SYS (I: 90-150)",
];

const sysAttachedItem = [
    "Average SBP",
    "SYS Median (SF: >150)",
];

const diaItem = [
    "DIA (I: 50-99)",
    "DIA (I: 50-90)"
];

const diaAttachedItem = [
    "DIA Median (SF: ≥100)",
];

const hrItem = [
    "HR (I: 45-100)",
];

const HRAttachedItem = [
    "HR Median (SF: >100)",
];

logger("Current item: " + itemJson.item.name);
var list = [];
var median = 0;

var sysMaxCount = 6;
var diaMaxCount = 6;
var hrMaxCount = 6;
const item = itemJson.item;
const sigfig = item.significantDigits;

if (HRAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, hrItem, HRAttachedItem, hrMaxCount)) {
    list = filterList(list);
    median = calculateMedian(list, sigfig);
    log()
    return (median).toFixed(sigfig);
}

if (sysAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, sysItem, sysAttachedItem, sysMaxCount)) {
    list = filterList(list);
    median = calculateMedian(list, sigfig);
    log()
    return (median).toFixed(sigfig);
}

if (diaAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, diaItem, diaAttachedItem, diaMaxCount)) {
    list = filterList(list);
    median = calculateMedian(list, sigfig);
    log()
    return (median).toFixed(sigfig);
}

return null;

function log() {
    logger("List: " + list);
    logger("List length: " + list.length)
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

function filterList(list) {
    if (!list || list.length <= 3) return list;
    return list.slice(-3);
}


function populateList(form, targetItem, attachedItem, maxCount) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    list = [];
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                    logger("List length: " + list.length);
                    if (list.length >= maxCount) return list;
                }
            }
        }
    }
    if (list.length == 0) return false;
    return list;
}