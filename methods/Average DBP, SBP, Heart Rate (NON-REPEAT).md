const dbpItems = [
    "Diastolic BP", 
    "SCRN_Diastolic BP",
    "DIA", 
    "Repeat DIA (2 of 3)",
    "Repeat DIA (3 of 3)"
];

const dbpAttachedItem = [
    "Average DBP",
    "🧮 DIA AVERAGE:",
];

const sysItem = [
    "Systolic BP", 
    "SCRN_Systolic BP", 
    "ET Threshold SBP",
    "SYS", 
    "Repeat SYS (2 of 3)",
    "Repeat SYS (3 of 3)"
];

const sysAttachedItem = [
    "Average SBP",
    "🧮 SYS_AVERAGE:",
];

var item = itemJson.item;
const sigfig = itemJson.item.significantDigits;
var maxCount = 0; 
var list = [];
var avg = 0;

logger("Attached item: " + item.name);
if (sysAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, sysItem, sysAttachedItem)) {
    avg = calculateAverage(list, sigfig);
    if (list.length == maxCount) return (avg).toFixed(sigfig);
}

if (dbpAttachedItem.indexOf(item.name.trim()) !== -1 && populateList(formJson, dbpItems, dbpAttachedItem)) {
    avg = calculateAverage(list, sigfig);
    if (list.length == maxCount) return (avg).toFixed(sigfig);
}

return null;

function log() {
    logger("List: " + list);
    logger("List length: " + list.length);
    logger("Max count: " + maxCount);
    logger("Average: " + avg);
}

function populateList(form, targetItem, list) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                }
            }
        }
    }
    return list;
}

function calculateAverage(values, sigfig) {
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
    var factor = Math.pow(10, sigfig);
    return Math.round(avg * factor) / factor;
}