// Add item names
const dbpItems = [
    "Diastolic BP", 
    "SCRN_Diastolic BP",
    "DIA", 
    "Repeat DIA (2 of 3)",
    "Repeat DIA (3 of 3)",
    "DIA (1 of 3)",
    "DIA (2 of 3)",
    "DIA (3 of 3)",
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
    "SYS (1 of 3)",
    "SYS (2 of 3)",
    "SYS (3 of 3)",
];

const sysAttachedItem = [
    "Average SBP",
    "🧮 SYS_AVERAGE:",
];

const hrItem = [
    "HR (1 of 3)",
    "HR (2 of 3)",
    "HR (3 of 3)"
]

const hrAttachedItem = [
    "🧮 HR AVERAGE:"
]

const repeatItem = [
    "▶ Is VS repeat required?",    
]

// Inclusive (Edit)
const sysAvg_maxRange = 170;
const diaAvg_maxRange = 100;

// ======== Don't modify ========
var item = itemJson.item;
const sigfig = itemJson.item.significantDigits;

var repeat = pullItemFromForm(formJson, repeatItem);
logger("Repeat: " + repeat);
var isRepeat = containsValue(repeat, "yes");

if (isRepeat) {
    var syslist = populateListRepeat(formJson, sysItem, sysAttachedItem)
    var dialist = populateListRepeat(formJson, dbpItems, dbpAttachedItem)
}
else {
    var syslist = populateListNonRepeat(formJson, sysItem, sysAttachedItem)
    var dialist = populateListNonRepeat(formJson, dbpItems, dbpAttachedItem)
}

var sysAvg = calculateAverage(syslist, sigfig);
var diaAvg = calculateAverage(dialist, sigfig);

logger("Sys: " + sysAvg);
logger("Dia: " + diaAvg);

if (sysAvg >= sysAvg_maxRange || diaAvg >= diaAvg_maxRange) return "Yes";

return "No";

function populateList(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                }
            }
        }
    }
    return list;
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function populateListNonRepeat(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var maxCount = 3;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                    if (list.length >= 3) return list;
                }
            }
        }
    }
    return list;
}


function populateListRepeat(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var maxCount = 3;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                    if (list.length >= 3) return list;
                }
            }
        }
    }
    return list;
}