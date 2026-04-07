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

// Replace codelist 
const attachedItemCodeList = [
    "⭕Pending Results",
    "✅ Within Protocol Range",
    "🛑 Out of protocol range, SF",
    "❗Out of Normal Range",
    "✅ Within Normal Range",
    " Yes, within normal range",
    "No",
]

// Inclusive (edit)
var sys_min_range = 90;
var sys_max_range = 169;

var dia_min_range = 50
var dia_max_range = 99;

var hr_min_range = 50;
var hr_max_range = 100;

// ======== Don't modify ========
var item = itemJson.item;
const sigfig = itemJson.item.significantDigits;
var pending = false;

logger("Attached item: " + item.name);

var syslist = populateList(formJson, sysItem, sysAttachedItem)
var dialist = populateList(formJson, dbpItems, dbpAttachedItem)
var hrlist = populateList(formJson, hrItem, hrAttachedItem);

var sys = calculateAverage(syslist, sigfig);
var dia = calculateAverage(dialist, sigfig);
var hr = calculateAverage(hrlist, sigfig);

logger("Sys: " + sys);
logger("Dia: " + dia);
logger("HR: " + hr);

if (pending) return attachedItemCodeList[0];
if (
    sys > sys_max_range ||
    sys < sys_min_range ||
    dia > dia_max_range ||
    dia < dia_min_range ||
    hr > hr_max_range ||
    hr < hr_min_range
) return attachedItemCodeList[2]; // Out of Normal Range
else if ( // IR
    sys <= sys_max_range &&
    sys >= sys_min_range &&
    dia <= dia_max_range &&
    dia >= dia_min_range &&
    hr <= hr_max_range &&
    hr >= hr_min_range
) return attachedItemCodeList[1]; // Within Normal Range

return attachedItemCodeList[0];

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
            if (item && targetItem.indexOf(item.name) !== -1) {
                count++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                }
            }
        }
    }
    if (list.length < count) pending = true;
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