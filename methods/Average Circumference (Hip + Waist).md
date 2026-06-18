// Version: v1
// Purpose: Computes sum of average hip and waist circumference values.

// Add item names
const hipItemList = [
    "Hip Circumference #1", 
    "Hip Circumference #2"
];

const waistItemList = [
    "Waist Circumference #1", 
    "Waist Circumference #2",
    "Waist Circumference #3",
    "Waist Circumference #4"
];

const hipAttachedItem = [
    "Average Hip Circumference"
]

const waistAttachedItem = [
    "Average Waist Circumference"
]

// ======== Don't modify ========
const item = itemJson.item;
const sigfig = item.significantDigits;
var maxCount = 2; 
var waistList = [];
var hipList = [];
var waistAvg = 0;
var hipAvg = 0;

try {
    waistList = populateList(formJson, waistItemList);
    hipList = populateList(formJson, hipItemList);

    waistAvg = calculateAverage(waistList, sigfig);
    hipAvg = calculateAverage(hipList, sigfig);

    logger("Waist List: " + waistList);
    logger("Hip List: " + hipList);

    logger("Waist List length: " + waistList.length);
    logger("Hip List length: " + hipList.length);

    logger("Max count: " + maxCount);

    logger("Waist Average: " + waistAvg);
    logger("Hip Average: " + hipAvg);

    if (hipAttachedItem.indexOf(item.name) !== -1 && hipList.length === maxCount) return hipAvg;
    if (waistAttachedItem.indexOf(item.name) !== -1 && waistList.length === maxCount) return waistAvg;

    return null;
} catch (e) {
    logger("Error in main method: " + e);
    return null;
}

// Collects numeric values from form items matching target names, stopping when an attached item is encountered (first-to-last order).
function populateList(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var list = [];
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                list.push(parseFloat(item.value));
                if (list.length >= maxCount) {
                    return list;
                }
            }
        }
    }
    return list;
}

// Calculates the arithmetic mean of an array of numeric values, ignoring non-numeric entries.
function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (values[i] == null || values[i] == "") return null;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);
    return (Math.round(avg * factor) / factor).toFixed(sigfig);
}