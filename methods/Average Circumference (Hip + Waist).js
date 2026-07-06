/* jshint strict: false */

// Version: v1
// Purpose: Computes sum of average hip and waist circumference values.

// Add item names
var hipItemList = [
    "Hip Circumference #1",
    "Hip Circumference #2"
];

var waistItemList = [
    "Waist Circumference #1",
    "Waist Circumference #2",
    "Waist Circumference #3",
    "Waist Circumference #4"
];

var hipAttachedItem = [
    "Average Hip Circumference"
]

var waistAttachedItem = [
    "Average Waist Circumference"
]

var item = itemJson.item;
var sigfig = item.significantDigits;
var maxCount = 2;
var waistList = [];
var hipList = [];
var waistAvg = 0;
var hipAvg = 0;


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
            if (item && containsItemName(targetItem, item.name)) {
                list.push(parseFloat(item.value));
                if (list.length >= maxCount) {
                    return list;
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
        if (values[i] == null || values[i] == "") return null;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);
    return (Math.round(avg * factor) / factor).toFixed(sigfig);
}

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

    if (containsItemName(hipAttachedItem, item.name) && hipList.length === maxCount) return hipAvg;
    if (containsItemName(waistAttachedItem, item.name) && waistList.length === maxCount) return waistAvg;

    return null;
} catch (e) {
    logger("Error in main method: " + e);
    return null;
}
