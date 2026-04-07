// Add item names
const itemList = [
    "Hip Circumference #1", 
    "Hip Circumference #2"
];

// ======== Don't modify ========
const sigfig = itemJson.item.significantDigits;
var maxCount = 2; 
var list = [];
var avg = 0;

list = populateList(formJson, itemList);

avg = calculateAverage(list, sigfig);
logger("List: " + list);
logger("List length: " + list.length)
logger("Max count: " + maxCount);
logger("Average: " + avg)
if (list.length === maxCount) return (avg);
return null;

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