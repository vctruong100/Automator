const itemList = [
    "Waist Circumference #1", 
    "Waist Circumference #2"
];
const attachedItem = [
    "Average Waist Circumference"
];

const sigfig = itemJson.item.significantDigits;
var maxCount = 0; 
var list = [];
var avg = 0;

list = populateList(formJson, itemList, list);

avg = calculateAverage(list, sigfig);
logger("List: " + list);
logger("List length: " + list.length)
logger("Max count: " + maxCount);
logger("Average: " + avg)
if (list.length === maxCount) return (avg).toFixed(sigfig);
return null;

function populateList(form, targetItem, ilist) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
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