
// Add item names
const heightItemList = [
    "BMI_HEIGHT #1", 
    "BMI_HEIGHT #2"
];

const weightItemList = [
    "BMI_WEIGHT #1", 
    "BMI_WEIGHT #2"
];

const heightAttachedItem = [
    "Average Height"
];

const weightAttachedItem = [
    "Average Weight"
];

// ======== Don't modify ========
const item = itemJson.item;
const sigfig = item.significantDigits;
var maxCount = 0;
var heightList = [];
var weightList = [];
var heightAvg = 0;
var weightAvg = 0;

heightList = populateList(formJson, heightItemList);
weightList = populateList(formJson, weightItemList);

heightAvg = calculateAverage(heightList, sigfig);
weightAvg = calculateAverage(weightList, sigfig);
logger("Height List: " + heightList);
logger("Weight List: " + weightList);

logger("Height List length: " + heightList.length);
logger("Weight List length: " + weightList.length)

logger("Max count: " + maxCount);
logger("Height Average: " + heightAvg);
logger("Weight Average: " + weightAvg);

if (heightAttachedItem.indexOf(item.name) !== -1) {
    return heightAvg;
}
if (weightAttachedItem.indexOf(item.name) !== -1) {
    return weightAvg;
}
return null;

function populateList(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
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
    return (Math.round(avg * factor) / factor).toFixed(sigfig);
}