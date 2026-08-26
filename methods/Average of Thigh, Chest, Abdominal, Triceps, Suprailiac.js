/* jshint strict: false */

// Version: v1
// Description: Calculates average height from explicitly configured height measurement items on the current form, excluding blank or canceled entries and summary fields.

var form = formJson.form;
var attachedItem = itemJson.item;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function containsStandaloneKeyword(input, keyword) {
    var value = normalizeName(input);
    var target = normalizeName(keyword);
    var startIndex = 0;
    var index;
    var before;
    var after;

    while (startIndex < value.length) {
        index = value.indexOf(target, startIndex);
        if (index === -1) return false;

        before = index === 0 ? "" : value.charAt(index - 1);
        after = index + target.length >= value.length ? "" : value.charAt(index + target.length);

        if ((before === "" || !/[A-Z0-9]/.test(before)) && (after === "" || !/[A-Z0-9]/.test(after))) return true;

        startIndex = index + target.length;
    }

    return false;
}

function matchesMetric(itemName, metric) {
    var name = normalizeName(itemName);

    if (metric === "CHEST") return containsStandaloneKeyword(name, "CHEST");
    if (metric === "THIGH") return containsStandaloneKeyword(name, "THIGH");
    if (metric === "ABDOMINAL") return containsStandaloneKeyword(name, "ABDOMINAL");
    if (metric === "TRICEPS") return containsStandaloneKeyword(name, "TRICEPS");
    if (metric === "SUPRAILIAC") return containsStandaloneKeyword(name, "SUPRAILIAC");
    return false;
}

function getMetricFromAverageItem(itemName) {
    var name = normalizeName(itemName);

    if (containsValue(name, "CHEST")) return "CHEST";
    if (containsValue(name, "THIGH")) return "THIGH";
    if (containsValue(name, "ABDOMINAL")) return "ABDOMINAL";
    if (containsValue(name, "TRICEPS")) return "TRICEPS";
    if (containsValue(name, "SUPRAILIAC")) return "SUPRAILIAC";

    return null;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;

    return false;
}

function addNumericValue(list, value) {
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) list.push(numericValue);
}

function populateList(formJsonValue, metric, attachedItemName, isFemale) {
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return list;

    if (!isFemale) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;
            if (containsValue(group.name, "female")) continue; 
            items = group.items;
            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItemName) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    } else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;
            if (containsStandaloneKeyword(group.name, "male")) continue;
            items = group.items;

            for (j = items.length - 1; j >= 0; j--) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItemName && list.length > 1) return list;

                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                    addNumericValue(list, groupItem.value);
                }
            }
        }
    }

    return list;
}

function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    var avg = sum / values.length;
    var factor = Math.pow(10, sigfig);
    return (Math.round(avg * factor) / factor).toFixed(sigfig);
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isFemale = parsedGroupName ? containsValue(parsedGroupName, "female") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is female: " + isFemale);

    var metric = getMetricFromAverageItem(attachedItem.name);
    logger("Attached Item name: " + attachedItem.name);
    logger("Metric: " + metric);

    if (isFemale) {
        var tricepList = populateList(formJson, "TRICEPS", attachedItem.name, isFemale);
        var thighList = populateList(formJson, "THIGH", attachedItem.name, isFemale);
        var suprailiacList = populateList(formJson, "SUPRAILIAC", attachedItem.name, isFemale);
        
        logger("TRICEPS List: " + tricepList);
        logger("THIGH List: " + thighList);
        logger("SUPRAILIAC List: " + suprailiacList);
    
        logger("TRICEPS List length: " + tricepList.length);
        logger("THIGH List length: " + thighList.length);
        logger("SUPRAILIAC List length: " + suprailiacList.length);
    
        var tricepAvg = calculateAverage(tricepList, attachedItem.significantDigits);
        var thighAvg = calculateAverage(thighList, attachedItem.significantDigits);
        var suprailAvg = calculateAverage(suprailiacList, attachedItem.significantDigits);
        
        logger("TRICEPS Average: " + tricepAvg);
        logger("THIGH Average: " + thighAvg);
        logger("SUPRAILIAC Average: " + suprailAvg);
    
        if (metric === "TRICEPS") return tricepAvg;
        if (metric === "THIGH") return thighAvg;
        if (metric === "SUPRAILIAC") return suprailAvg;
    } else {
        var chestList = populateList(formJson, "CHEST", attachedItem.name, isFemale);
        var thighList = populateList(formJson, "THIGH", attachedItem.name, isFemale);
        var abdominaltList = populateList(formJson, "ABDOMINAL", attachedItem.name, isFemale);
        
        logger("CHEST List: " + chestList);
        logger("THIGH List: " + thighList);
        logger("ABDOMINAL List: " + abdominaltList);
    
        logger("CHEST List length: " + chestList.length);
        logger("THIGH List length: " + thighList.length);
        logger("ABDOMINAL List length: " + abdominaltList.length);
    
        var chestAvg = calculateAverage(chestList, attachedItem.significantDigits);
        var thighAvg = calculateAverage(thighList, attachedItem.significantDigits);
        var abAvg = calculateAverage(abdominaltList, attachedItem.significantDigits);
        
        logger("CHEST Average: " + chestAvg);
        logger("THIGH Average: " + thighAvg);
        logger("ABDOMINAL Average: " + abAvg);
    
        if (metric === "CHEST") return chestAvg;
        if (metric === "THIGH") return thighAvg;
        if (metric === "ABDOMINAL") return abAvg;
    }

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}