/* jshint strict: false */

// Version: v2
// Purpose: General vital signs range protocol check. Does not require any item names.

// Inclusive
var sys_min_range = 91;
var sys_max_range = 150;

var dia_min_range = 51;
var dia_max_range = 100;

var hr_min_range = 61;
var hr_max_range = 100;

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

    if (metric === "SYS") return containsStandaloneKeyword(name, "SYS") || containsStandaloneKeyword(name, "SBP") || containsValue(name, "SYSTOLIC");
    if (metric === "DIA") return containsStandaloneKeyword(name, "DIA") || containsStandaloneKeyword(name, "DBP") || containsValue(name, "DIASTOLIC");
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "RATE") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE");

    return false;
}

function getMetricFromAverageItem(itemName) {
    var name = normalizeName(itemName);

    if (containsValue(name, "SYS") || containsValue(name, "SBP") || containsValue(name, "SYSTOLIC")) return "SYS";
    if (containsValue(name, "DIA") || containsValue(name, "DBP") || containsValue(name, "DIASTOLIC")) return "DIA";
    if (containsValue(name, "HR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE") || containsValue(name, "RATE")) return "HR";

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

function populateList(formJsonValue, metric, attachedItemName, isRepeat) {
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return list;

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = items.length - 1; j >= 0; j--) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItemName && list.length > 1) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    if (list.length >= 3) return list;
                }
            }
        }
    } else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            items = group.items;

            for (j = 0; j < items.length; j++) {
                groupItem = items[j];
                if (!groupItem) continue;
                if (groupItem.name == attachedItemName) return list;
                if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                    addNumericValue(list, groupItem.value);
                    if (list.length >= 3) return list;
                }
            }
        }
    }

    return list;
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

function pullItemFromForm(formJsonValue, metric, groupName, isRepeat) {
    logger("Target metric: " + metric);
    var itemGroups = formJsonValue.form.itemGroups;
    var list = [];
    var average = null;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;

        for (var j = group.items.length - 1; j >= 0; j--) {
            var item = group.items[j];
            if (!item) continue;

            logger("Item name: " + item.name);
            if (getMetricFromAverageItem(item.name) === metric) {
                list = populateList(formJsonValue, metric, attachedItem.name, isRepeat);
            }
            if (list.length > 0) {
                logger("List: " + list);
                average = calculateAverage(list);
                logger("Average: " + average);
            }

            if (average !== null) return average;
            if (matchesMetric(item.name, metric) && item.value !== null && item.value !== "") return item.value;
        }
    }

    return null;
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    logger("Group name: " + parsedGroupName);
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    var sys = pullItemFromForm(formJson, "SYS", parsedGroupName, isRepeat);
    var dia = pullItemFromForm(formJson, "DIA", parsedGroupName, isRepeat);
    var hr = pullItemFromForm(formJson, "HR", parsedGroupName, isRepeat);

    logger("Systolic: " + sys);
    logger("Diastolic: " + dia);
    logger("Heart rate: " + hr);

    if (!sys || sys === null || !dia || dia == null || !hr || hr == null) return attachedItem.codeListItems[4].codedValue;

    // OOR
    if (
        sys > sys_max_range ||
        sys < sys_min_range ||
        dia > dia_max_range ||
        dia < dia_min_range ||
        hr > hr_max_range ||
        hr < hr_min_range
    ) return attachedItem.codeListItems[1].codedValue; // Out of Protocol Range
    else if ( // IR
        sys <= sys_max_range &&
        sys >= sys_min_range &&
        dia <= dia_max_range &&
        dia >= dia_min_range &&
        hr <= hr_max_range &&
        hr >= hr_min_range
    ) return attachedItem.codeListItems[0].codedValue; // Within Normal Range

    return attachedItem.codeListItems[4].codedValue;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
