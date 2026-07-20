/* jshint strict: false */

// Version: v3
// Purpose: Calculates the average SYS, DIA, and HR, then checks protocol ranges.
//          Combines Protocol Check Vital Signs Range v2 and Average DIA, SYS, Heart Rate v2.

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
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "PR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE");

    return false;
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

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

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

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Is repeat: " + isRepeat);

    var sysList = populateList(formJson, "SYS", attachedItem.name, isRepeat);
    var diaList = populateList(formJson, "DIA", attachedItem.name, isRepeat);
    var hrList = populateList(formJson, "HR", attachedItem.name, isRepeat);

    logger("sysList: " + sysList);
    logger("diaList: " + diaList);
    logger("hrList: " + hrList);

    var sys = calculateAverage(sysList);
    var dia = calculateAverage(diaList);
    var hr = calculateAverage(hrList);

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
