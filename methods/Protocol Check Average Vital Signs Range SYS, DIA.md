/* jshint strict: false */

// Version: v2
// Purpose: Protocol range check for averaged SBP and DBP.

// Add item names
var dbpItems = [
    "Diastolic BP",
    "SCRN_Diastolic BP",
    "DIA",
    "Repeat DIA (2 of 3)",
    "Repeat DIA (3 of 3)",
    "DIA (1 of 3)",
    "DIA (2 of 3)",
    "DIA (3 of 3)"
];

var dbpAttachedItem = [
    "Average DBP",
    "🧮 DIA AVERAGE:"
];

var sysItem = [
    "Systolic BP",
    "SCRN_Systolic BP",
    "ET Threshold SBP",
    "SYS",
    "SYS (1 of 3)",
    "SYS (2 of 3)",
    "SYS (3 of 3)"
];

var sysAttachedItem = [
    "Average SBP",
    "🧮 SYS_AVERAGE:"
];

var repeatItem = [
    "▶ Is VS repeat required?"
];

// Inclusive
var sysAvg_maxRange = 170;
var diaAvg_maxRange = 100;

// ======== Don't modify ========
var item = itemJson.item;
var sigfig = item.significantDigits;

try {
    var repeat = pullItemFromForm(formJson, repeatItem);
    var isRepeat = containsValue(repeat, "yes");

    logger("Repeat: " + repeat);

    var sysList = populateList(formJson, sysItem, sysAttachedItem, isRepeat);
    var diaList = populateList(formJson, dbpItems, dbpAttachedItem, isRepeat);

    var sysAvg = calculateAverage(sysList, sigfig);
    var diaAvg = calculateAverage(diaList, sigfig);

    logger("Sys Avg: " + sysAvg);
    logger("Dia Avg: " + diaAvg);

    if (
        sysAvg !== null &&
        diaAvg !== null &&
        (sysAvg >= sysAvg_maxRange || diaAvg >= diaAvg_maxRange)
    ) {
        return item.codeListItems[1].codedValue; // Yes
    }

    return item.codeListItems[2].codedValue; // No
}
catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function pullItemFromForm(form, targetItems) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) {
        return null;
    }

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];

        if (!group || group.canceled) {
            continue;
        }

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];

            if (
                targetItems.indexOf(item.name) !== -1 &&
                item.value !== null &&
                item.value !== "" &&
                !item.canceled
            ) {
                return item.value;
            }
        }
    }

    return null;
}

function calculateAverage(values, sigfig) {
    if (!values || values.length === 0) {
        return null;
    }

    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (!isNaN(values[i])) {
            sum += values[i];
            count++;
        }
    }

    if (count === 0) {
        return null;
    }

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);

    return Math.round(avg * factor) / factor;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    return input.toString().toLowerCase().indexOf(keyword) !== -1;
}

function populateList(form, targetItems, attachedItems, repeat) {
    var itemGroups = form.form.itemGroups;
    var list = [];
    var i, j, group, item;

    if (!itemGroups || itemGroups.length < 1) {
        return [];
    }

    if (repeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];

            if (!group || group.canceled) {
                continue;
            }

            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];

                if (
                    item &&
                    targetItems.indexOf(item.name) !== -1 &&
                    item.value !== null &&
                    item.value !== "" &&
                    !isNaN(item.value)
                ) {
                    list.push(parseFloat(item.value));

                    if (list.length >= 3) {
                        return list;
                    }
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];

            if (!group || group.canceled) {
                continue;
            }

            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];

                if (
                    item &&
                    targetItems.indexOf(item.name) !== -1 &&
                    item.value !== null &&
                    item.value !== "" &&
                    !isNaN(item.value)
                ) {
                    list.push(parseFloat(item.value));

                    if (list.length >= 3) {
                        return list;
                    }
                }
            }
        }
    }

    return list;
}
