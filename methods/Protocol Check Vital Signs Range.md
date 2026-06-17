/* jshint strict: false */

// Version: v1
// Purpose: General vital signs range protocol check.

var sysItems = [
    "SYS (60 - 200) mmHg", "SYS (60 - 200)",
];

var diaItems = [
    "DIA (40 - 110) mmHg", "DIA (40 - 110)",
];

var hrItems = [
    "HR (50 - 100) bpm", "HR (30 - 200)",
]

// Inclusive
var sys_min_range = 60;
var sys_max_range = 200;

var dia_min_range = 40
var dia_max_range = 110;

var hr_min_range = 30;
var hr_max_range = 200;

// ======== Don't modify ========
var item = itemJson.item;

try {
    var isRepeat = isItemInRepeat(formJson);

    var sys = pullItemFromForm(formJson, sysItems, isRepeat);
    var dia = pullItemFromForm(formJson, diaItems, isRepeat);
    var hr = pullItemFromForm(formJson, hrItems, isRepeat);

    log();

    if (!sys || sys === null || !dia || dia == null || !hr || hr == null) return itemJson.item.codeListItems[4].codedValue;
    // OOR
    if (
        sys > sys_max_range ||
        sys < sys_min_range ||
        dia > dia_max_range ||
        dia < dia_min_range ||
        hr > hr_max_range ||
        hr < hr_min_range
    ) return itemJson.item.codeListItems[1].codedValue; // Out of Protocol Range
    else if ( // IR
        sys <= sys_max_range &&
        sys >= sys_min_range &&
        dia <= dia_max_range &&
        dia >= dia_min_range &&
        hr <= hr_max_range &&
        hr >= hr_min_range
    ) return itemJson.item.codeListItems[0].codedValue; // Within Normal Range

    return itemJson.item.codeListItems[4].codedValue;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function log() {
    logger("sys: " + sys);
    logger("dia:  " + dia);
    logger("hr: " + hr);
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            keepers.push(formData);
        } else {

        }
    }
    return keepers;
}

function pullItemFromForm(form, targetItem, repeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    if (repeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            if (group.id !== groupid) continue;
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
            }
        }
    }

    return null;
}


function calculateAverage(values) {
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
    return avg;
}
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function isItemInRepeat(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                if (containsValue(group.name, "repeat")) return true;
            }
        }
    }
    return false;
}
