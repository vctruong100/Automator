// Add item names (for different naming conventions)
const sysItems = [
    "SYS (2 of 3)",
    "Repeat SYS (2 of 3)",
    "Repeat SYS",
];

const diaItems = [
    "DIA (2 of 3)",
    "Repeat DIA (2 of 3)",
    "Repeat DIA",
];

const hrItems = [
    "HR (2 of 3)",
    "Repeat HR (2 of 3)",
    "HR (60 - 100 bpm)",
    "Repeat HR",
]

const tempItems = [
    "Repeat Temp:",
]

// Inclusive (Edit)
var sys_min_range = 90;
var sys_max_range = 140;

var dia_min_range = 50
var dia_max_range = 90;

var hr_min_range = 60;
var hr_max_range = 100;

// ======== Don't modify ========
var item = itemJson.item;

var isRepeat = isItemInRepeat(formJson);

if (isRepeat) {
    var sys = pullItemFromLast(formJson, sysItems);
    var dia = pullItemFromLast(formJson, diaItems);
    var hr = pullItemFromLast(formJson, hrItems);
}
else {
    var sys = pullItemFromFirst(formJson, sysItems);
    var dia = pullItemFromFirst(formJson, diaItems);
    var hr = pullItemFromFirst(formJson, hrItems);
}

log();

if (!sys || sys === null || !dia || dia == null || !hr || hr == null) return itemJson.item.codeListItems[4];
// OOR
if (
    sys > sys_max_range ||
    sys < sys_min_range ||
    dia > dia_max_range ||
    dia < dia_min_range ||
    hr > hr_max_range ||
    hr < hr_min_range
) return itemJson.item.codeListItems[1]; // Out of Protocol Range
else if ( // IR
    sys <= sys_max_range &&
    sys >= sys_min_range && 
    dia <= dia_max_range &&
    dia >= dia_min_range &&
    hr <= hr_max_range &&
    hr >= hr_min_range
) return itemJson.item.codeListItems[0]; // Within Normal Range

return itemJson.item.codeListItems[4];

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

function pullItemFromLast(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
    return null;
}

function pullItemFromFirst(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
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

function isItemInRepeat(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                if (containsRepeat(group.name)) return true;
            }
        }
    }
    return false;
}

function containsRepeat(input) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf("repeat") !== -1;
}