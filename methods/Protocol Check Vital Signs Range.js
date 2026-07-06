/* jshint strict: false */

// Version: v1
// Purpose: General vital signs range protocol check.

const sysItems = ["SYS MEAN (P: 91 - 150)", "SYS (P: 91 - 150)", "SYS (60 - 200) mmHg", "SYS  (I: 90 - 150)", "SYS (60 - 200)"];
const diaItems = ["DIA MEAN (P: 51 - 100)", "DIA (P: 51 - 100)", "DIA (P:  51 - 100)", "DIA (40 - 110) mmHg", "DIA (I: 50 - 100)", "DIA (40 - 110)",];
const hrItems = ["HR MEAN (P: 61 - 100)", "HR (P: 61 - 100)", "HR (50 - 100) bpm", "HR (30 - 100)"];

const sysAvg = ["HR MEAN (P: 61 - 100)", "SYS MEAN AVERAGE", "SYS MEAN AVERAGE REPEAT",];
const diaAvg = ["DIA MEAN (P: 51 - 100)", "DIA MEAN AVERAGE", "DIA MEAN AVERAGE REPEAT",];
const hrAvg = ["HR MEAN (P: 61 - 100)", "HR MEAN AVERAGE", "HR MEAN AVERAGE REPEAT"];

// Inclusive
var sys_min_range = 91;
var sys_max_range = 150;

var dia_min_range = 51;
var dia_max_range = 100;

var hr_min_range = 61;
var hr_max_range = 100;

var isRepeat = false;
var item = itemJson.item;

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

function pullItemFromForm(form, targetItem, itemAvg, groupName, isRepeat) {
    logger("Target Item: " + targetItem)
    logger("Item Avg: " + itemAvg)
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var average = null;
	if (!itemGroups || itemGroups.length < 1) return null;
	
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;

        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            
            logger("Item name: " + item.name);
            if (containsItemName(itemAvg, item.name)) list = populateList(formJson, targetItem, itemAvg, isRepeat);
            if (list.length > 0) {
                logger("List: " + list)
                average = calculateAverage(list);
                logger("Average: " + average);
            }
            
            if (average !== null) return average;
            if (containsItemName(targetItem, item.name) && item.value !== null) return item.value;
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

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

function populateList(form, targetItem, attachedItem, repeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    if (repeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];

                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
                    if (list.length >= 3) return list;
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name)) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
                    if (list.length >= 3) return list;
                }
            }
        }
    }
    return list;
}

try {
    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
    logger("Group name: " + parsedGroupName);
    if (containsValue(parsedGroupName, "repeat")) isRepeat = true;
    var sys = pullItemFromForm(formJson, sysItems, sysAvg, parsedGroupName, isRepeat);
    var dia = pullItemFromForm(formJson, diaItems, diaAvg, parsedGroupName, isRepeat);
    var hr = pullItemFromForm(formJson, hrItems, hrAvg, parsedGroupName, isRepeat);

    logger("Systolic: " + sys);
    logger("Diastolic: " + dia);
    logger("Heart rate: " + hr);
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