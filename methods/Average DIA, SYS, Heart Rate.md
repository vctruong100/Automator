/* jshint strict: false */

// Add item names
// items to pull for calculation
const sysItem = ["SYS (60 - 200) mmHg", "SYS (I: 90 - 150)"];
const diaItem = ["DIA (40 - 110) mmHg", "DIA (I: 50 - 100)",];
const hrItem = ["HR (50 - 100) bpm"];

// items to attach
const sysAttach = ["SYS MEAN AVERAGE"];
const diaAttach = ["DIA MEAN AVERAGE"];
const hrAttach = ["HR MEAN AVERAGE"];

// ======== Don't modify ========
var form = formJson.form;
var item = itemJson.item;

try {
    var isRepeat = false;
    var groupName = getItemGroupName(formJson);
    logger("Group Name: " + groupName)
    if (groupName) {
        isRepeat = containsValue(groupName, "repeat")
    }

    var sysList = populateList(formJson, sysItem, sysAttach, isRepeat);
    var diaList = populateList(formJson, diaItem, diaAttach, isRepeat);
    var hrList = populateList(formJson, hrItem, hrAttach, isRepeat);

    logger("Is it a repeat? " + isRepeat);
    logger("sysList: " + sysList);
    logger("diaList: " + diaList);
    logger("hrList: " + hrList);
    
    var avgSys = calculateAverage(sysList);
    var avgDia = calculateAverage(diaList);
    var avgHR = calculateAverage(hrList);
    
    logger("Average PR: " + avgSys);
    logger("AVerage QRS: " + avgDia);
    logger("Average QT: " + avgHR);
    
    if (sysAttach.indexOf(item.name) !== -1) return avgSys;
    if (diaAttach.indexOf(item.name) !== -1) return avgDia;
    if (hrAttach.indexOf(item.name) !== -1) return avgHR;
    
    return null;
} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

function getItemGroupName(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;
    
        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                return group.name;
            }
        }
    }
    return null;
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
                if (attachedItem.indexOf(item.name) !== -1 && list.length > 1) return list;
                if (item && targetItem.indexOf(item.name) !== -1) {
                    logger(item.value);
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
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
                if (attachedItem.indexOf(item.name) !== -1) return list;
                if (item && targetItem.indexOf(item.name) !== -1) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
                }
            }
        }
    }
    return list;
}