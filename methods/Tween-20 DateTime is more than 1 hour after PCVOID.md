/* jshint strict: false */

// Version: v1
// Purpose: Validates Tween-20 datetime is >1 hour after PCVOID.

var pcvoidItem = [
    "PCVOID",
    "PCVOID w/edit check",
]
var enteredTimeItem = [
    "Tween-20 Date/Time"
]
var attachedCodeList = [
    "YES",
    "NO",
]
var difference = 60;

var item = itemJson.item;
var groupid = null;

try {
    getItemGroupID(formJson.form);
    var pcvoid = getItemValueFromSameGroup(formJson, pcvoidItem);

    var enteredTime = getItemValueFromSameGroup(formJson, enteredTimeItem);
    if (!pcvoid || pcvoid == null || !enteredTime || enteredTime == null) return null;

    var pcvoidMs = pcvoid.dateValueMs;
    var collectedTimeMs = enteredTime.dateValueMs;

    var differenceMs = collectedTimeMs - pcvoidMs;
    logger(differenceMs)
    if (differenceMs < 0) {
        return attachedCodeList[0]; // return YES
    }
    var differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

    logger(differenceInMins)
    if(differenceInMins >= difference){
        return attachedCodeList[0]; // return YES
    }

    return attachedCodeList[1]; // return NO
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function getItemGroupID(form) {
    for (var i = 0; i < form.itemGroups.length; i++) {
    var group = form.itemGroups[i];
    var items = group.items;
    if (!items || items.length < 1) continue;

    for (var j = 0; j < items.length; j++) {
        var it = items[j];
        if (it.id === item.id) {
            groupid = group.id;
            break;
        }
    }
    if (groupid) break;
    }
}

function getItemValueFromSameGroup(formJ, itemName) {
    var value = null;
    var form = formJ.form;
    for (var i = 0; i < form.itemGroups.length; i++) {
        var group = form.itemGroups[i];
        if (group.id !== groupid) continue;

        var items = group.items;
        for (var j = 0; j < items.length; j++) {
            var item = items[j];
            if (item && itemName.indexOf(item.name) !== -1) {
                if (item.value && item.value !== null) return item;
            }
        }
    }
    return null;
}
