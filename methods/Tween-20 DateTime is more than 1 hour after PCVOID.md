const pcvoidItem = [
    "PCVOID"
]
const enteredTimeItem = [
    "Tween-20 Date/Time"
]
const attachedCodeList = [
    "YES",
    "NO",
]
const difference = 60;

const item = itemJson.item;
var groupid = null;

getItemGroupID(formJson.form);
var pcvoid = getItemValueFromSameGroup(formJson, pcvoidItem);

const enteredTime = getItemValueFromSameGroup(formJson, enteredTimeItem);
if (!pcvoid || pcvoid == null || !enteredTime || enteredTime == null) return null;

var pcvoidMs = pcvoid.dateValueMs;
var collectedTimeMs = enteredTime.dateValueMs;

var differenceMs = collectedTimeMs - pcvoidMs;
logger(differenceMs)
if (differenceMs < 0) {
    return attachedCodeList[0];
}
var differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger(differenceInMins)
if(differenceInMins >= difference){
    return attachedCodeList[0];
}

return attachedCodeList[1];

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