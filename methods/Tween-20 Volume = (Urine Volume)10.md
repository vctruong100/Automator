// Version: v1
// Purpose: Calculates Tween-20 volume as urine volume divided by 10.

const urineJugItem = [
    "Weight of Void Jug+Urine"
];
const emptyJugItem = [
    "Weight of Empty Void Jug"
];
const item = itemJson.item;
var groupid = null;

try {
    getItemGroupID(formJson.form);

    const urineJug = getItemValueFromSameGroup(formJson.form, urineJugItem);
    const voidJug = getItemValueFromSameGroup(formJson.form, emptyJugItem);

    const sigfig = item.significantDigits;

    var weight = (urineJug - voidJug).toFixed(0);
    if (!weight) return null;

    var volume = weight / 1.02;
    if (!volume) return null;

    var tween = volume / 10;
    if (tween) return tween.toFixed(sigfig);

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Returns the ID of the item group containing the current item within the given form.
function getItemGroupID(form) {
    for (var i = 0; i < form.itemGroups.length; i++) {
    var group = form.itemGroups[i];
    var items = group.items;
    if (!items || items.length < 1) continue;

    for (var j = 0; j < items.length; j++) {
        var currentItem = items[j];
        if (currentItem.id === item.id) {
            groupid = group.id;
            break;
        }
    }
    if (groupid) break;
    }
}

// Retrieves data for: getItemValueFromSameGroup.
function getItemValueFromSameGroup(form, itemName) {
    var value = null;
    for (var i = 0; i < form.itemGroups.length; i++) {
        var group = form.itemGroups[i];
        if (group.id !== groupid) continue;

        var items = group.items;
        for (var j = 0; j < items.length; j++) {
            var item = items[j];
            if (item && item.name === itemName) {
                value = item.value;
                if (value && value !== null) return value;
            }
        }
    }
    return null; 
}