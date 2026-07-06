/* jshint strict: false */

// Version: v2
// Purpose: Calculates net urine weight (jug+urine minus empty jug).

var urineJugItem = [
    "Weight of Void Jug+Urine",
    "Total Weight of Urine and Jug",
];
var emptyJugItem = [
    "Weight of Empty Void Jug",
    "Weight of Empty Jug",
];

var sigfig = itemJson.item.significantDigits;

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
function GetNetTotal(form, urineItem, emptyItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    var total = 0;
	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        var urineJug = 0;
        var voidJug = 0;
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && containsItemName(urineItem, item.name) && item.value !== null) urineJug = item.value;
            if (item && containsItemName(emptyItem, item.name) && item.value !== null) voidJug = item.value;
        }
        total += urineJug - voidJug;
    }
    logger("Total: " + total);
    return total;
}

try {
    var netTotal = GetNetTotal(formJson, urineJugItem, emptyJugItem);

    if (netTotal) return netTotal.toFixed(sigfig);
    if (netTotal == 0) return (0).toFixed(sigfig);

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
