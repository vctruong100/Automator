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
            if (item && urineItem.indexOf(item.name) !== -1 && item.value !== null) urineJug = item.value;
            if (item && emptyItem.indexOf(item.name) !== -1 && item.value !== null) voidJug = item.value;
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
