const urineJugItem = [
    "Weight of Void Jug+Urine"
];
const emptyJugItem = [
    "Weight of Empty Void Jug"
];

const sigfig = itemJson.item.significantDigits;

var urineJug = pullItemFromForm(formJson, urineJugItem);
var voidJug = pullItemFromForm(formJson, emptyJugItem);

var total = (urineJug - voidJug);

if (total) return total.toFixed(sigfig);

if (total == 0) return (0).toFixed(sigfig);

return null;

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    var total = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item && item.value !== null) {
                logger(item.value);
                total += parseInt(item.value);
            }
        }
    }
    logger("Total: " + total);
    return total;
}