/* jshint strict: false */

// Version: v1
// Purpose: Populate adverse event item in conmeds item group

var itemName = [
    "AE_Adverse event",
    "AE_TERM"
]

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

try {
    return pullItemFromForm(formJson, itemName);
} catch (e) {
    logger("Error: " + e);
    return null;
}

