/* jshint strict: false */

// Version: v1
// Purpose: Evaluates and flags dipstick results across all relevant items.

var itemid = itemJson.item.id;

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
function pullItemFromForm(form) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    var firstItem = null;
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (j == 0) firstItem = item;
            if (item.id === itemid) return firstItem;
        }
    }
    return null;
}

try {
    var item = pullItemFromForm(formJson);
    logger("Item name: " + item.name + ", value: " + item.value);

    var normalizedItem = item.name.trim().toLowerCase();
    if (normalizedItem.indexOf("ph") !== -1) {
        if (item.value == item.codeListItems[6].codedValue) return "YES";
        else return "NO";
    }
    if (normalizedItem.indexOf("gravity") !== -1) {
        if (item.value == item.codeListItems[0].codedValue) return "YES";
        else return "NO";
    }
    if (normalizedItem.indexOf("urobilinogen") !== -1) {
        if (item.value == item.codeListItems[0].codedValue || item.value == item.codeListItems[1].codedValue) return "NO";
    }
    if (item.value == item.codeListItems[0].codedValue) return "NO";

    if (!item || item.value == null) return "NO";

    return "YES";
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
