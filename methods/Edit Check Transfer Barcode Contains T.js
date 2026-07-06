/* jshint strict: false */

// Version: v1
// Purpose: Validates transfer barcodes contain required T prefix.

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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

try {
    var itemGroupName = getItemGroupName(formJson);

    if (containsValue(itemGroupName, "self-collection") || containsValue(itemGroupName, "pg dna")) {
        if (containsValue(item.value, "t")) {
            customErrorMessage("Barcode cannot contain 'T'")
            return false;
        }
    }

    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
