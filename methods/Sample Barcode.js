/* jshint strict: false */

// Version: v1
// Purpose: Transfer barcode validation and formatting.
var item = itemJson.item;

var itemNames = [
    "BE_Date of sample 0 EOI",
    "BE_Date of sample 0",
    "BE_Date of sample 0 (PK)",
    "BE_Date of sample 1 (ADA, nAb)",
    "BE_Date of sample 2",
    "BE_Date of sample 3",
    "BE_Date of sample 4 (Serum Biomarker)",
    "BE_Date of sample 5 (Plasma biomarker)",
    "BE_Date of sample 6 (WB DNA)",

]
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
function pullItemFromForm(form, targetItem, groupName) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name)) return item;
        }
    }
    return null;
}

function parseItemContext(item) {
    var context = JSON.parse(getRelatedItemDataContext(item))
    var collectedBarcodes = context.collectedBarcodes;
    if (collectedBarcodes && collectedBarcodes != null) return collectedBarcodes;
    return null;
}

try {
    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
    logger("Group name: " + parsedGroupName);
    
    var item = pullItemFromForm(formJson, itemNames, parsedGroupName)
    return parseItemContext(item.name);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}