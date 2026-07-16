var itemName = ["EGTXT"]

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
function pullItemFromForm(form, targetItem, groupid) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || group.id !== groupid) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
}

function replaceWithDelimiter(value) {
    if (!value) return "";

    return value
        .toString()
        .replace(/\bWITH\b/gi, "|");
}

try {
    var itemDataContext = JSON.parse(getItemDataContext());
    var groupid = itemDataContext.itemGroupDataId;
    var ecgtext = pullItemFromForm(formJson, itemName, groupid);
    logger("ECG Text: " + ecgtext);
    var cleanedText = replaceWithDelimiter(ecgtext);
    logger("Cleaned ECG Text: " +  cleanedText);
    return cleanedText;
} catch (e) {
    logger("Error: " + e);
    return null;
}