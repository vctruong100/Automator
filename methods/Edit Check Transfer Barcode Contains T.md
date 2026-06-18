// Version: v1
// Purpose: Validates transfer barcodes contain required T prefix.

const item = itemJson.item;

try {
    const itemGroupName = getItemGroupName(formJson);

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

// Retrieves the name of the item group containing the current item.
function getItemGroupName(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;
    
        for (var j = 0; j < items.length; j++) {
            var currentItem = items[j];
            if (currentItem.id === item.id) {
                return group.name;
            }
        }
    }
    return null;
}

// Checks if the input string contains the specified keyword (case-insensitive). Returns false for null/undefined input.
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

<<<<<<< Updated upstream
    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}
=======
    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}
>>>>>>> Stashed changes
