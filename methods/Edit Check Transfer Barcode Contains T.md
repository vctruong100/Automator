const item = itemJson.item;
const itemGroupName = getItemGroupName(formJson);

if (containsValue(itemGroupName, "self-collection") || containsValue(itemGroupName, "pg dna")) {
    if (containsValue(item.value, "t")) {
        customErrorMessage("Barcode cannot contain 'T'")
        return false;
    }
}

return true;

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