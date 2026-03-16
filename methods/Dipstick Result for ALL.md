var itemid = itemJson.item.id;
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

return "YES";

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