const itemName = [
    "Glucometer Reading"    
]
const item = itemJson.item;
var result = getItemGroupID(formJson);
var groupid = result[0];
var itemGroupKey = result[1]
logger("groupid: " + groupid)
logger("Item Group Key: " + itemGroupKey);

var glycemia = parseInt(pullItemFromForm(formJson, itemName, groupid));

logger("Glycemia: " + glycemia)
if (itemGroupKey && parseInt(itemGroupKey) > 1) {
    try {
        if (glycemia <= 80) return "Repeat 15g of Glucose IV in Bolus.";
        else return "Provide a carbohydrate-rich meal if the study participant can safely eat (eg. bread meal). Monitor the glycemia every 30 to 60 minutes until study particiapnt has a glycemia of > 80mg/dl for at least 4 hours"
    } catch (e) {
        return null;
    }
}
else {
    try {
        if (glycemia <= 70) return "Proceed to IV Protocol.";
        else if (glycemia > 70 && glycemia <= 80) return "Recheck after 15 minutes, and repeat administration of +/- 15g of oral glucose if the study participant can safely swallow."
        else if (glycemia > 80) return "Provide a carbohydrate-rich meal (eg., bread meal) to prevent recurrence. If symptoms ongoing, continue glucose monitoring. If not, stop monitoring."
    } catch (e) {
        return null;
    }
}

return null;

function getItemGroupID(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;
    
        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                return [group.id, group.itemGroupRepeatKey];
            }
        }
    }
    return null;
}

function pullItemFromForm(form, targetItem, groupid) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        if (groupid !== group.id) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}