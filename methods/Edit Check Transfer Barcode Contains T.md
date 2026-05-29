const item = itemJson.item;
var triggerLetter = "T";
var groupName = getItemGroupName(formJson);
var isSerum = containsValue(groupName, "serum");
var isPlasma = containsValue(groupName, "plasma");

logger("Group Name: " + groupName);
if (isSerum || isPlasma) return true;

logger(item.value);

var value = (item.value || "").toString().trim().toUpperCase();
var letter = triggerLetter.toUpperCase();

if (!value) return false;

if (value.indexOf(triggerLetter) !== -1) {
    customErrorMessage("Value cannot contain " + triggerLetter);
    return false;
}

return true;

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
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