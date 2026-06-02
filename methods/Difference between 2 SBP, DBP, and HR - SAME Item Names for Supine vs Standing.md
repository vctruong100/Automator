const sysItem = [
    "Systolic BP (P: 90 - 140 mmHg)",
    "SYS (I: 80-150)",
    "Systolic Blood Pressure",
]
const diaItem = [
    "Diastolic BP (E: 50 - 95 mmHg)",
    "DIA (I: 50-99)",
    "DIA (I: 50-90)"
]

const sysAttachedItem = [
    "Systolic_BP_DIFF_@3min",
    "Systolic_BP_DIFF_@3min_repeat",
    "Systolic_BP_Diff",
    "ortho sbp check",
]

const diaAttachedItem = [
    "Diastolic_BP_DIFF_@3min",
    "Diastolic_BP_Diff",
    "Repeat_ortho dbp check",
]

const item = itemJson.item;
var isRepeat = false;
var semiValue, standing;

try {
    var groupName = getItemGroupName(formJson);
    if (containsValue(groupName, "repeat")) isRepeat = true;
    logger("Is repeat: " + isRepeat)
    logger("Attached item: " + item.name);
    if (sysAttachedItem.indexOf(item.name) !== -1) {
        semi = pullItemFromForm(formJson, sysItem, isRepeat, false);
        standing = pullItemFromForm(formJson, sysItem, isRepeat, true);
    } else if (diaAttachedItem.indexOf(item.name) !== -1) {
        semi = pullItemFromForm(formJson, diaItem, isRepeat, false);
        standing = pullItemFromForm(formJson, diaItem, isRepeat, true);
    }
    logger("Semi: " + semi);
    logger("Standing: " + standing);
    if (semi && standing) return String(semi - standing);
    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
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

function pullItemFromForm(form, targetItem, repeat, isStanding) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    if (repeat) { 
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") {
                    if (isStanding && containsValue(group.name, "standing")) return item.value;
                    else if (!isStanding && containsValue(group.name, "supine")) return item.value;
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") {
                    if (isStanding && containsValue(group.name, "standing")) return item.value;
                    else if (!isStanding && containsValue(group.name, "supine")) return item.value;
                }
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