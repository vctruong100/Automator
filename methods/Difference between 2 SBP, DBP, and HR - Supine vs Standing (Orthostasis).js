/* jshint strict: false */

// Version: v1
// Purpose: Computes differences between supine and standing vitals (same item names for both repeat and non-repeat item groups).

// Add item names
var sysItem = ["SYS (P: 91 - 150)",]
var diaItem = ["DIA (P:  51 - 100)",]
var hrItem = ["HR (P: 61 - 100)", "HR (I: 45-100)"]

var semi = null;
var standing = null;

var item = itemJson.item;

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

function getOrthostasis(form, targetItems, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var i, j, group, item;

    var semi = null;
    var standing = null;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];

        if (!group || group.canceled) continue;

        if (isRepeat) {
            if (!containsValue(group.name, "repeat")) continue;
        } else {
            if (containsValue(group.name, "repeat")) continue;
        }

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];

            if (
                targetItems.indexOf(item.name) !== -1 &&
                item.value !== null &&
                item.value !== "" &&
                !item.canceled
            ) {
                if (containsValue(group.name, "semi") && semi === null) {
                    semi = parseFloat(item.value);
                    logger("Captured Semi from group: " + group.name + " value: " + semi);
                }

                if (containsValue(group.name, "stand") && standing === null) {
                    standing = parseFloat(item.value);
                    logger("Captured Standing from group: " + group.name + " value: " + standing);
                }
            }
        }

        if (semi !== null && standing !== null) {
            logger("Final Semi: " + semi);
            logger("Final Standing: " + standing);

            return semi - standing;
        }
    }

    return null;
}

try {
    var isRepeat = false;
    var difference = null;
    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
    
    logger("Group name: " + parsedGroupName);
    if (containsValue(parsedGroupName, "repeat")) isRepeat = true;
    logger("Attached item: " + item.name);

    if (containsValue(item.name, 'sys')) {
        difference = getOrthostasis(formJson, sysItem, isRepeat);
    } else if (containsValue(item.name, 'dia')) {
        difference = getOrthostasis(formJson, diaItem, isRepeat);
    } else if (containsValue(item.name, 'hr')) {
        difference = getOrthostasis(formJson, hrItem, isRepeat);
    }

    return difference.toFixed(0);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}