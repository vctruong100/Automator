/* jshint strict: false */

// Version: v3
// Purpose: Computes orthostatic differences between supine/semi and standing vitals.
//          Uses keyword matching on item and group names instead of hardcoded item-name lists.

var attachedItem = itemJson.item;

var sysDifferenceRange = 20;
var diaDifferenceRange = 10;
var hrDifferenceRange = 30;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function containsStandaloneKeyword(input, keyword) {
    var value = normalizeName(input);
    var target = normalizeName(keyword);
    var startIndex = 0;
    var index;
    var before;
    var after;

    while (startIndex < value.length) {
        index = value.indexOf(target, startIndex);
        if (index === -1) return false;

        before = index === 0 ? "" : value.charAt(index - 1);
        after = index + target.length >= value.length ? "" : value.charAt(index + target.length);

        if ((before === "" || !/[A-Z0-9]/.test(before)) && (after === "" || !/[A-Z0-9]/.test(after))) return true;

        startIndex = index + target.length;
    }

    return false;
}

function matchesMetric(itemName, metric) {
    var name = normalizeName(itemName);

    if (metric === "SYS") return containsStandaloneKeyword(name, "SYS") || containsStandaloneKeyword(name, "SBP") || containsValue(name, "SYSTOLIC");
    if (metric === "DIA") return containsStandaloneKeyword(name, "DIA") || containsStandaloneKeyword(name, "DBP") || containsValue(name, "DIASTOLIC");
    if (metric === "HR") return containsStandaloneKeyword(name, "HR") || containsStandaloneKeyword(name, "PR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE");

    return false;
}

function getMetricFromAttachedItem(itemName) {
    if (matchesMetric(itemName, "SYS")) return "SYS";
    if (matchesMetric(itemName, "DIA")) return "DIA";
    if (matchesMetric(itemName, "HR")) return "HR";

    return null;
}

function isStanding(name, groupName) {
    if (name != null && (containsValue(name, "STANDING") || containsStandaloneKeyword(name, "STAND"))) return true;
    if (groupName != null && (containsValue(groupName, "STANDING") || containsStandaloneKeyword(groupName, "STAND"))) return true;

    return false;
}

function isAttachedItem(groupItem) {
    if (attachedItem.id != null && groupItem.id != null) return groupItem.id === attachedItem.id;
    return groupItem.name === attachedItem.name;
}

function getDifferenceRange(metric) {
    if (metric === "SYS") return sysDifferenceRange;
    if (metric === "DIA") return diaDifferenceRange;
    if (metric === "HR") return hrDifferenceRange;
    return null;
}

function getUnit(metric) {
    if (metric === "HR") return "bpm";
    return "mmHg";
}

function checkDifference(range, semi, standing, unit) {
    if (semi == null || standing == null || isNaN(semi) || isNaN(standing)) return null;

    var diff = parseFloat(standing) - parseFloat(semi);
    var absDiff = Math.abs(diff);

    if (diff > range) return "YES, decrease by " + absDiff + " " + unit + ".";
    if (diff >= 0) return "NO, decrease by " + absDiff + " " + unit + ".";
    return "NO, increase by " + absDiff + " " + unit + ".";
}

function getOrthostasisValues(metric, isRepeat) {
    var itemGroups = formJson.form.itemGroups;
    var semi = null;
    var standing = null;
    var attachedFound = false;
    var i, j, group, groupItem;

    if (!itemGroups || itemGroups.length < 1) return {semi: null, standing: null};

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            for (j = 0; j < group.items.length; j++) {
                groupItem = group.items[j];
                if (!groupItem) continue;

                if (isAttachedItem(groupItem)) return {semi: semi, standing: standing};

                if (groupItem.value === null || groupItem.value === "" || groupItem.canceled) continue;
                if (!matchesMetric(groupItem.name, metric)) continue;

                if (isStanding(groupItem.name, group.name)) {
                    if (standing === null) standing = parseFloat(groupItem.value);
                } else {
                    if (standing === null) semi = parseFloat(groupItem.value);
                }
            }
        }
    } else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled || !group.items) continue;

            for (j = group.items.length - 1; j >= 0; j--) {
                groupItem = group.items[j];
                if (!groupItem) continue;

                if (isAttachedItem(groupItem)) {
                    attachedFound = true;
                    continue;
                }

                if (!attachedFound) continue;

                if (groupItem.value === null || groupItem.value === "" || groupItem.canceled) continue;
                if (!matchesMetric(groupItem.name, metric)) continue;

                if (isStanding(groupItem.name, group.name)) {
                    if (standing === null) standing = parseFloat(groupItem.value);
                } else {
                    if (standing !== null && semi === null) {
                        semi = parseFloat(groupItem.value);
                        return {semi: semi, standing: standing};
                    }
                }
            }
        }
    }

    return {semi: semi, standing: standing};
}

try {
    var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
    var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
    var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

    logger("Group name: " + parsedGroupName);
    logger("Attached item: " + attachedItem.name);
    logger("Is repeat: " + isRepeat);

    var metric = getMetricFromAttachedItem(attachedItem.name);
    logger("Metric: " + metric);

    if (!metric) return null;

    var values = getOrthostasisValues(metric, isRepeat);
    logger("Semi: " + values.semi);
    logger("Standing: " + values.standing);

    var range = getDifferenceRange(metric);
    var unit = getUnit(metric);

    return checkDifference(range, values.semi, values.standing, unit);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
