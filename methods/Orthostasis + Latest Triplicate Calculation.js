/* jshint strict: false */

// Version: v3
// Description: Calculates the difference between a standing vital-sign value and the latest matching supine or semi-recumbent triplicate value for orthostasis workflows.
// Triplicate item groups (non-repeat): VS_READ #1 TRIPLICATE, VS_READ #2 TRIPLICATE, VS_READ #3 TRIPLICATE
// Standing item group (non-repeat): VS_READ #4 STANDING
// Triplicate item groups (repeat): VS_READ #5 TRIPLICATE (REPEAT), VS_READ #6 TRIPLICATE (REPEAT), VS_READ #7 TRIPLICATE (REPEAT)
// Standing item group (repeat): VS_READ #8 STANDING (REPEAT)
// Used in Merck MK-4082-005

var item = itemJson.item;

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
    var index, before, after;

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

function getMetricFromItemName(itemName) {
    if (matchesMetric(itemName, "SYS")) return "SYS";
    if (matchesMetric(itemName, "DIA")) return "DIA";
    if (matchesMetric(itemName, "HR")) return "HR";

    return null;
}

function isCalculatedVitalItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;
    if (containsValue(itemName, "MEDIAN")) return true;
    if (containsValue(itemName, "CHANGE")) return true;
    if (containsValue(itemName, "DIFFERENCE")) return true;
    if (containsStandaloneKeyword(itemName, "DIFF")) return true;
    if (containsValue(itemName, "ORTHOSTATIC")) return true;
    if (containsValue(itemName, "ORTHOSTASIS")) return true;

    return false;
}

function getLatestTriplicateValue(form, metric, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var values = [];
    var latest = null;
    var i, group, j, groupItem, num;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        if (isRepeat) {
            if (!containsValue(group.name, "repeat")) continue;
        } else {
            if (containsValue(group.name, "repeat")) continue;
        }

        if (!containsValue(group.name, "triplicate")) continue;

        for (j = 0; j < group.items.length; j++) {
            groupItem = group.items[j];
            if (!groupItem || groupItem.value === null || groupItem.value === "" || groupItem.canceled) continue;

            if (isCalculatedVitalItem(groupItem.name)) continue;
            if (!matchesMetric(groupItem.name, metric)) continue;

            num = parseFloat(groupItem.value);
            if (!isNaN(num)) {
                values.push(num);
                latest = num;
                logger("Captured triplicate value from group: " + group.name + " item: " + groupItem.name + " value: " + num);
            }
        }
    }

    if (values.length === 0) {
        logger("No triplicate values found for metric: " + metric);
        return null;
    }

    logger("Triplicate values for " + metric + ": " + values.join(", "));
    logger("Latest triplicate value for " + metric + ": " + latest);
    return latest;
}

function getStandingValue(form, metric, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var i, group, j, groupItem, num;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        if (isRepeat) {
            if (!containsValue(group.name, "repeat")) continue;
        } else {
            if (containsValue(group.name, "repeat")) continue;
        }

        if (!containsValue(group.name, "standing")) continue;

        for (j = 0; j < group.items.length; j++) {
            groupItem = group.items[j];
            if (!groupItem || groupItem.value === null || groupItem.value === "" || groupItem.canceled) continue;

            if (isCalculatedVitalItem(groupItem.name)) continue;
            if (!matchesMetric(groupItem.name, metric)) continue;

            num = parseFloat(groupItem.value);
            if (!isNaN(num)) {
                logger("Captured standing value from group: " + group.name + " item: " + groupItem.name + " value: " + num);
                return num;
            }
        }
    }

    return null;
}

function getOrthostasis(form, metric, isRepeat) {
    var semi = getLatestTriplicateValue(form, metric, isRepeat);
    if (semi === null) {
        logger("No semi/triplicate latest value found for metric: " + metric);
        return null;
    }

    var standing = getStandingValue(form, metric, isRepeat);
    if (standing === null) {
        logger("No standing value found for metric: " + metric);
        return null;
    }

    logger("Final Semi (latest triplicate value) for " + metric + ": " + semi);
    logger("Final Standing for " + metric + ": " + standing);

    return standing - semi;
}

try {
    var isRepeat = false;
    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;

    logger("Group name: " + parsedGroupName);
    if (containsValue(parsedGroupName, "repeat")) isRepeat = true;
    logger("Attached item: " + item.name);

    var metric = getMetricFromItemName(item.name);
    logger("Metric: " + metric);
    if (!metric) return null;

    var difference = getOrthostasis(formJson, metric, isRepeat);

    logger("Difference: " + difference);
    if (difference === null || isNaN(difference)) return null;
    return difference.toFixed(0);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
