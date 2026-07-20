/* jshint strict: false */

// Version: v4
// Purpose: Computes the numeric orthostatic difference between supine/semi and standing vitals.
//          Uses keyword matching on item and group names instead of hardcoded item-name lists.

var formNames = [
    "❤️ VS_TRIPLICATE (HR, BP, RR, ORAL TEMP.) SCRN"    
]
var studyevent = formJson.form.studyEventName;

var attachedItem = itemJson.item;

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

function calculateDifference(semi, standing) {
    if (semi == null || standing == null || isNaN(semi) || isNaN(standing)) return null;

    var diff = parseFloat(standing) - parseFloat(semi);
    return diff.toFixed(0);
}
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false &&
            (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') ||
                formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        }
    }
    return completedForms;
}

function pullItemFromForm(formJsonValue, metric, isRepeat) {
    logger("Target metric: " + metric);
    var itemGroups = formJsonValue.form.itemGroups;
    var group, item, i, j;
    if (!itemGroups || itemGroups.length < 1) return null;

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
    
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];
                if (!item) continue;
    
                if (matchesMetric(item.name, metric) && item.value !== null && item.value !== "") return item.value;
            }
        }
    } else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
    
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];
                if (!item) continue;
    
                if (matchesMetric(item.name, metric) && item.value !== null && item.value !== "") return item.value;
            }
        }
    }
    return null;
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

var rawGroupName = getItemDataContextByItemDataId(attachedItem.id);
var parsedGroupName = JSON.parse(rawGroupName).foundItemGroupName;
var isRepeat = parsedGroupName ? containsValue(parsedGroupName, "repeat") : false;

logger("Group name: " + parsedGroupName);
logger("Attached item: " + attachedItem.name);
logger("Is repeat: " + isRepeat);

var metric = getMetricFromAttachedItem(attachedItem.name);
logger("Metric: " + metric);

if (!metric) return null;

var standing, semi;

if (containsValue(formJson.form.name, "standing")) {
    logger("Pull from screening vital form")
    standing = pullItemFromForm(formJson, metric, isRepeat);

    var form = pullForm([studyevent], formNames);
    if (!form) return null;
    semi = pullItemFromForm(form, metric, true);

} else {
    var values = getOrthostasisValues(metric, isRepeat);
    semi = values.semi;
    standing = values.standing;
}

logger("Semi: " + semi);
logger("Standing: " + standing);

return calculateDifference(semi, standing);