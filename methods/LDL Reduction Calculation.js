/* jshint strict: false */

// Version: v1
// Purpose: Calculates LDL reduction percentage from baseline and current values.

var studyEventNames = [
    "Day 1",
    "Screening",
    "SCREENING",
]

var formName = [
    "LDL Continuation (Baseline LDL)"
]

var baselineItem = [
    "Baseline LDL (Day 1)"
]

var valueItem = [
    "LDL-C Value at Week 10",
    "LDL-C Value at Week 8",
]

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
        }
    }
    return false;
}
function calculateReduction(baseline, value) {
    var parsedBaseline = parseInt(baseline);
    var parsedValue = parseInt(value);
    return Math.round(100 - ((parsedValue * 100) / parsedBaseline))
}

function pullItemFromForm(form, targetItem, groupid) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        if (groupid !== null && group.id !== groupid) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
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
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            keepers.push(formData);
        } else {

        }
    }
    return keepers;
}

function getItemGroupNameId(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                return [group.name, group.id];
            }
        }
    }
    return null;
}

try {
    var form = pullForm(studyEventNames, formName);
    if (!form) return null;

    var item = itemJson.item;
    var groupName, groupID = getItemGroupNameId(formJson);

    var baseline = pullItemFromForm(form, baselineItem, null);
    var value = pullItemFromForm(formJson, valueItem, groupID);

    if (!baseline || baseline == null || !value || value == null) return null;

    return String(calculateReduction(baseline, value));
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
