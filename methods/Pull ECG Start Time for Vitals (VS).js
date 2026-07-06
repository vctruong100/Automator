/* jshint strict: false */

    }
    return false;
}

// Version: v1
// Purpose: Pull ECG Start Time for Vitals (VS) from the most recent completed ECG form

var formNames = [
    "❤️ (PREDOSE) 12-LEAD ECG (SINGLE) V2.0",
    "❤️ 12-LEAD ECG (SINGLE) V2.0",
    "❤️ 12-LEAD ECG (SINGLE) V1.0"
]

var itemNames = [
    "REPEAT START SUPINE",
    "START SUPINE'"
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
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item.value;
            }
        }
    }
    return null;
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

function checkFormMostRecent(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[completedForm.length - 1];
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

var studyEventName = formJson.form.studyEventName;

try {
    logger("Study event: " + studyEventName);
    var form = pullForm([studyEventName], formNames);
    if (!form) return null;

    return pullItemFromForm(form, itemNames);
} catch (e) {
    logger("Error: " + e);
    return null;
}