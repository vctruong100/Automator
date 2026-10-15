/* jshint strict: false */

var item = itemJson.item;
var studyEvent = ["Subject"];
var pullFormName = ["Subject Eligibility", "Subject Eligibility v2.0"];

var form = pullForm(studyEvent, pullFormName);

if (!form) return null;

var result = pullItemFromForm(form, item.name);

return result !== null ? result : null;

function normalizeItemName(name) {
    if (!name) return "";

    return name.toString()
        .replace(/\s+v\d+(?:\.\d+)*$/i, "")
        .replace(/\s+/g, "")
        .toLowerCase();
}

function pullItemFromForm(form, targetItemName) {
    var itemGroups = form.form.itemGroups;
    var normalizedTarget = normalizeItemName(targetItemName);

    if (!itemGroups || itemGroups.length === 0) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];

        if (!group || group.canceled || !group.items) continue;

        for (var j = 0; j < group.items.length; j++) {
            var sourceItem = group.items[j];

            if (
                !sourceItem.canceled &&
                normalizeItemName(sourceItem.name) === normalizedTarget &&
                sourceItem.value !== null &&
                sourceItem.value !== undefined &&
                sourceItem.value !== ""
            ) {
                return sourceItem.value;
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

    return null;
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);

    if (!completedForm || completedForm.length === 0) return null;

    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) return [];

    var completedForms = [];

    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];

        if (
            formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (
                formData.form.dataCollectionStatus == "Complete" ||
                (
                    INCLUDE_NONCONFORMANT_DATA == true &&
                    formData.form.dataCollectionStatus == "Nonconformant"
                ) ||
                formData.form.dataCollectionStatus == "Incomplete"
            )
        ) {
            completedForms.push(formData);
        }
    }

    return completedForms;
}