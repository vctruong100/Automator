/* jshint strict: false */

// Version: v1
// Purpose: Evaluates FSH and pregnancy test result logic.

var formName = formJson.form.name;

var studyEvent = [
    "SCREENING",
    "Screening"
];
var formNames = [
    "REP_Reproductive Status and Contraception – Female",
    "REP_Female Reproductive Status and Contraceptive Use",
    "♀️ REP_CONTRACEPTIVE & REPRODUCTIVE STATUS (Female)"
];
var childbearingItem = [
    "Is the female subject of childbearing potential?",
    "Childbearing potential?"
];
var statusItem = [
    "Female Reproductive Status",
    "Female subject is considered non-childbearing potential due to"
];

var gender = formJson.form.subject.volunteer.sexMale;
var age = formJson.form.subject.volunteer.age;

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item;
        }
    }
    return null;
}

function checkForm(studyevent, form) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms, true);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm[0];
    }
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

try {
    logger(age);
    logger(gender);
    if (gender) return "None";

    var form = pullForm(studyEvent, formNames);
    if (!form) return null;

    var childbearing = pullItemFromForm(form, childbearingItem);
    if (childbearing && childbearing.value !== null && childbearing.value == "Y") return "Pregnancy";

    var status = pullItemFromForm(form, statusItem);
    if (status && status.value !== null && (status.value == status.codeListItems[2].codedValue)) return "FSH";
    else if (status && status.value !== null) return "Pregnancy";

    return "None";
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
