const studyEvent = [
    "SCREENING", 
    "Screening"
];
const formName = [
    "REM_Study Reminders (SCRN)",
];
const itemName = [
    "POCBP or Partner POCBP"
];
const attachedItemCodeList = [
    "YES",
    "NO",
]
const pulledItemCodeList = [
    "YES",
    "NO",
]

const gender = formJson.form.subject.volunteer.sexMale;

if (gender) return attachedItemCodeList[1];

var form = pullForm(studyEvent, formName);
if (!form) return null;

var POCBP = pullItemFromForm(form, itemName);
if (!POCBP) return null;

log();

if (POCBP == pulledItemCodeList[0]) return attachedItemCodeList[0];
else if (POCBP == pulledItemCodeList[1]) return attachedItemCodeList[1]
return null;

function log() {
    logger("POCBP: " + POCBP);
}

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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
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