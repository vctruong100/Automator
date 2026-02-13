const baselineStudyEvent = [
    "Period 1 Day 1",
    "Visit 2 Week 1 Day 0",
];
const baselineFormName = [
    "ECG_Single 12 - Lead ECG + Triplicate",
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)",
    "ECG_Single 12 - Lead ECG + Triplicate (baseline - PREDOSE)"
];

const qtcfItem = [
    "QTcF (protocol range)",
    "QTcF (baseline)",
    "Baseline QTcF"
];
const sigfig = itemJson.item.significantDigits;

var form = pullForm(baselineStudyEvent, baselineFormName);
if (!form) return null;

var baseline = pullBaselineItemFromForm(form, qtcfItem);
var qtcfValue = pullItemFromForm(formJson, qtcfItem)
if (!qtcfValue || qtcfValue == null || !baseline || baseline == null) return null;
logger("qtcfValue: " + qtcfValue);
logger("Baselien: " + baseline);

var difference = Math.round(qtcfValue - baseline).toFixed(sigfig);
logger("Difference: " + difference);
return difference;

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function pullBaselineItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled) return parseInt(item.value);
        }
    }
    return null;
}
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled) return parseInt(item.value);
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