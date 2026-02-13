const baselineForms = [
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)"
];
const baselineFormStudyEvents = [
    "Visit 2 Week 1 Day 0"
];
const baselineItem = [
    "Baseline QTcF"
];
const qtcfItems = [
    "QTcF #1", 
    "QTcF #2", 
    "QTcF #3"
];
const item = itemJson.item;
const protocol_avg_qtcf_range = 500;
const protocol_difference = 60;

var maxCount = 3; // max number of qtcf items
var qtcfOOR = false;
var list = [];

var form = pullForm(baselineFormStudyEvents, baselineForms);
if (!form) return true;
var baseline = pullItemFromForm(form, baselineItem);

list = populateList(formJson, qtcfItems, list);

if (list.length < 3 || !baseline || baseline == null) return null;

var result = calculateAverageOOR();
log()

return result;

function log() {
    logger("Baseline: " + baseline);
    logger("List: " + list);
    logger("Average: " + average + ", Protocol range: " + protocol_avg_qtcf_range);
    logger("QTcF OOR: " + qtcfOOR);
    logger("Result: " + result);
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function populateList(form, targetItem, list) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;

	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                    if (list.length >= maxCount) return list;
                }
            }
        }
    }
    return list;
}

function calculateAverageOOR() {
    var total = 0;
    for (var i = 0; i < list.length; i++) {
        if (checkBaseline(parseFloat(list[i]))) {
            qtcfOOR =  true;
        }
        total += list[i];
    }
    average = Math.round(total / list.length);
    if (average > protocol_avg_qtcf_range || qtcfOOR) {
        return "Y";
    }
    else return "N";
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
    return null;
}

function checkBaseline(itemValue) {
    if (baseline !== null && baseline !== undefined) {
        var diff = (itemValue - baseline)
        logger("Item value: " + itemValue + ", Baseline: " + baseline + ", Difference: " + diff)
        if (diff >= 60) {
            return true;
        }
        return false;
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