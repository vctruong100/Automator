const sigfig = itemJson.item.significantDigits;

// Function to pull a specific item from form (not by first item name)
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
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

// By last to first
function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "" && !isNaN(item.value)) return item.value;
        }
    }
    return null;
}

//Use this to get item from same item group:
function getItemGroupID(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
    var group = form.form.itemGroups[i];
    var items = group.items;
    if (!items || items.length < 1) continue;

    for (var j = 0; j < items.length; j++) {
        var it = items[j];
        if (it.id === item.id) {
            groupid = group.id;
            break;
        }
    }
    if (groupid) break;
    }
}

function getItemValueFromSameGroup(form, itemName) {
    var value = null;
    for (var i = 0; i < form.itemGroups.length; i++) {
        var group = form.itemGroups[i];
        if (group.id !== groupid) continue;

        var items = group.items;
        for (var j = 0; j < items.length; j++) {
            var item = items[j];
            if (item && item.name === itemName) {
                value = item.value;
                if (value && value !== null) return value;
            }
        }
    }
    return null; 
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[completedForm.length - 1];
}


// for matching timepoint forms
function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    var clean = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && 
            formData.form.itemGroups[0].canceled == false && 
            (formData.form.dataCollectionStatus == 'Complete' || 
            (INCLUDE_NONCONFORMANT_DATA == true && 
            formData.form.dataCollectionStatus == 'Nonconformant') || 
            formData.form.dataCollectionStatus == "Incomplete")) 
        {
            keepers.push(formData);
        } else {}
    }
    for (var j = 0; j < keepers.length; j++) {
        var unfilter = keepers[j];
        if (unfilter.form.timepoint == timepoint) clean.push(keepers[j]);
    }
    return clean;
}


// set the name of the item which contains the desired value
var relatedItemDataContext = JSON.parse(getRelatedItemDataContext('BE_Date of sample 2'))
var itemDataContext = JSON.parse(getRelatedItemDataContext());
var transferBarcodes = relatedItemDataContext.transferBarcodes;
logger(JSON.stringify(relatedItemDataContext, null, 2));
return transferBarcodes;
