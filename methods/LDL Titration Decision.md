// Version: v1
// Purpose: Determines LDL titration decision based on reduction and thresholds.

const studyEventNames = [
    "Day 1",
    "Screening",
    "SCREENING",
]

const formName = [
    "LDL Continuation (Baseline LDL)"    
]

const baselineItem = [
    "Baseline LDL (Day 1)"    
]

const valueItem = [
    "LDL-C Value at Week 10",
    "LDL-C Value at Week 8",
]
try {
    var form = pullForm(studyEventNames, formName);
    if (!form) return null;

    const item = itemJson.item;
    var groupName, groupID = getItemGroupName(formJson);

    var baseline = pullItemFromForm(form, baselineItem, null);
    var value = pullItemFromForm(formJson, valueItem, groupID);

    if (!baseline || baseline == null || !value || value == null) return null;
    var percentRed = calculateReduction(baseline, value);

    if (containsValue(groupName, "no ascvd")) {
        if (percentRed >= 50 && value < 55) return item.codeListItems[1].codedValue;
        else if (percentRed < 50 || value >= 55) return item.codeListItems[0].codedValue;
    }
    else {
        if (percentRed >= 50 && value < 70) return item.codeListItems[1].codedValue;
        else if (percentRed < 50 || value >= 70) return item.codeListItems[0].codedValue;
    }

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Performs calculation: calculateReduction.
function calculateReduction(baseline, value) {
    var parsedBaseline = parseInt(baseline);
    var parsedValue = parseInt(value);
    return Math.round(100 - ((parsedValue * 100) / parsedBaseline))
}

// Iterates through study events and form names to find the first matching completed form.
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
}

// Retrieves the first completed (or nonconformant) form instance from a study event.
function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

// Filters an array of form data to return only entries with valid completion status (Complete, Nonconformant, or Incomplete).
function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var completedForms = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' || 
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            completedForms.push(formData);
        } else {

        }
    }
    return completedForms;
}

// Searches a form's item groups for an item matching the target name and returns its value or the item object.
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}

// Checks if the input string contains the specified keyword (case-insensitive). Returns false for null/undefined input.
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

// Retrieves the name of the item group containing the current item.
function getItemGroupName(form) {
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;
    
        for (var j = 0; j < items.length; j++) {
            var currentItem = items[j];
            if (currentItem.id === item.id) {
                return group.name, group.id;
            }
        }
    }
    return null;
}