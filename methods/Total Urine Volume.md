// Version: v1
// Purpose: Calculates total urine volume across collection intervals.

const currentEvent = formJson.form.studyEventName;
const studyevents = {
    "Day 1": "Day-1",
    "Day 2": "Day 1",
    "Day 3": "Day 2",
    "Day 4": "Day 3",
    "Day 5": "Day 4",
    "Day 6": "Day 5",
    "Day 7": "Day 6",
    "Day 8": "Day 7",
    "Day 9": "Day 8",
    "Day 10": "Day 9",
    "Day 11": "Day 10",
    "Day 12": "Day 11",
    "Day 13": "Day 12",
    "Day 18": "Day 17"
}

const mealForms = [
    'STOP SNACKS', "Snack End", 
    'STOP DINNER', "Dinner End",
    'STOP LUNCH', "Lunch End",
    'STOP BREAKFAST', "Breakfast End",
];

try {
    const prevEvent = studyevents[currentEvent];
    var form = pullForm([prevEvent], mealForms);
    if (!form) return null;

    var mealTime = form.form.itemGroups[0].items[0].value;
    if (mealTime && mealTime !== null) return mealTime;

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Searches a form's item groups for an item matching the target name and returns its value or the item object.
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
<<<<<<< Updated upstream
    return keepers;
}
=======
    return completedForms;
}
>>>>>>> Stashed changes
