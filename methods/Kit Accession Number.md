// Version: v1
// Purpose: Generates or validates kit accession numbers with D1-prefix logic.

const formNames = [
    "(*)🩸D1, D2, D3, D4, D7, D10, D14_Predose_Safety Labs + PD/PK - 6 TUBES", 
    "(*)🩸Sched B_Visit 2 W1, D0_Safeties + PK/ADA/Biomarkers_PREDOSE - 8 TUBES", 
    "(*)🩸D15 24h_Safety Labs + PD/PK - 6 TUBES",
    "(*)🩸D1, D2, D3, D4, D7, D10, D14_Predose_Safety Labs + PD/PK - 6 TUBES V2",
    "(*)🩸D15 24h_Safety Labs + PD/PK - 6 TUBES V2",
    "(*)LAB_🩸Part 2_D1 PREDOSE, D3,  D4, D18 PREDOSE _Safeties + PK - 3 TUBES",
    "(*) LAB_🩸Part 2_D5 PREDOSE, D6 PREDOSE, D7, D8, D11, D14, D19 24h - 6 TUBES",
];

const d1FormNames = [
    "(*)🩸D1, D2, D3, D4, D7, D10, D14_Predose_Safety Labs + PD/PK - 6 TUBES", 
    "(*)🩸D1, D2, D3, D4, D7, D10, D14_Predose_Safety Labs + PD/PK - 6 TUBES V2",
]
const itemName = ["Process No.", "Kit Accession Number"];
const currentStudyEvent = formJson.form.studyEventName;

var form = null;
logger("Study event: " + currentStudyEvent);

const D1events = [
    "Day 1 P2 (0.5hr)",
    "D1 (PRE)",
];
const D5events = [
    "D5 (PRE) -1h",
]
const D14events = [
    "D14 (PRE) -1h",
];

const D15events = [
    "D15 (24hr)"
];
const D18events = [
    "Day 18 (PRE)"
]

try {
    if (containsValue(currentStudyEvent, "day 18")) {
        logger('P5 D18');
        form = pullForm(D18events, formNames);
    }
    else if (containsValue(currentStudyEvent, "day 5")) {
        logger("P2 D5");
        form = pullForm(D5events, formNames);
    }
    else if (containsValue(currentStudyEvent, "d15")) {
        form = pullForm(D15events, formNames);
    }
    else if (containsValue(currentStudyEvent, "d14")) {
        form = pullForm(D14events, formNames);
    }
    else if (containsValue(currentStudyEvent, "day 1")) {
        form = pullForm(D1events, formNames);
    }
    else if (containsValue(currentStudyEvent, "d1")) {
        form = pullForm(D1events, d1FormNames);
    }

    if (!form) return null;
    var result = pullItemFromForm(form, itemName);

    return result;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Iterates through study events and form names to find the first matching completed form.
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = formNameList.length - 1; j >= 0; j--) {
            var matchedForm = checkForm(studyeventList[i], formNameList[j]);
            if (matchedForm) return matchedForm;
        }
    }
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

// Checks if the input string contains the specified keyword (case-insensitive). Returns false for null/undefined input.
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

<<<<<<< Updated upstream
    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}
=======
    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}
>>>>>>> Stashed changes
