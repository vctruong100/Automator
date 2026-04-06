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

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = formNameList.length - 1; j >= 0; j--) {
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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}