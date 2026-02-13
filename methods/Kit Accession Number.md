const formName = [
    "(*)🩸D1, D2, D3, D4, D7, D10, D14_Predose_Safety Labs + PD/PK - 6 TUBES", 
    "(*)🩸Sched B_Visit 2 W1, D0_Safeties + PK/ADA/Biomarkers_PREDOSE - 8 TUBES", 
    "(*)🩸D15 24h_Safety Labs + PD/PK - 6 TUBES",
];
const itemName = ["Process No.", "Kit Accession Number"];
const currentStudyEvent = formJson.form.studyEventName;
var form = null;
logger("Study event: " + currentStudyEvent);

const D1events = [
  "D1 (PRE)",
  "D1 (OHR)",
  "D1 (0.5hr) ±5",
  "D1 (1hr) ±5",
  "D1 (1.5hr) ±1.5",
  "D1 (2hr) ±30",
  "D1 (3hr) ±30",
  "D1 (4hr) ±30",
  "D1 (6hr) ±30",
  "D1 (8hr) ±30",
  "D1 (12hr) ±1h"
];

const D2events = [
  "D14 (PRE) -1h",
  "D14 (0.5hr) ±5",
  "D14(1hr) ±5",
  "D14 (1.5hr) ±5",
  "D14 (2hr) ±30",
  "D14 (3hr) ±30",
  "D14 (4hr) ±30",
  "D14 (6hr) ±30",
  "D14 (8hr) ±30",
  "D14(12hr) ±1h",
  "D14(18 hr) ±1h"
];

const D3events = [
    "D15 (24hr) ±1h", "D15 (36hr) ±1h"
];

if (D1events.indexOf(currentStudyEvent) !== -1) {
    logger("D1");
    form = pullForm(D1events, formName)
}
else if (D2events.indexOf(currentStudyEvent) !== -1) {
    logger("D2")
    form = pullForm(D2events, formName)
}
else if (D3events.indexOf(currentStudyEvent) !== -1) {
    logger("D3")
    form = pullForm(D3events, formName);
}

if (!form) return null;
var result = pullItemFromForm(form, itemName);

log()

return result;

function log() {
    logger("Day " + D1events[0] + ": " + D1events.indexOf(currentStudyEvent) !== -1);
    logger("Day " + D2events[0] + ": " + D2events.indexOf(currentStudyEvent) !== -1);
    logger("Day " + D3events[0] + ": " + D3events.indexOf(currentStudyEvent) !== -1);
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
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[completedForm.length - 1];
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