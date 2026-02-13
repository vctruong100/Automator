const formName = [
    "BM_Waist/Hip Circumference"
];
const screeningStudyEvent = [
    "Screening", 
    "SCREENING"
]
const screeningFormName = [
    "BM_Waist/Hip Circumference (SCRN)"
];

const waistItem1 = [
    "Waist Circumference #1"
];
const waistItem2 = [
    "Waist Circumference #2"
];
const hipItem1 = [
    "Hip Circumference #1"
];
const hipItem2 = [
    "Hip Circumference #2"
];

const studyEvents = {
    "Visit 2 Week 1 Day 0": "Screening",
    "Visit 3 Week 2 Day 7": "Visit 2 Week 1 Day 0",
    "Visit 4 Week 3 Day 14": "Visit 3 Week 2 Day 7",
    "Visit 5 Week 5 Day 28": "Visit 4 Week 3 Day 14",
    "Visit 6 Week 8 Day 49": "Visit 5 Week 5 Day 28",
    "Visit 7 Week 11 Day 70": "Visit 6 Week 8 Day 49",
    "Visit 8 Week 15 Day 98": "Visit 7 Week 11 Day 70",
    "Visit 9 Week 20 Day 133 (Final Dose)": "Visit 8 Week 15 Day 98",
    "EOT Visit 10 Week 21 Day 140": "Visit 9 Week 20 Day 133 (Final Dose)",
    "EOS Visit 11 Week 24 Day 161": "EOT Visit 10 Week 21 Day 140"
}

const itemName = itemJson.item.name;
const currentEvent = formJson.form.studyEventName;

var newEvent = studyEvents[currentEvent];
var form = null;
if (screeningStudyEvent.indexOf(newEvent) !== -1) {
    form = pullForm([newEvent], screeningFormName);
} else {
    form = pullForm([newEvent], formName);
}
if (!form) return null;
if (itemName == waistItem1) return pullItemFromForm(form, waistItem1);
if (itemName == waistItem2) return pullItemFromForm(form, waistItem2);
if (itemName == hipItem1) return pullItemFromForm(form, hipItem1);
if (itemName == hipItem2) return pullItemFromForm(form, hipItem2);

return null;

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