const studyevents = {
    "D1": "D -1",
    "D2": "D1",
    "D3": "D2",
    "D4": "D3",
    "D5": "D4",
    "D6": "D5",
    "D7": "D6",
    "D8": "D7",
    "D9": "D8"
};
const currentEvent = formJson.form.studyEventName;

const mealForms = [
    'STOP SNACKS', "Snack End", "🍽️ MEAL_SNACK END",
    'STOP DINNER', "Dinner End", "🍽️ MEAL_DINNER END",
    'STOP LUNCH', "Lunch End", "🍽️ MEAL_LUNCH END",
    'STOP BREAKFAST', "Breakfast End", "BREAKFAST END",
];

var parsedEvent = null;
var lowerEvent = currentEvent ? currentEvent.toLowerCase() : "";

// Filter study event names
for (var key in studyevents) {
    if (lowerEvent.indexOf(key.toLowerCase()) !== -1) {
        parsedEvent = key;
        break;
    }
}
logger("Parsed event: " + parsedEvent)
if (!parsedEvent) return null;

const prevEvent = studyevents[parsedEvent];
logger("Previous event: " + prevEvent);
var form = pullForm([prevEvent], mealForms);
if (!form) {
    var preprevEvent = studyevents[prevEvent]; 
    form = pullForm([preprevEvent], mealForms)
}
if (!form) return null;
var mealTime = form.form.itemGroups[0].items[0].value;
if (mealTime && mealTime !== null) return mealTime;

return null;

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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}