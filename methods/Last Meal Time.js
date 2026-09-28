const currentEvent = formJson.form.studyEventName;
const studyevents  = {
    "D1 (PRE)": "Day -1",
    "D2": "D1",
    "D3": "D2",
    "D4": "D3",
    "D5": "D4",
    "D6": "D5",
    "D7": "D6",
    "D8": "D7",
    "D9": "D8",
    "D10": "D9",
    "D11": "D10",
    "D12": "D11",
    "D13": "D12",
    "D14": "D13",
    "D15": "D14",
    "D16": "D15",
    "D17": "D16",
    "D18": "D17",
    "D5 (PRE)": "D4",
    "D6 (PRE)": "D5",
    "D7 (PRE)": "D6",
    "D8 (PRE)": "D7",
    "D9 (PRE)": "D8",
    "D10 (PRE)": "D9",
    "D11 (PRE)": "D10",
    "D12 (PRE)": "D11",
    "D13 (PRE)": "D12",
    "D14 (PRE)": "D13",
    "Day 15 (PRE)": "D14",
    "Day 16 (PRE)": "D15",
    "Day 17 (PRE)": "D16",
    "Day 18 (PRE)": "D17",
    "D19": "D18",
    "D20": "D19",
    "D21": "D20",
    "Day 22": "D21"
};

const mealForms = [
    'STOP SNACKS', "Snack End", "MEAL STOP SNACKS",
    'STOP DINNER', "Dinner End", "MEAL STOP DINNER",
    'STOP LUNCH', "Lunch End", "MEAL STOP LUNCH",
    'STOP BREAKFAST', "Breakfast End", "MEAL STOP BREAKFAST", "STOP BREAKFAST 🔵"
];

try {
    var parsedEvent = null;
    var lowerEvent = currentEvent ? currentEvent.toLowerCase() : "";

    for (var key in studyevents) {
        if (lowerEvent.indexOf(key.toLowerCase()) !== -1) {
            parsedEvent = key;
            break;
        }
    }

    logger("Parsed event: " + parsedEvent);

    if (!parsedEvent) return null;

    return findLastMeal(parsedEvent);

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
        }
    }

    return false;
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];

        if (!group || group.canceled || !group.items) continue;

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];

            if (
                containsItemName(targetItem, item.name) &&
                item.value !== null
            ) {
                return item.value;
            }
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

    return null;
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);

    if (!completedForm || completedForm.length === 0) return null;

    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) return [];

    var keepers = [];

    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];

        if (
            formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (
                formData.form.dataCollectionStatus == "Complete" ||
                (
                    INCLUDE_NONCONFORMANT_DATA == true &&
                    formData.form.dataCollectionStatus == "Nonconformant"
                ) ||
                formData.form.dataCollectionStatus == "Incomplete"
            )
        ) {
            keepers.push(formData);
        }
    }

    return keepers;
}

function findLastMeal(startEvent) {
    var eventToCheck = studyevents[startEvent];

    while (eventToCheck) {
        logger("Checking event for last meal: " + eventToCheck);

        var form = pullForm([eventToCheck], mealForms);

        if (form) {
            var mealTime = form.form.itemGroups[0].items[0].value;
            if (mealTime !== null) {
                logger("Last meal found in " + eventToCheck + ": " + mealTime);
                return mealTime;
            }
        }

        eventToCheck = studyevents[eventToCheck];
    }

    logger("No previous meal time found.");
    return null;
}

