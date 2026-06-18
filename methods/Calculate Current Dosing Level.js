/* jshint strict: false */

// Version: v1
// Purpose: Determines current dosing level from kit label and tablet count.

var formNames = [
    "🟡IP_ MK-4082/Placebo Administration (D36-D84)"
]

var studyeventMap = {
    "Wk7 Day 43": "Wk6 Day 38",
    "Wk8 Day 50": "Wk7 Day 43",
    "Wk9 Day 57": "Wk8 Day 50",
    "Wk10 Day 64": "Wk9 Day 57",
    "Wk11 Day 71": "Wk10 Day 64",
    "Wk12 Day 78": "Wk11 Day 71",
    "Wk13 Day 85": "Wk12 Day 78",
};

var placeboItem = [
    "Label Description MK-4082/Placebo"
]

var numTabletsItem = [
    "Number of Units Taken - MK-4082/Placebo",
]
var studyevent = formJson.form.studyEventName;
var expectedDays = 0;

var screeningNumber = formJson.form.subject.screeningNumber;
logger("Subject ID: " + screeningNumber);

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

try {
    var expectedDays = 0;
    var day = parseInt(studyevent.split(" ")[2]);
    var prevDay = studyeventMap[studyevent];
    var prevPrevVisit = studyeventMap[prevDay];

    logger("Studyevent: " + studyevent)
    logger("Previous visit: " + prevDay)
    var prevVisitform = pullForm([prevDay], formNames);

    var placebo = pullItemFromForm(prevVisitform, placeboItem);
    var tablet = pullItemFromForm(prevVisitform, numTabletsItem);

    if (!placebo || !tablet || placebo == null || tablet == null) {
        var prevPrevVisitform = pullForm([prevPrevVisit], formNames);
        placebo = pullItemFromForm(prevPrevVisitform, placeboItem);
        tablet = pullItemFromForm(prevPrevVisitform, numTabletsItem);
    }
    logger("Placebo: " + placebo);
    logger("Tablet: " + tablet);

    var parts = placebo.split(" ");

    var count = parseInt(parts[0], 10);
    var description = parts.slice(1).join(" ");

    logger("Count: " + count);
    logger("Description: " + description);

    var total = parseInt(count) * tablet;

    return total + " " + description;

} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
}

