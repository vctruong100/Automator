/* jshint strict: false */

// Version: v1
// Purpose: Computes expected tablet return count from dosing regimen.

var formNames = [
    "REV_Review Questions (w/ Glucose monitoring) v3.0",
    "REV_Review Questions (w/ Glucose monitoring) (D36, D43, D50, D57, D64, D71, D78, D80, D81, D98, Early discontinuation) v3.0"
]

var itemName = [
    "Visit Date",
]

var studyeventMap = {
    "Wk7 Day 43": "Wk6 Day 38",
    "Wk8 Day 50": "Wk7 Day 43",
    "Wk9 Day 57": "Wk8 Day 50",
    "Wk10 Day 64": "Wk9 Day 57",
    "Wk11 Day 71": "Wk10 Day 64",
    "Wk12 Day 78": "Wk11 Day 71"
};

var dispensedItem = [
    "Number of tablets dispensed:",
]

var prescribedItem = [
    "Number of tablets prescribed per day:",
]

var studyevent = formJson.form.studyEventName;
var expectedDays = 0;

var screeningNumber = formJson.form.subject.screeningNumber;
logger("Subject ID: " + screeningNumber);

try {
    var day = parseInt(studyevent.split(" ")[2]);
    var prevDay = studyeventMap[studyevent];
    logger("studyevent: " + studyevent + ", previous visit: " + prevDay)
    var pastVisitform = pullForm([prevDay], formNames);
    var currentVisitform = pullForm([studyevent], formNames);
    if (!pastVisitform) {
        if (day > 43 && day <= 84) expectedDays = 7;
        else if (day <= 43) expectedDays = 4;
    }
    else {
        var pastdate = pullItemFromForm(pastVisitform, itemName);
        var currentdate = pullItemFromForm(currentVisitform, itemName);
        logger(pastdate);
        logger(currentdate);

        var pastday = pastdate.split("-")[2];
        var currentday = currentdate.split("-")[2];
        expectedDays = currentday - pastday - 1;
    }
    var dispensed = pullItemFromForm(formJson, dispensedItem);
    var prescribed = pullItemFromForm(formJson, prescribedItem);
    if (!dispensed || dispensed == null || !prescribed || prescribed == null) return null;

    logger("Current Day: " + day);
    logger("expectedDays: " + expectedDays);

    var totalPrescribed = expectedDays * prescribed;
    logger("Tablets Dispensed: " + dispensed)
    logger("Tablet Prescribed: " + totalPrescribed)
    var expectedReturn = dispensed - totalPrescribed;
    return String(expectedReturn.toFixed(0));

} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
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
