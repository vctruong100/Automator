var studyevents = [
    "Wk12 Day 78"
]
var doseForm = [
    "IP_Rosuvastatin Administration"   
]

var doseItem = [
    "Start Date/Time Rosuvastatin"    
]

var collectedTimeItem = [
    "AE_Onset_Date/Time"
]

const difference = 24; // in hours
var methodType = "B";

var subjectScreeningNumber = formJson.form.subject.screeningNumber;
logger(subjectScreeningNumber)

// ======== Don't modify ========
var startTime = pullForm(studyevents, doseForm, doseItem);
var endTime = pullItemFromForm(formJson, collectedTimeItem);

if (!startTime || startTime.value == null || !endTime || endTime.value == null) return "N";
logger("startTime Time: " + startTime.value);
logger("Collected Time: " + endTime.value);

var startTimeMs = getDateValueMs(startTime);
var endTimeMs = getDateValueMs(endTime);

logger("Start Time MS: " + startTimeMs);
logger("End Time MS: " + endTimeMs);

if (startTimeMs == null || endTimeMs == null) {
    logger("Unable to determine date/time values.");
    return "N";
}

var differenceInMins;

if (methodType === "A") {

    // Method A: true difference, then floor
    var diffMs = endTimeMs - startTimeMs;
    differenceInMins = Math.floor(diffMs / (1000 * 60));

    logger("Method A used");

} else {

    // Method B: floor each first, then subtract
    var startMin = Math.floor(startTimeMs / (1000 * 60));
    var endMin = Math.floor(endTimeMs / (1000 * 60));

    differenceInMins = endMin - startMin;

    logger("Method B used");
    logger("Start (hrs): " + (startMin / 60));
    logger("End (hrs): " + (endMin / 60));
}

logger("Diff (hrs): " + (differenceInMins / 60));

if (differenceInMins < 0) return "N";

if (differenceInMins >= (difference * 60)) {
    return "N";
}
else if (differenceInMins < (difference * 60)) return "Y";

return null;

function getDateValueMs(item) {
    if (!item) return null;

    if (item.dateValueMs != null) {
        return item.dateValueMs;
    }

    if (!item.value) {
        return null;
    }

    logger("dateValueMs missing. Attempting date-only parse for value: " + item.value);

    var datePart = item.value.split("T")[0];

    if (!datePart) {
        return null;
    }

    var datePieces = datePart.split("-");

    if (datePieces.length !== 3) {
        return null;
    }

    var year = parseInt(datePieces[0], 10);
    var month = parseInt(datePieces[1], 10) - 1;
    var day = parseInt(datePieces[2], 10);

    var parsedMs = new Date(year, month, day, 0, 0, 0, 0).getTime();

    logger("Date-only parsed MS: " + parsedMs);

    return parsedMs;
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item;
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList, itemName) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) {
                logger("Study event: " + studyeventList[i])
                var startTime = pullItemFromForm(temp, itemName);
                if (startTime || startTime.value != null) return startTime;
            }
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