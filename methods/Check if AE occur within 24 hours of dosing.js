var studyevents = [
    "Wk12 Day 84",
    "Day -1",
]
var doseForm = [
    "IP_Acetaminophen Administration"   
]

var doseItem = [
    "Start Date/Time (acetaminophen)"    
]

var collectedTimeItem = [
    "AE_Onset_Date/Time"
]

const difference = 24; // in hours
var methodType = "B";

// ======== Don't modify ========
try {
    var endTime = pullItemFromForm(formJson, collectedTimeItem);
    if (!endTime || endTime.value == null || endTime.value === "") return null;

    logger("Collected Time: " + endTime.value);

    var doseResult = pullForm(studyevents, doseForm, doseItem, endTime.dateValueMs);
    if (!doseResult || !doseResult.item) {
        logger("No dose found before collected time");
        return "N";
    }

    var startTime = doseResult.item;
    var differenceInMins = doseResult.differenceInMins;

    logger("Study event: " + doseResult.studyEvent);
    logger("Start time: " + startTime.value);

    if (methodType === "A") {
        logger("Method A used");
    } else {
        logger("Method B used");
        logger("Start (hrs): " + (doseResult.startMin / 60));
        logger("End (hrs): " + (doseResult.endMin / 60));
    }

    logger("Diff (hrs): " + (differenceInMins / 60));

    if (differenceInMins >= (difference * 60)) {
        return "N";
    }

    return "Y";
} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return "N";
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

function pullForm(studyeventList, formNameList, itemName, endTimeMs) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);

            if (temp) {
                var startTime = pullItemFromForm(temp, itemName);
                if (!startTime || startTime.value == null || startTime.value === "" || !startTime.dateValueMs) continue;

                var differenceInMins;
                var startMin;
                var endMin;

                if (methodType === "A") {
                    differenceInMins = Math.floor((endTimeMs - startTime.dateValueMs) / (1000 * 60));
                } else {
                    startMin = Math.floor(startTime.dateValueMs / (1000 * 60));
                    endMin = Math.floor(endTimeMs / (1000 * 60));
                    differenceInMins = endMin - startMin;
                }

                logger("Candidate study event: " + studyeventList[i]);
                logger("Candidate start time: " + startTime.value);
                logger("Candidate diff (hrs): " + (differenceInMins / 60));

                if (differenceInMins >= 0) {
                    return {
                        item: startTime,
                        studyEvent: studyeventList[i],
                        differenceInMins: differenceInMins,
                        startMin: startMin,
                        endMin: endMin
                    };
                }

                logger("Skipping candidate because collected time is before dose time");
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