const studyevents = [
    "Day 1"
]
const formName = [
    "🟡IP_EVOLOCUMAB ADMINISTRATION",
]
const itemName = [
    "IP_StartDate"
]

const sampleItem = [
    "🟡 Evolocumab PK",
    "🟡 PCSK9 Serum",
    "🔴 Lipid Panel (LabCorp)",
]

var currentStudyName = formJson.form.studyEventName;
var item = itemJson.item;

try {

    logger("Study event: " + currentStudyName);

    if (currentStudyName == "Screening") {
        return true;
    }

    var form = pullForm(studyevents, formName);
    if (!form) return null;

    var doseItem = pullItemFromForm(form, itemName);

    if (!doseItem) {
        return null;
    }

    logger("Dosing time: " + doseItem.value);
    
    var sample = pullItemFromForm(formJson, sampleItem);
    if (!sample) return null;
    logger("Sample time: " + sample.value);
    var doseMs = doseItem.dateValueMs;
    var sampleMs = sample.dateValueMs;
    
    // var doseDate = parseDateOnly(doseItem.value);
    // var sampleDate = parseDateOnly(sample.value);
    
    // var differenceInDays = Math.floor((sampleDate.getTime() - doseDate.getTime()) / (1000 * 60 * 60 * 24));
    // logger("Dosing day: " + doseDate);
    // logger("Sample day: " + sampleDate);
    // logger("Difference In Days: " + differenceInDays);

    const minuteMs = 1000 * 60;
    const hourMs = minuteMs * 60;
    const dayMs = hourMs * 24;
    // var DaysInMS = differenceInDays * dayMs;
    
    var diffMs = sampleMs - doseMs + dayMs;
    logger("Difference in Ms: " + diffMs);
    logger("Difference in Days: " + msToDaysHours(diffMs))
    var lowerBound;
    var upperBound;

    // PREDOSE
    if (containsValue(currentStudyName, "predose")) {
        lowerBound = -30 * minuteMs;
        upperBound = 0;
        logger("Predose window");
    }

    // HOURLY VISITS (4h - 168h)
    var hourMatch = currentStudyName.match(/(\d+)\s*h/i);
    if (hourMatch) {
        var targetHour = parseInt(hourMatch[1]);
        if (targetHour >= 4 && targetHour <= 168) {
            var targetMs = targetHour * hourMs;
            lowerBound = targetMs - hourMs;
            upperBound = targetMs + hourMs;
            logger("Hourly window: " + targetHour + "h");
        }
    }

    // DAY VISITS
    var dayMatch = currentStudyName.match(/Day\s+(\d+)/i);

    if (dayMatch) {
        var targetDay = parseInt(dayMatch[1]);
        var targetMs = targetDay * dayMs;
        // Day 11 - 15 : ±24h 
        if (targetDay >= 11 && targetDay <= 15) {
            lowerBound = targetMs - dayMs;
            upperBound = targetMs + dayMs;
            logger("Day window ±24h");
        }

        // Day 22 - 64 : ±48h
        if (targetDay >= 22 && targetDay <= 64) {
            lowerBound = targetMs - (2 * dayMs);
            upperBound = targetMs + (2 * dayMs);
            logger("Day window ±48h");
        }
    }
    logger("Lowerbound: " + msToDaysHours(lowerBound));
    logger("upperbound: " + msToDaysHours(upperBound));
    if (diffMs >= lowerBound && diffMs <= upperBound) return "YES";
    return 'NO';
}
catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
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

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function msToDaysHours(milliseconds) {
    var totalMinutes = Math.floor(milliseconds / (1000 * 60));

    var days = Math.floor(totalMinutes / (60 * 24));
    var hours = Math.floor((totalMinutes % (60 * 24)) / 60);
    var minutes = totalMinutes % 60;

    return days + ":" + hours + ":" + minutes;
}

function parseDateOnly(value) {
    if (!value) return null;

    var datePart = value.split("T")[0];
    var parts = datePart.split("-");

    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1; 
    var day = parseInt(parts[2], 10);

    return new Date(year, month, day);
}