/* jshint strict: false */

var studyevents = [
    "Day 1"
];

var formName = [
    "🟡IP_Exposure"
];

var itemName = [
    "End Date/Time of Infusion"
];

var currentStudyName = formJson.form.studyEventName;
var item = itemJson.item;

try {
    logger("Study Event: " + currentStudyName);

    if (currentStudyName != "Day 2, EOI" &&
        currentStudyName != "Day 2, 12Hr Post EOI") {
        return true;
    }

    var form = pullForm(studyevents, formName);

    if (!form) {
        logger("Dosing form not found.");
        return null;
    }

    var eoiItem = pullItemFromForm(form, itemName);

    if (!eoiItem || !item || item.value == null || item.value === "") {
        logger("End of Infusion or collected date/time is missing.");
        return true;
    }


    var eoiMs = Number(eoiItem.dateValueMs);
    var collectedMs = Number(item.dateValueMs);

    if (isNaN(eoiMs) || isNaN(collectedMs)) {
        logger("Unable to determine dateValueMs.");
        return true;
    }

    var eoiMin = Math.floor(eoiMs / 60000);
    var collectedMin = Math.floor(collectedMs / 60000);
    var diffMinutes = collectedMin - eoiMin;
    
    logger("End of Infusion: " + formatDateTime(eoiItem.value));
    logger("Collected Date/Time: " + formatDateTime(item.value));
    logger("EOI floored minute: " + eoiMin);
    logger("Collected floored minute: " + collectedMin);
    logger("Difference in minutes: " + diffMinutes);

    if (currentStudyName == "Day 2, EOI") {
        logger("Allowed window: -20 to +20 minutes from EOI.");

        return diffMinutes >= 0 && diffMinutes <= 20;
    }

    if (currentStudyName == "Day 2, 12Hr Post EOI") {
        logger("Allowed window: 710 to 730 minutes after EOI.");

        return diffMinutes >= 710 && diffMinutes <= 730;
    }

    return true;

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function formatDateTime(dateTime) {
    if (!dateTime) return "";

    var parts = dateTime.toString().split("T");
    if (parts.length !== 2) return dateTime;

    var dateParts = parts[0].split("-");
    if (dateParts.length !== 3) return dateTime;

    var months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];

    var year = dateParts[0];
    var month = months[Number(dateParts[1]) - 1];
    var day = dateParts[2];
    var time = parts[1].substring(0, 8);

    return day + " " + month + " " + year + " T " + time;
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

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) return [];

    var keepers = [];

    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];

        if (!formData ||
            !formData.form ||
            !formData.form.itemGroups ||
            formData.form.itemGroups.length < 1) {
            continue;
        }

        if (formData.form.canceled == false &&
            formData.form.itemGroups[0].canceled == false &&
            (
                formData.form.dataCollectionStatus == "Complete" ||
                (INCLUDE_NONCONFORMANT_DATA == true &&
                    formData.form.dataCollectionStatus == "Nonconformant") ||
                formData.form.dataCollectionStatus == "Incomplete"
            )) {
            keepers.push(formData);
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
    if (!form || !form.form) return null;

    var itemGroups = form.form.itemGroups;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];

        if (!group || group.canceled) continue;
        if (!group.items || group.items.length < 1) continue;

        for (var j = 0; j < group.items.length; j++) {
            var pulledItem = group.items[j];

            if (pulledItem &&
                containsItemName(targetItem, pulledItem.name) &&
                pulledItem.value !== null &&
                pulledItem.value !== "" &&
                !pulledItem.canceled) {
                return pulledItem;
            }
        }
    }

    return null;
}

