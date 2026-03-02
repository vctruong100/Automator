const item = itemJson.item;
const studyevent = formJson.form.studyEventName;
const form = formJson.form;

const doseForms = [
    "IP_Study Drug Exposure - Blinded Study Medication Part A",
    "IP_Study Drug Exposure - Blinded Study Medication Part B",
]

const doseStudyEvent = [
    "Day 1",
]

const doseItem = [
    "Datetime of Administration",    
]

var upperInterval = getUpperIntervalValue(studyevent);
if (upperInterval) {
    upperInterval = parseInt(upperInterval);
}
logger("Upper interval: " + upperInterval)
var doseForm = pullForm(doseStudyEvent, doseForms);
if (!doseForm) return null;
var doseTime = pullItemFromForm(doseForm, doseItem)
var doseTimeMs = parseDateTime(doseTime);

logger("Dose Time: "  + formatDateTimeByType(doseTime));

var intervalMs = upperInterval * 3600000;
var endTimeMs = doseTimeMs + intervalMs;

var endDate = new Date(endTimeMs);
var formatted = formatDateObject(endDate);
var converted = convertToOriginalFormat(formatted)
logger("End Time: " + formatDateObject(endDate));
logger("Converted End Time: " + converted);

return converted;

function convertToOriginalFormat(displayString) {
    if (!displayString) return "";

    var months = {
        "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
        "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
        "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
    };

    var parts = displayString.split(" ");
    if (parts.length < 4) return "";

    var day = parts[0];
    var monthName = parts[1];
    var year = parts[2];
    var time = parts[3];

    var month = months[monthName];
    if (!month) return "";

    day = ("0" + parseInt(day, 10)).slice(-2);

    return year + "-" + month + "-" + day + "T" + time;
}

function parseDateTime(doseTime) {
    var splitDT = doseTime.value.split("T");
    var datePart = splitDT[0];
    var timePart = splitDT.length > 1 ? splitDT[1] : "00:00:00";

    var datePieces = datePart.split("-");
    var timePieces = timePart.split(":");

    var year = parseInt(datePieces[0], 10);
    var month = parseInt(datePieces[1], 10) - 1;
    var day = parseInt(datePieces[2], 10);

    var hour = parseInt(timePieces[0], 10);
    var minute = parseInt(timePieces[1], 10);
    var second = parseInt(timePieces[2], 10);

    var doseDate = new Date(year, month, day, hour, minute, second);
    var doseTimeMs = doseDate.getTime();    
    return doseTimeMs;
}
function getUpperIntervalValue(studyEventName) {
    if (!studyEventName) return null;

    var parts = studyEventName.split(" ");

    if (parts.length >= 3 && parts[1].toLowerCase() === "to") {
        var value = parseInt(parts[2], 10);
        if (!isNaN(value)) {
            return value;
        }
    }

    return null;
}

function formatDateObject(dateObj) {
    if (!dateObj) return "";

    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    var day = ("0" + dateObj.getDate()).slice(-2);
    var month = months[dateObj.getMonth()];
    var year = dateObj.getFullYear();

    var hour = ("0" + dateObj.getHours()).slice(-2);
    var minute = ("0" + dateObj.getMinutes()).slice(-2);
    var second = ("0" + dateObj.getSeconds()).slice(-2);

    return day + " " + month + " " + year + " "
           + hour + ":" + minute + ":" + second;
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


function formatDateTimeByType(item) {
    if (!item || !item.value) return "";

    var value = item.value;
    var type = (item.dataType || "").toLowerCase();

    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    var year = "";
    var month = "";
    var day = "";
    var hour = "00";
    var minute = "00";
    var second = "00";

    // DATETIME TYPES
    if (type.indexOf("datetime") !== -1) {

        var splitDT = value.split("T");
        var datePart = splitDT[0];
        var timePart = splitDT.length > 1 ? splitDT[1] : "";

        if (datePart) {
            var datePieces = datePart.split("-");
            if (datePieces.length >= 3) {
                year = datePieces[0];
                month = parseInt(datePieces[1], 10) - 1;
                day = datePieces[2];
            }
        }

        if (timePart) {
            var timePieces = timePart.split(":");
            if (timePieces.length >= 1) hour = timePieces[0];
            if (timePieces.length >= 2) minute = timePieces[1];
            if (timePieces.length >= 3) second = timePieces[2];
        }

        if (!year || month === "" || !day) return value;

        return day + " " + months[month] + " " + year + " "
               + hour + ":" + minute + ":" + second;
    }

    // DATE TYPES
    if (type.indexOf("date") !== -1) {

        var dateOnly = value.split("-");
        if (dateOnly.length < 3) return value;

        year = dateOnly[0];
        month = parseInt(dateOnly[1], 10) - 1;
        day = dateOnly[2];

        return day + " " + months[month] + " " + year;
    }

    // TIME TYPES
    if (type.indexOf("time") !== -1) {

        var timeOnly = value.split(":");
        if (timeOnly.length >= 1) hour = timeOnly[0];
        if (timeOnly.length >= 2) minute = timeOnly[1];
        if (timeOnly.length >= 3) second = timeOnly[2];

        return hour + ":" + minute + ":" + second;
    }

    return value;
}