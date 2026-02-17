const item = itemJson.item;
const form = formJson.form;
const studyevent = formJson.form.studyEventName;
var ecgFormNames = [
  "⚡DAY-1 ECG SINGLE 12 LEAD V1",
  "⚡ECG SINGLE 12 LEAD"
];
const confirmationItem = [
    "Repeat WAS THE VITALS PERFORMED?",
    "VS_supine time check",
];
// If confirmation is YES
const vs_same_as_ecg = [
    "supine_time_start",
    "START SUPINE",
    "VS_Resting time (same as ECG)",
    "vs_same_as_ecg"
];

// If confirmation is NO or doesn't exist
const supinetimeStart = [
    "REPEAT START VITALS SUPINE TIME.",
    "vs_same_as_ecg",
    "vs_same_as_ecg.",
    "VS_Resting time",
    "supine_time_start",
    "supine_time_start.",
];

const difference = 5;

const confirmation = pullItemFromForm(formJson, confirmationItem);

var supineTime = null;
var isYes = false;

if (confirmation && confirmation.value != null) {
    var normalized = String(confirmation.value).trim().toLowerCase();
    isYes = normalized.indexOf("y") === 0;
}

if (isYes) {
    supineTime = pullItemFromForm(formJson, vs_same_as_ecg);
    if (!supineTime || supineTime.value === null) {
        var ecgform = pullForm([studyevent], ecgFormNames)
        supineTime = pullItemFromForm(ecgform, vs_same_as_ecg);
    }
} else {
    supineTime = pullItemFromForm(formJson, supinetimeStart);
}

if (!supineTime || supineTime.value == null) return null;

var supineTimeMs = supineTime.dateValueMs;
const collectedTimeMs = item.dateValueMs;

logger("supine Time :"  + formatDateTimeByType(supineTime));
logger("collected time: " + formatDateTimeByType(item));

const differenceMs = collectedTimeMs - supineTimeMs;
logger(differenceMs)
if (differenceMs < 0) {
    customErrorMessage("Selected time is less than previous supine time. Previous Supine Time: " + formatDateTimeByType(supineTime))
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger(differenceInMins)
if(differenceInMins >= difference ){
    return true;
}

customErrorMessage("Difference is out of range: " + Math.abs(differenceInMins) + ", Previous Supine Time: " + formatDateTimeByType(supineTime));
return false;


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


function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

  if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item;
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