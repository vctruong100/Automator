const item = itemJson.item;
const form = formJson.form;
const confirmationItem = [
    "Scn_remained_Semi_recumbent_ECG",
];
// If confirmation is YES
const vs_same_as_ecg = [
    "START SUPINE",
    "vs_same_as_ecg"
];

// If confirmation is NO or doesn't exist
const supinetimeStart = [
    "START SUPINE",
    "REPEAT_START_INITIAL_TIME",
    "ECG_Supine resting time",
    "supine_time_start",
    "vs_same_as_ecg",
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
} else {
    supineTime = pullItemFromForm(formJson, supinetimeStart);
}


if (!supineTime || supineTime == null) return null;

var supineTimeMs = supineTime.dateValueMs;
const collectedTimeMs = item.dateValueMs;

logger("Data type: " + itemJson.item.dataType);
logger("supine Time :"  + formatDateTimeByType(supineTime));
logger("collected time: " + formatDateTimeByType(item));

const differenceMs = collectedTimeMs - supineTimeMs;
if (differenceMs < 0) {
    customErrorMessage("Selected time is less than previous supine time. Previous Supine Time: " + formatDateTimeByType(supineTime))
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger("Difference in minutes: " + differenceInMins)

if (differenceInMins >= difference ) return true;

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