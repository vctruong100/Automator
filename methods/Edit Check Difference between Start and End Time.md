const startTimeItem = [
    "CSSRS_COLLECTION_DATETIME"
]

const difference = 5;

var startTime = pullItemFromForm(formJson, startTimeItem);
var endTime = itemJson.item;

if (!startTime || startTime.value == null || !endTime || endTime.value == null) return null;

var startTimeMs = startTime.dateValueMs;
var endTimeMs = endTime.dateValueMs;

logger("Start Time :"  + formatDateTimeByType(startTime));
logger("End Time: " + formatDateTimeByType(endTime));

const differenceMs = endTimeMs - startTimeMs;
logger(differenceMs)
if (differenceMs < 0) {
    customErrorMessage("Collected End Time is less than Start Time. Start Time: " + formatDateTimeByType(startTime))
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger(differenceInMins)
if(differenceInMins >= difference ){
    return true;
}

customErrorMessage("Difference is out of range: " + Math.abs(differenceInMins) + " minutes, Start Time: " + formatDateTimeByType(startTime));
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