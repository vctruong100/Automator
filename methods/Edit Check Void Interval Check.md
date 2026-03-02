const item = itemJson.item;
const studyevent = formJson.form.studyEventName;
const form = formJson.form;

const allForms = [
    "(*) 💧 -2 to 0 hr Urine Interval v3", // 0
    "(*)💧0 to 4 hr Urine Interval v4",// 1
    "(*) 💧 4 to 8 hr Urine Interval v2",// 2
    "(*) 💧 4 to 8 hr Urine Interval v3",// 2
    "(*) 💧 8 to 12 hrs Urine Interval v3",// 3
    "(*) 💧 12 to 24 hrs Urine Interval v3",// 4
    "(*) 💧24 to 48 hrs Urine Interval v3",// 5
    "(*) 💧 48 to 72 hrs Urine Interval v3",// 6
    "(*) 💧 72 to 96 hrs Urine Interval v3",// 7
]

const attachedItemName = [
    "PCVOID w/edit check",
    "PCVOID",
]
const formMaps = {
    "(*)💧0 to 4 hr Urine Interval v4" : "(*) 💧 -2 to 0 hr Urine Interval v3",
    "(*) 💧 4 to 8 hr Urine Interval v2" : "(*)💧0 to 4 hr Urine Interval v2",
    "(*) 💧 4 to 8 hr Urine Interval v3" : "(*)💧0 to 4 hr Urine Interval v4",
    "(*) 💧 8 to 12 hrs Urine Interval v3": "(*) 💧 4 to 8 hr Urine Interval v3",
    "(*) 💧 12 to 24 hrs Urine Interval v3": "(*) 💧 8 to 12 hrs Urine Interval v3",
    "(*) 💧24 to 48 hrs Urine Interval v3" : "(*) 💧 12 to 24 hrs Urine Interval v3",
    "(*) 💧 48 to 72 hrs Urine Interval v3": "(*) 💧24 to 48 hrs Urine Interval v3",
    "(*) 💧 72 to 96 hrs Urine Interval v3" : "(*) 💧 48 to 72 hrs Urine Interval v3",
}

const afterDoseItem = [
    "0 to 4 hrs END/4 to 8 hrs START", // 0
    "4 to 8 hrs END/8 to 12 hrs START", // 1
    "8 to 12 hrs END/12 to 24 hrs START", // 2
    "12 to 24 hrs END/24 to 48 hrs START",  // 3
    "24 to 48 hrs END/48 to 72 hrs START",// 4
    "48 to 72 hrs END/72 to 96 hrs START",// 5
    "72 to 96 hrs END", // 6
    
]

const studyEvents = [
    "-2 to 0 hr",
    "0 to 4 hrs",
    "4 to 8 hrs",
    "8 to 12 hrs",
    "12 to 24 hrs",
    "24 to 48 hrs",
    "48 to 72 hrs",
    "72 to 96 hrs",
]

const sameFormDifference = 15;
const prevFormDifference = 15;

var prevForm = null;
var prevEndTime = null;

var normalizedMaps = {};
for (var key in formMaps) {
    normalizedMaps[normalize(key)] = formMaps[key];
}

var endTime = pullItemFromForm(formJson, afterDoseItem)

if (form.name != allForms[1]) {
    logger("Form name: " + form.name);
    var prevFormName = normalizedMaps[normalize(form.name)];
    logger("Previous Form Name: " + prevFormName);
    prevForm = pullForm(studyEvents, [prevFormName]);
    if (!prevForm) {
        customErrorMessage("Could not detect " + prevFormName);
        return false;
    }
    prevEndTime = pullItemFromForm(prevForm, afterDoseItem);
}

var collectedTimeMs = null;
if (!item || item.dateValueMs == null || item.dateValueMs == undefined) {
    var voidTime = pullItemFromForm(formJson, attachedItemName);
    collectedTimeMs = voidTime.dateValueMs;
}
else {
    var voidTime = item;
    var collectedTimeMs = item.dateValueMs;
}
logger(voidTime.name)
var endTimeMs = endTime.dateValueMs;

logger("Collected Time: "  + formatDateTimeByType(voidTime));
logger("End Time: " + formatDateTimeByType(endTime));

const differenceMs =  collectedTimeMs - endTimeMs;
logger(differenceMs)
if (differenceMs < 0) {
    customErrorMessage("Difference between Void time and Interval time is less than 0.")
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

if(differenceInMins > sameFormDifference){
    customErrorMessage("Difference between Void time and Interval is more than 15 minutes.");
    return false;
}

if (form.name === allForms[1]) return true;

if (!prevEndTime) {
    customErrorMessage("Could not determine previous interval end time.");
    return false;
}
var prevEndTimeMs = prevEndTime.dateValueMs;
var differenceMs2 = collectedTimeMs - prevEndTimeMs;
logger(differenceMs2);
if (differenceMs2 < 0) {
    customErrorMessage("Difference between Void time and Interval time is less than 0.")
    return false;
}

const differenceInMins2 = Math.abs(Math.floor(differenceMs2 / (1000 * 60)))
if(differenceInMins2 < prevFormDifference){
    customErrorMessage("Void Time is within 15 minutes of previous interval '" + prevFormName + "'");
    return false;
}

return true;

function normalize(str) {
    return str.replace(/\s+/g, " ").trim();
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