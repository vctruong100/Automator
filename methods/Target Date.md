const studyEvent = formJson.form.studyEventName;
const studyEventNames = [
    "Day 1 (pre)",
    "DAY 1 (Pre)",
    "V1 (Pre-Randomization)",
];
const formNames = [
    "RANDOMIZATION",
    "DISPOSITION RANDOMIZATION_V2.0",
    "DISPOSITION RANDOMIZATION",
    "RANDOMIZATION IMPALA (IRT)",
    "RANDOMIZATION IMPALA (IRT) V2.0",
];
const startDateItem = [
    "Date..",
    "Date of Randomization"
];

const map = {
    "V3 (D28 to D35)": 31,
    "V4 (D84 to D98)": 91,
    "V5 (D175 to D189)": 182,
    "V6 (Within 4 weeks)": 28
};

const range = {
    "V3 (D28 to D35)": 4,
    "V4 (D84 to D98)": 7,
    "V5 (D175 to D189)": 7,
    "V6 (Within 4 weeks)": 28
};

var form = pullForm(formNames, studyEventNames);
var day1Date = pullItemFromForm(form, startDateItem);
if (!day1Date || day1Date == null) return true;

var day1DateMs = isoToLocalMidnight(day1Date.value);
var day1Dateformat = msToDate(day1DateMs);

var addDays = map[studyEvent];
var allowedRange = range[studyEvent];
if (!addDays || addDays == null || addDays == -1) return true;

var targetMs = day1DateMs + addDays * 86400000;

var allowedRangeMs = allowedRange * 86400000;

var targetDateformat = msToDate(targetMs);
var minTargeDateformat = msToDate(targetMs - allowedRangeMs);
var maxTargetDateformat = msToDate(targetMs + allowedRangeMs);

log();

var rangeDateString = minTargeDateformat + " - " + maxTargetDateformat;

return rangeDateString;

function log() {
    logger("Study event: " + studyEvent);
    logger("Day 1 Date: " + day1Dateformat);
    logger("Days to add from Day 1: " + addDays);
    logger("Allowed range: " + allowedRange);
    logger("Target date: " + targetDateformat);
    logger("Target date range: " + minTargeDateformat + " to " + maxTargetDateformat);
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function collectCompleted(formDataArray) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (
          formData.form.canceled == false &&
          formData.form.itemGroups[0].canceled == false &&
          (
            formData.form.dataCollectionStatus == 'Complete' ||
            formData.form.dataCollectionStatus == 'Incomplete' ||
            formData.form.dataCollectionStatus == 'Nonconformant'
          )
        ) {
            keepers.push(formData);
        }
    }
    return keepers;
}

function checkForm(form, studyevent) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm[completedForm.length - 1];
    }
}

function msToDate(ms) {
    var d = new Date(Number(ms));

    var day = d.getDate();
    if (day < 10) day = "0" + day;

    var months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    var month = months[d.getMonth()];

    var year = d.getFullYear();

    return day + "" + month + "" + year;
}

function toLocalMidnight(ms) {
    var d = new Date(Number(ms));
    return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
    return null;
}


function isoToLocalMidnight(isoStr) {
    if (!isoStr) return null;
    var d = isoStr.split("T")[0]; 
    var p = d.split("-");
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getTime();
}