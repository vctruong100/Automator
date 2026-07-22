/* jshint strict: false */

// Version: v2
// Purpose: Validates visit dates fall within protocol-defined windows. Uses current study event to calculate previous and time window.

var studyevents = [
    "Day 1"
]
var formName = [
    "🟡IP_DRUG ADMINISTRATION - INJECTION","🟡IP_DRUG ADMINISTRATION - INJECTION_V2",
]
var itemName = [
    "IP_StartDate", "Datetime of Administration"
]

var currentStudyName = formJson.form.studyEventName;
var item = itemJson.item;

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
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (!formData || !formData.form || !formData.form.itemGroups || formData.form.itemGroups.length < 1) continue;

        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
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
function parseDate(dateStr) {
    if (!dateStr) return null;
    var clean = dateStr.toString().split('T')[0];
    var parts = clean.split('-');
    if (parts.length !== 3) return null;
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var day = Number(parts[2]);
    if (!year || !month || !day) return null;
    var parsed = new Date(year, month - 1, day);
    if (parsed.getFullYear() !== year || parsed.getMonth() !== month - 1 || parsed.getDate() !== day) return null;
    return parsed;
}

function addCalendarDays(dateObj, days) {
    var copy = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    copy.setDate(copy.getDate() + days);
    return copy;
}

function calendarDayNumber(dateObj) {
    return Math.floor(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()) / 86400000);
}

function formatDateText(dateObj) {
    var MMM = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    var d = dateObj.getDate();
    var m = MMM[dateObj.getMonth()];
    var y = dateObj.getFullYear();
    var hy = "\u2011";
    return (d < 10 ? "0"+d : ""+d) + hy + m + hy + y;
}

function pullItemFromForm(form, targetItem) {
    if (!form || !form.form) return null;
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        if (!group.items || group.items.length < 1) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") return item;
        }
    }
    return null;
}

logger("Study event: " + currentStudyName);
if (currentStudyName == "Screening") return true;
var dayMatch = currentStudyName ? currentStudyName.toString().match(/^Day\s+(\d+)\b/i) : null;
if (!dayMatch) return true;

var day = parseInt(dayMatch[1], 10);
logger("Day: " + day);

if (day < 8) return true;

var form = pullForm(studyevents, formName);

if (!form) return null;
var val = pullItemFromForm(form, itemName);

if (!val || !item || !item.value) return true;
logger("Start Date: " + val.value);
logger("Collected Date: " + item.value);
var baseDate = parseDate(val.value);
var collectedDate = parseDate(item.value);
if (!baseDate || !collectedDate) return true;

var addDays = day - 1;
logger("Add days: " + addDays)
var allowedRange = day <= 141 ? 2 : 7;

logger("Allowed range: " + allowedRange)

var targetDate = addCalendarDays(baseDate, addDays);

var targetMs = baseDate.getTime() + addDays * 86400000;

var collectedFromMs = new Date(Number(item.dateValueMs));
var targetFromMs = new Date(Number(targetMs));

logger("Collected raw value: " + item.value);
logger("Collected dateValueMs: " + item.dateValueMs);
logger("Collected from dateValueMs local: " + collectedFromMs.toString());
logger("Collected from dateValueMs ISO: " + collectedFromMs.toISOString());

logger("Target date local: " + targetFromMs.toString());
logger("Target date ISO: " + targetFromMs.toISOString());

logger("Raw ms difference: " + Math.abs(item.dateValueMs - targetMs));
logger("Raw day difference: " + (Math.abs(item.dateValueMs - targetMs) / 86400000));

logger("Floored Collected Date: " + calendarDayNumber(collectedDate))
logger("Floored Target Date: " + calendarDayNumber(targetDate))
var diffDays = Math.abs(calendarDayNumber(collectedDate) - calendarDayNumber(targetDate));
logger("Differences: " + diffDays);
return diffDays <= allowedRange;
