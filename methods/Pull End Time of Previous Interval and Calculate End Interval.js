// used in AbbVie m26-089

var formMapping = {
    "(*) LAB_URINE PK (4-8 Hours)": {
        previousForm: "(*) LAB_URINE PK (0-4 Hours)",
        studyEvents: ["Period 1 Day 85", "Period 1 Day 134"]
    },
    "(*) LAB_URINE PK (8-12 Hours)": {
        previousForm: "(*) LAB_URINE PK (4-8 Hours)",
        studyEvents: ["Period 1 Day 85", "Period 1 Day 134"]
    },
    "(*) LAB_URINE PK (12-24 Hours)": {
        previousForm: "(*) LAB_URINE PK (8-12 Hours)",
        studyEvents: ["Period 1 Day 85", "Period 1 Day 134"]
    },
    "(*) LAB_URINE PK (24-48 Hours)": {
        previousForm: "(*) LAB_URINE PK (12-24 Hours)",
        studyEvents: ["Period 1 Day 85", "Period 1 Day 134"]
    },
    "(*) LAB_URINE PK (48-96 Hours)": {
        previousForm: "(*) LAB_URINE PK (24-48 Hours)",
        studyEvents: ["Period 1 Day 86", "Period 1 Day 135"]
    },
    "(*) LAB_URINE PK (96-120 Hours)": {
        previousForm: "(*) LAB_URINE PK (48-96 Hours)",
        studyEvents: ["Period 1 Day 87", "Period 1 Day 136"]
    },
    "(*) LAB_URINE PK (120-144 Hours)": {
        previousForm: "(*) LAB_URINE PK (96-120 Hours)",
        studyEvents: ["Period 1 Day 89", "Period 1 Day 138"]
    }
};

var formName = formJson.form.name;
logger("Current form: " + formName);

var isV2 = / v2$/.test(formName);
var baseFormName = isV2 ? formName.replace(/ v2$/, "") : formName;
var mapping = formMapping[baseFormName];

if (!mapping) {
    logger("ERROR: No mapping found for " + formName);
    return null;
}

var previousBaseName = mapping.previousForm;
var studyEventList = mapping.studyEvents;
var currentStudyEvent = formJson.form.studyEventName;

logger("Current Study Event: " + currentStudyEvent);

if (!currentStudyEvent) {
    logger("ERROR: Cannot determine current form's Study Event.");
    return null;
}

var useAlternateSchedule = /Day\s+13[4-9]\b/i.test(currentStudyEvent);
var expectedScheduleStart = useAlternateSchedule ? 134 : 85;
var expectedScheduleEnd = useAlternateSchedule ? 139 : 89;
var filteredStudyEvents = [];

for (var i = 0; i < studyEventList.length; i++) {
    var dayMatch = studyEventList[i].match(/Day\s+(\d+)/i);
    if (!dayMatch) continue;

    var dayNumber = parseInt(dayMatch[1], 10);

    if (dayNumber >= expectedScheduleStart && dayNumber <= expectedScheduleEnd) {
        filteredStudyEvents.push(studyEventList[i]);
    }
}

if (filteredStudyEvents.length === 0) {
    logger("ERROR: No Study Events match the current schedule: " + currentStudyEvent);
    return null;
}

logger("Selected schedule: " + (useAlternateSchedule ? "Day 134-139" : "Day 85-89"));
logger("Study Events to search: " + filteredStudyEvents.join(", "));

var currentInterval = getInterval(baseFormName);
var previousInterval = getInterval(previousBaseName);

if (!currentInterval || !previousInterval) {
    logger("ERROR: Could not determine interval endpoints.");
    return null;
}

var hoursToAdd = currentInterval.end - previousInterval.end;

if (hoursToAdd <= 0) {
    logger("ERROR: Invalid interval difference: " + hoursToAdd);
    return null;
}

var previousItemName = previousInterval.start + " to " + previousInterval.end + " hr Interval End (both)";
logger("Previous item: " + previousItemName);
logger("Hours to add: " + hoursToAdd);

var previousForms = isV2 ? [previousBaseName + " v2", previousBaseName] : [previousBaseName, previousBaseName + " v2"];
var previousValue = null;
var selectedSource = null;
var selectedStudyEvent = null;

for (var i = 0; i < previousForms.length; i++) {
    for (var j = 0; j < filteredStudyEvents.length; j++) {
        var sourceForm = pullForm([filteredStudyEvents[j]], [previousForms[i]]);

        if (!sourceForm) {
            logger("Source unavailable: " + filteredStudyEvents[j] + " / " + previousForms[i]);
            continue;
        }

        var value = pullItemFromForm(sourceForm, previousItemName);

        if (value === null || value === undefined || value === "") {
            logger("Target item unavailable: " + filteredStudyEvents[j] + " / " + previousForms[i]);
            continue;
        }

        previousValue = value;
        selectedSource = previousForms[i];
        selectedStudyEvent = filteredStudyEvents[j];
        break;
    }

    if (previousValue !== null) break;
}

if (previousValue === null) {
    logger("ERROR: No usable previous interval value found.");
    return null;
}

logger("Selected Study Event: " + selectedStudyEvent);
logger("Selected source form: " + selectedSource);
logger("Previous datetime: " + previousValue);

var result = addHoursToDateTime(previousValue, hoursToAdd);

if (result === null) {
    logger("ERROR: Invalid datetime: " + previousValue);
    return null;
}

logger("Calculated interval end: " + result);
return result;

function getInterval(name) {
    if (!name) return null;

    var match = name.match(/\((\d+)\s*-\s*(\d+)\s*Hours\)/i);
    if (!match) return null;

    return {
        start: parseInt(match[1], 10),
        end: parseInt(match[2], 10)
    };
}

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function pullItemFromForm(form, targetItemName) {
    if (!form || !form.form) return null;

    var itemGroups = form.form.itemGroups;
    if (!itemGroups || itemGroups.length === 0) return null;

    var target = normalizeItemName(targetItemName);

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        for (var j = 0; j < group.items.length; j++) {
            var item = group.items[j];
            if (!item || !item.name) continue;
            if (normalizeItemName(item.name) !== target) continue;

            if (item.value !== null && item.value !== undefined && item.value !== "") {
                return item.value;
            }
        }
    }

    return null;
}

function pullForm(studyEventList, formNameList) {
    for (var i = 0; i < studyEventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyEventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }

    return null;
}

function checkForm(studyEvent, formName) {
    var arrayForms = findFormData(studyEvent, formName);
    var availableForms = collectCompleted(arrayForms, true);

    if (!availableForms || availableForms.length === 0) return null;
    return availableForms[0];
}

function collectCompleted(formDataArray, includeNonconformantData) {
    if (!formDataArray) return [];

    var keepers = [];

    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (!formData || !formData.form) continue;

        var sourceForm = formData.form;
        var groups = sourceForm.itemGroups;

        if (sourceForm.canceled === true) continue;
        if (!groups || groups.length === 0) continue;
        if (!groups[0] || groups[0].canceled === true) continue;

        var status = sourceForm.dataCollectionStatus;

        if (status === "Complete" || (includeNonconformantData && status === "Nonconformant") || status === "Incomplete") {
            keepers.push(formData);
        }
    }

    return keepers;
}

function addHoursToDateTime(value, hours) {
    if (typeof value !== "string") return null;

    var match = value.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    if (!match) return null;

    var year = parseInt(match[1], 10);
    var month = parseInt(match[2], 10);
    var day = parseInt(match[3], 10);
    var hour = parseInt(match[4], 10);
    var minute = parseInt(match[5], 10);
    var second = parseInt(match[6], 10);

    if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59 || second > 59) {
        return null;
    }

    var date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));

    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
        return null;
    }

    date.setTime(date.getTime() + hours * 60 * 60 * 1000);

    function pad(number) {
        return number < 10 ? "0" + number : "" + number;
    }

    return date.getUTCFullYear() + "-" +
        pad(date.getUTCMonth() + 1) + "-" +
        pad(date.getUTCDate()) + "T" +
        pad(date.getUTCHours()) + ":" +
        pad(date.getUTCMinutes()) + ":" +
        pad(date.getUTCSeconds());
}