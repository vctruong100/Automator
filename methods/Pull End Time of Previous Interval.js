// Used in AbbVie M26-089

var baseFormMapping = {
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

var itemNames = [
    "0 to 4 hr Interval End (both)",
    "4 to 8 hr Interval End (both)",
    "8 to 12 hr Interval End (both)",
    "12 to 24 hr Interval End (both)",
    "24 to 48 hr Interval End (both)",
    "48 to 96 hr Interval End (both)",
    "96 to 120 hr Interval End (both)",
    "120 to 144 hr Interval End (both)"
];

var formName = formJson.form.name;
var currentStudyEvent = formJson.form.studyEventName;

logger("Current form: " + formName);
logger("Current Study Event: " + currentStudyEvent);

// Determine whether the current form is v1 or v2.
var isV2 = / v2$/.test(formName);
var baseFormName = isV2 ? formName.replace(/ v2$/, "") : formName;
var mapping = baseFormMapping[baseFormName];

if (!mapping) {
    logger("ERROR: No mapping found for " + formName);
    return null;
}

if (!currentStudyEvent) {
    logger("ERROR: Cannot determine the current form's Study Event.");
    return null;
}

// Determine which schedule the current form belongs to.
var useAlternateSchedule = /Day\s+13[4-8]\b/i.test(currentStudyEvent);
var useOriginalSchedule = /Day\s+8[5-9]\b/i.test(currentStudyEvent);

if (!useAlternateSchedule && !useOriginalSchedule) {
    logger("ERROR: Current Study Event is outside the supported schedules: " + currentStudyEvent);
    return null;
}

var expectedScheduleStart = useAlternateSchedule ? 134 : 85;
var expectedScheduleEnd = useAlternateSchedule ? 138 : 89;
var filteredStudyEvents = [];

for (var i = 0; i < mapping.studyEvents.length; i++) {
    var dayMatch = mapping.studyEvents[i].match(/Day\s+(\d+)/i);
    if (!dayMatch) continue;

    var dayNumber = parseInt(dayMatch[1], 10);

    if (dayNumber >= expectedScheduleStart && dayNumber <= expectedScheduleEnd) {
        filteredStudyEvents.push(mapping.studyEvents[i]);
    }
}

if (filteredStudyEvents.length === 0) {
    logger("ERROR: No mapped Study Events match the current schedule: " + currentStudyEvent);
    return null;
}

logger("Selected schedule: " + (useAlternateSchedule ? "Day 134-138" : "Day 85-89"));
logger("Study Events to search: " + filteredStudyEvents.join(", "));

// Determine source form version priority.
var previousBaseName = mapping.previousForm;
var previousForms = isV2 ? [previousBaseName + " v2", previousBaseName] : [previousBaseName, previousBaseName + " v2"];

logger("Source form priority: " + previousForms.join(" -> "));

var result = null;
var selectedSource = null;
var selectedStudyEvent = null;

for (var i = 0; i < previousForms.length; i++) {
    for (var j = 0; j < filteredStudyEvents.length; j++) {
        var sourceForm = pullForm([filteredStudyEvents[j]], [previousForms[i]]);

        if (!sourceForm) {
            logger("Source form not found or unusable: " + filteredStudyEvents[j] + " / " + previousForms[i]);
            continue;
        }

        var value = pullItemFromForm(sourceForm, itemNames);

        if (value !== null && value !== undefined && value !== "") {
            result = value;
            selectedSource = previousForms[i];
            selectedStudyEvent = filteredStudyEvents[j];
            break;
        }

        logger("No usable target item in: " + filteredStudyEvents[j] + " / " + previousForms[i]);
    }

    if (result !== null) break;
}

if (result === null) {
    logger("ERROR: No usable interval-end value found.");
    return null;
}

logger("Selected Study Event: " + selectedStudyEvent);
logger("Selected source form: " + selectedSource);
logger("Returned interval-end value: " + result);

return result;


// Helper functions

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, targetName) {
    var normalizedName = normalizeItemName(targetName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) return true;
    }

    return false;
}

function pullItemFromForm(form, targetItems) {
    if (!form || !form.form) return null;

    var itemGroups = form.form.itemGroups;
    if (!itemGroups || itemGroups.length < 1) return null;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        for (var j = 0; j < group.items.length; j++) {
            var item = group.items[j];

            if (!item || !item.name) continue;

            if (containsItemName(targetItems, item.name) && item.value !== null && item.value !== undefined && item.value !== "") {
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
    var completedForms = collectCompleted(arrayForms, true);

    if (!completedForms || completedForms.length === 0) return null;

    return completedForms[0];
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