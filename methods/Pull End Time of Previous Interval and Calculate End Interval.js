var formMapping = {
    "(*) LAB_URINE PK (0-4 Hours)": {
        previousForm: "🟡IP_Dose ABBV-295 or Placebo flex dosing (Day 134)",
        studyEvents: ["Period 1 Day 85", "Period 1 Day 134"]
    },
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

// First interval: check these dosing datetime items in order.
var dosingItemNames = [
    "EC_Dose Date/Time SubQ 3mg",
    "EC_Dose Date/Time SubQ 6mg",
    "EC_Dose Date/Time SubQ 9mg",
    "EC_Dose Date/Time SubQ 12mg",
    "EC_Dose Date/Time SubQ 15mg",
    "EC_Dose Date/Time SubQ 18mg",
    "EC_Dose Date/Time SubQ 22mg"
];

// Subsequent intervals: manual time takes priority.
// The expected interval-end item is tried first among the fallbacks.
var manualIntervalEndItem = "Manual Urine Interval End Time";

var intervalEndItems = [
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
logger("Current form: " + formName);

var isV2 = / v2$/.test(formName);
var baseFormName = isV2
    ? formName.replace(/ v2$/, "")
    : formName;

var mapping = formMapping[baseFormName];

if (!mapping) {
    logger("ERROR: No mapping found for " + formName);
    return null;
}

var previousBaseName = mapping.previousForm;
var studyEventList = mapping.studyEvents;
var currentStudyEvent = formJson.form.studyEventName;

logger("Previous base name: " + previousBaseName);
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

    if (dayNumber >= expectedScheduleStart &&
        dayNumber <= expectedScheduleEnd) {
        filteredStudyEvents.push(studyEventList[i]);
    }
}

if (filteredStudyEvents.length === 0) {
    logger("ERROR: No Study Events match current schedule.");
    return null;
}

logger("Selected schedule: " +
    (useAlternateSchedule ? "Day 134-139" : "Day 85-89"));
logger("Study Events to search: " + filteredStudyEvents.join(", "));

var currentInterval = getInterval(baseFormName);

if (!currentInterval) {
    logger("ERROR: Could not determine current interval endpoints.");
    return null;
}

var isDosingForm = containsValue(previousBaseName, "dosing");
var previousInterval = isDosingForm
    ? null
    : getInterval(previousBaseName);

if (!isDosingForm && !previousInterval) {
    logger("ERROR: Could not determine previous interval endpoints.");
    return null;
}

var hoursToAdd;
var fallbackIntervalItem = null;

if (isDosingForm) {
    hoursToAdd = currentInterval.end;
} else {
    hoursToAdd = currentInterval.end - previousInterval.end;

    fallbackIntervalItem =
        previousInterval.start + " to " +
        previousInterval.end +
        " hr Interval End (both)";
}

if (hoursToAdd <= 0) {
    logger("ERROR: Invalid interval difference: " + hoursToAdd);
    return null;
}

logger("Hours to add: " + hoursToAdd);

var previousForms;

if (isDosingForm) {
    previousForms = [previousBaseName];
} else {
    previousForms = isV2
        ? [previousBaseName + " v2", previousBaseName]
        : [previousBaseName, previousBaseName + " v2"];
}

var previousValue = null;
var selectedSource = null;
var selectedStudyEvent = null;
var selectedItem = null;

// Search forms using the existing schedule and V2 priority.
for (var i = 0; i < previousForms.length; i++) {
    for (var j = 0; j < filteredStudyEvents.length; j++) {
        var sourceForm = pullForm(
            [filteredStudyEvents[j]],
            [previousForms[i]]
        );

        if (!sourceForm) {
            logger(
                "Source unavailable: " +
                filteredStudyEvents[j] + " / " + previousForms[i]
            );
            continue;
        }

        var value = null;
        var matchedItem = null;

        if (isDosingForm) {
            // First interval: find the first available dosing datetime.
            var dosingResult = pullFirstAvailableItem(
                sourceForm,
                dosingItemNames
            );

            value = dosingResult ? dosingResult.value : null;
            matchedItem = dosingResult ? dosingResult.name : null;

        } else {
            // Later intervals: manual time ALWAYS has first priority.
            value = pullItemFromForm(
                sourceForm,
                manualIntervalEndItem
            );

            if (isUsableValue(value)) {
                matchedItem = manualIntervalEndItem;
            } else {
                // Try the expected interval first, then other interval ends.
                var orderedFallbacks = [fallbackIntervalItem];

                for (var k = 0; k < intervalEndItems.length; k++) {
                    if (normalizeItemName(intervalEndItems[k]) !==
                        normalizeItemName(fallbackIntervalItem)) {
                        orderedFallbacks.push(intervalEndItems[k]);
                    }
                }

                value = null;

                for (var k = 0; k < orderedFallbacks.length; k++) {
                    value = pullItemFromForm(
                        sourceForm,
                        orderedFallbacks[k]
                    );

                    if (isUsableValue(value)) {
                        matchedItem = orderedFallbacks[k];
                        break;
                    }
                }
            }
        }

        if (!isUsableValue(value)) {
            logger(
                "No usable datetime found in: " +
                filteredStudyEvents[j] + " / " + previousForms[i]
            );
            continue;
        }

        previousValue = value;
        selectedItem = matchedItem;
        selectedSource = previousForms[i];
        selectedStudyEvent = filteredStudyEvents[j];

        break;
    }

    if (isUsableValue(previousValue)) break;
}

if (!isUsableValue(previousValue)) {
    logger("ERROR: No usable previous datetime value found.");
    return null;
}

logger("Selected Study Event: " + selectedStudyEvent);
logger("Selected source form: " + selectedSource);
logger("Selected source item: " + selectedItem);
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

function isUsableValue(value) {
    return value !== null &&
        value !== undefined &&
        value !== "";
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

            if (!item || !item.name || item.canceled) continue;
            if (normalizeItemName(item.name) !== target) continue;

            if (isUsableValue(item.value)) {
                return item.value;
            }
        }
    }

    return null;
}

function pullFirstAvailableItem(form, itemNames) {
    for (var i = 0; i < itemNames.length; i++) {
        var value = pullItemFromForm(form, itemNames[i]);

        if (isUsableValue(value)) {
            return {
                name: itemNames[i],
                value: value
            };
        }
    }

    return null;
}

function pullForm(studyEventList, formNameList) {
    for (var i = 0; i < studyEventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(
                studyEventList[i],
                formNameList[j]
            );

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

        if (status === "Complete" ||
            (includeNonconformantData && status === "Nonconformant") ||
            status === "Incomplete") {
            keepers.push(formData);
        }
    }

    return keepers;
}

function addHoursToDateTime(value, hours) {
    if (typeof value !== "string") return null;

    var year, month, day, hour, minute, second;
    var match = value.match(
        /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/
    );

    if (match) {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
        day = parseInt(match[3], 10);
        hour = parseInt(match[4], 10);
        minute = parseInt(match[5], 10);
        second = parseInt(match[6], 10);
    } else {
        match = value.match(
            /^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/
        );

        if (!match) return null;

        var monthNames = {
            jan: 1, feb: 2, mar: 3, apr: 4,
            may: 5, jun: 6, jul: 7, aug: 8,
            sep: 9, oct: 10, nov: 11, dec: 12
        };

        day = parseInt(match[1], 10);
        month = monthNames[match[2].toLowerCase()];
        year = parseInt(match[3], 10);
        hour = parseInt(match[4], 10);
        minute = parseInt(match[5], 10);
        second = parseInt(match[6], 10);

        if (!month) return null;
    }

    if (month < 1 || month > 12 ||
        day < 1 || day > 31 ||
        hour > 23 || minute > 59 || second > 59 ||
        hour < 0 || minute < 0 || second < 0) {
        return null;
    }

    var date = new Date(Date.UTC(
        year, month - 1, day, hour, minute, second
    ));

    if (date.getUTCFullYear() !== year ||
        date.getUTCMonth() !== month - 1 ||
        date.getUTCDate() !== day) {
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

function containsValue(input, keyword) {
    if (input == null || keyword == null) return false;

    return input.toString().toLowerCase().indexOf(
        keyword.toString().toLowerCase()
    ) !== -1;
}