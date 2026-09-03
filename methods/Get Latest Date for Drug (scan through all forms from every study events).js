/* jshint strict: false */

var currentDate = itemJson.item.value;
if (currentDate && currentDate !== null) {
    logger('Current date is already set: ' + currentDate + '. Skipping calculation.');
    return currentDate; 
} else {
    logger('No date set for the current item. Proceeding with calculation.');
}

var formNames = [
    "CM_In-house administration of asthma medication"    
];

var categoryItemName = [
    "Standard of care Asthma Therapy Category",
]

var dateTimeItemNames = [
    "Date and time of administration (in-house dosing)",
]

// Step 1: Retrieve all forms across study events
var formData = [];

for (var i = 0; i < formNames.length; i++) {
    formData = formData.concat(findFormDataAcrossStudyEvents(formNames[i], false));
}

// Execution
try {
    var dateValues = getDateValues(formData);
    logger('All num values: ' + dateValues.join(', '));

    var latestDate = getLatestDateValue(dateValues);
    logger('Latest date value: ' + latestDate);

    return latestDate;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
// Step 2: Extract Date values
function getDateValues(formData) {
    var dateValues = [];
    logger('Form data length: ' + formData.length);

    for (var i = 0; i < formData.length; i++) {
        var form = formData[i].form;
        if (form.canceled) continue;
        
        logger('Inspecting form: ' + form.name);

        if (form.itemGroups && form.itemGroups.length > 0) {
            for (var j = 0; j < form.itemGroups.length; j++) {
                var itemGroup = form.itemGroups[j];
                var isLABA = false;
                
                if (itemGroup.canceled) continue;
                
                if (itemGroup.items && itemGroup.items.length > 0) {
                    for (var k = 0; k < itemGroup.items.length; k++) {
                        var item = itemGroup.items[k];
                        if (item && containsItemName(categoryItemName, item.name) && item.value == "ICS") isLABA = true;
                        if (!isLABA) continue;
                        if (item && containsItemName(dateTimeItemNames, item.name)) {
                            var value = item.value;
                            if (value !== null) {
                                dateValues.push(value);
                                logger('Date value added: ' + value);
                            }
                        }
                    }
                }
            }
        } else {
            logger('No item groups found for form: ' + form.name);
        }
    }

    logger('Collected num values: ' + dateValues.join(', '));
    return dateValues;
}

function normalizeDateValue(value) {
    if (!value) {
        return null;
    }

    if (value.indexOf("T") === -1) {
        return value + "T00:00:00";
    }

    if (value.indexOf("T  :  :") !== -1) {
        return value.split("T")[0] + "T00:00:00";
    }

    return value;
}

function getLatestDateValue(dateValues) {
    if (dateValues.length === 0) {
        return null;
    }

    var latestOriginal = dateValues[0];
    var latestNormalized = normalizeDateValue(latestOriginal);

    for (var i = 1; i < dateValues.length; i++) {
        var currentOriginal = dateValues[i];
        var currentNormalized = normalizeDateValue(currentOriginal);

        if (currentNormalized > latestNormalized) {
            latestNormalized = currentNormalized;
            latestOriginal = currentOriginal;
        }
    }

    return latestOriginal;
}

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