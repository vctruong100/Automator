/* jshint strict: false */

// Version: v1
// Purpose: Generates unique incremental IDs within repeated groups/forms.

var currentNum = parseInt(itemJson.item.value, 10);
if (!isNaN(currentNum)) {
    logger('Current num is already set: ' + currentNum + '. Skipping calculation.');
    return currentNum.toFixed(0);  // Return the current num if it's already set
} else {
    logger('No num set for the current item. Proceeding with calculation.');
}

var formNames = [
    "Inhalation Training - In-Check G16 (SCRN)",
    "Inhalation Training - In-Check G16",
];

// Step 1: Retrieve all forms across study events
var formData = [];

for (var i = 0; i < formNames.length; i++) {
    formData = formData.concat(findFormDataAcrossStudyEvents(formNames[i], false));
}

// Step 2: Extract num values
function getNumValues(formData) {
    var numValues = [];
    logger('Form data length: ' + formData.length);

    for (var i = 0; i < formData.length; i++) {
        var form = formData[i].form;
        logger('Inspecting form: ' + form.name);

        if (form.itemGroups && form.itemGroups.length > 0) {
            for (var j = 0; j < form.itemGroups.length; j++) {
                var itemGroup = form.itemGroups[j];

                if (itemGroup.items && itemGroup.items.length > 0) {
                    for (var k = 0; k < itemGroup.items.length; k++) {
                        var item = itemGroup.items[k];

                        if (item.name === "Inhalation Training Number") {
                            var value = parseInt(item.value, 10);
                            if (!isNaN(value)) {
                                numValues.push(value);
                                logger('Number value added: ' + value);
                            }
                        }
                    }
                }
            }
        } else {
            logger('No item groups found for form: ' + form.name);
        }
    }

    logger('Collected num values: ' + numValues.join(', '));
    return numValues;
}

// Step 3: Calculate the max num value
function getMaxNumValue(numValues) {
    if (numValues.length === 0) {
        return 1;
    }

    var maxValue = numValues[0];
    logger('Initial max num value: ' + maxValue);

    for (var i = 1; i < numValues.length; i++) {
        if (numValues[i] > maxValue) {
            maxValue = numValues[i];
        }
    }

    logger('Final max num value: ' + maxValue);
    return parseInt(maxValue, 10) + 1;
}

// Execution
try {
    var numValues = getNumValues(formData);
    logger('All num values: ' + numValues.join(', '));

    var nextNum = getMaxNumValue(numValues);
    logger('Next num value (forced integer): ' + nextNum);

    return nextNum.toFixed(0);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
