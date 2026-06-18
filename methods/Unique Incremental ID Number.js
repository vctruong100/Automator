/* jshint strict: false */

// Version: v1
// Purpose: Generates next incremental unique ID across all study events.

var currentAEID = parseInt(itemJson.item.value, 10);
if (!isNaN(currentAEID)) {
    logger('Current AEID is already set: ' + currentAEID + '. Skipping calculation.');
    return currentAEID;  // Return the current AEID if it's already set
} else {
    logger('No AEID set for the current item. Proceeding with calculation.');
}

// Step 1: Retrieve all forms with the name 'AE / CM' across study events
var formData = findFormDataAcrossStudyEvents('AE / CM', false);

// Step 2: Extract AEID values with added logging
function getAEIDValues(formData) {
    var aeidValues = [];
    logger('Form data length: ' + formData.length);

    // Loop through each form in the array
    for (var i = 0; i < formData.length; i++) {
        var form = formData[i].form;
        logger('Inspecting form: ' + form.name);

        // Check if itemGroups exists
        if (form.itemGroups && form.itemGroups.length > 0) {
            // Loop through each itemGroup
            for (var j = 0; j < form.itemGroups.length; j++) {
                var itemGroup = form.itemGroups[j];

                // Check if items exist in the itemGroup
                if (itemGroup.items && itemGroup.items.length > 0) {
                    // Loop through each item
                    for (var k = 0; k < itemGroup.items.length; k++) {
                        var item = itemGroup.items[k];

                        // Check if the item name is 'AEID'
                        if (item.name === "AEID") {
                            var value = parseInt(item.value, 10);  // Ensure the value is parsed as an integer
                            if (!isNaN(value)) {  // Only push valid integers
                                aeidValues.push(value);
                                logger('AEID value added: ' + value);
                            }
                        }
                    }
                }
            }
        } else {
            logger('No item groups found for form: ' + form.name);
        }
    }

    logger('Collected AEID values: ' + aeidValues.join(', '));
    return aeidValues;
}

// Step 3: Calculate the max AEID value
function getMaxAEIDValue(aeidValues) {
    if (aeidValues.length === 0) {
        return 1;  // No previous values, start with 1
    }

    var maxValue = aeidValues[0];
    logger('Initial max AEID value: ' + maxValue);

    for (var i = 1; i < aeidValues.length; i++) {
        if (aeidValues[i] > maxValue) {
            maxValue = aeidValues[i];
        }
    }

    logger('Final max AEID value: ' + maxValue);
    return parseInt(maxValue, 10) + 1;  // Increment the max value by 1 and force integer
}

// Execution
try {
    if (!formData) return 1;
    var aeidValues = getAEIDValues(formData);
    logger('All AEID values: ' + aeidValues.join(', '));

    var nextAEID = getMaxAEIDValue(aeidValues);
    logger('Next AEID value (forced integer): ' + nextAEID);

    return parseInt(nextAEID, 10);  // Return the next AEID as an explicit integer
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
