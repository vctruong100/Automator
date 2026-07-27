/* jshint strict: false */

// Version: v1
// Description: Triggers an edit-check response when a dispensed kit number is present, supporting kit reconciliation or conditional downstream validation.

var itemName = "Dispensed kit number #4";
var item = itemJson.item;

try {
    var value = findFirstItemValueByName(formJson, itemName);
    if (value && value !== null) {
        logger(value);
        if (item && item.value !== null && item.value !== "") return true;
        else {
            customErrorMessage("Please complete when kits are returned")
            return false;
        }
    }

    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
