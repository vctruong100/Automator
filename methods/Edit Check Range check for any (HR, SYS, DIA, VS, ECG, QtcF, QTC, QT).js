/* jshint strict: false */

// Version: v1
// Purpose: Lightweight range check for any vitals or ECG parameter.

var value = itemJson.item.value;
var minimum_range = 60;
var maximum_range = 100;

try {
    if (value >= minimum_range && value <= maximum_range) {
        return true;
    }

    return false;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
