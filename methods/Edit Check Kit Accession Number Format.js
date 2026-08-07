/* jshint strict: false */

// Version: v1
// Description: Validates that the value consists of exactly 2 letters followed by 6 digits.

var value = itemJson.item.value;

function isValidFormat(value)
{
    var pattern = /^[A-Za-z]{2}[0-9]{6}$/;

    if (typeof value !== "string") return false;
    return pattern.test(value);
}

try {
    return isValidFormat(value);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}