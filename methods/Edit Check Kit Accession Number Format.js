/* jshint strict: false */

// Version: v1
// Purpose: Validates kit accession number format.

var value = itemJson.item.value;

function isValidFormat(value)
{
    var pattern = /^[A-Za-z]{2}[0-9]{5}$/;

    if (typeof value !== "string") return false
    if (pattern.test(value)) return true;
    else return false;
}

try {
    return isValidFormat(value);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
