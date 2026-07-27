/* jshint strict: false */

// Version: v1
// Description: Returns false when the current collected date is later than today, allowing ClinSpark to query future-dated collection entries.

var val = itemJson.item.value;

function isFutureDate(dateStr) {
    var parts = dateStr.split("-");

    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1;
    var day = parseInt(parts[2], 10);

    var inputDate = new Date(year, month, day);
    var today = new Date();

    today.setHours(0, 0, 0, 0);

    logger("inputDate: " + inputDate);
    logger("today: " + today);

    return inputDate > today;
}

try {
    if (!val) return null;
    logger(val);
    return isFutureDate(val);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
