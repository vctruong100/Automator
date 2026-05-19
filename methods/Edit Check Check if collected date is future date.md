const val = itemJson.item.value;

logger(val);

return isFutureDate(val);

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