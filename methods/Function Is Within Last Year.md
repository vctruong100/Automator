function isWithinLastYear(inputDate) {
    if (!inputDate) return false;

    var parts = inputDate.split("-");
    if (parts.length !== 3) return false;

    var parsedDate = new Date(parts[0], parts[1] - 1, parts[2]);
    if (isNaN(parsedDate.getTime())) return false;

    var today = new Date();
    var oneYearAgo = new Date();
    oneYearAgo.setFullYear(today.getFullYear() - 1);

    parsedDate.setHours(0, 0, 0, 0);
    today.setHours(0, 0, 0, 0);
    oneYearAgo.setHours(0, 0, 0, 0);
    logger("Parsed Date: " + parsedDate);
    logger("One year ago: " + oneYearAgo);
    logger("Today: " + today);
    logger(parsedDate >= oneYearAgo && parsedDate <= today);
    return parsedDate >= oneYearAgo && parsedDate <= today;
}