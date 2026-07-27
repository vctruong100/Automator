/* jshint strict: false */

// Version: v1
// Description: Extracts the subject year of birth from the volunteer date-of-birth value and returns only the four-digit year.

var dob = formJson.form.subject.volunteer.dateOfBirth;

function getYearOnly(dateString) {
    if (!dateString) return "";
    return dateString.split("-")[0];
}

try {
    return getYearOnly(dob);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

