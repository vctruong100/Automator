/* jshint strict: false */

// Version: v1
// Purpose: Extracts or computes the subject's year of birth.

var dob = formJson.form.subject.volunteer.dateOfBirth;

try {
    return getYearOnly(dob);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

function getYearOnly(dateString) {
    if (!dateString) return "";
    return dateString.split("-")[0];
}
