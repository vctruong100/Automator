// Version: v1
// Purpose: Extracts or computes the subject's year of birth.

const dob = formJson.form.subject.volunteer.dateOfBirth;

try {
    return getYearOnly(dob);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Extracts the year component from an ISO 8601 date string.
function getYearOnly(dateString) {
    if (!dateString) return "";
    return dateString.split("-")[0];
}