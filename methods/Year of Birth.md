const dob = formJson.form.subject.volunteer.dateOfBirth;

try {
    return getYearOnly(dob);
} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
}

function getYearOnly(dateString) {
    if (!dateString) return "";
    return dateString.split("-")[0];
}