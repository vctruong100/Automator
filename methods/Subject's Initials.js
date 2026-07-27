/* jshint strict: false */

// Version: v1
// Description: Returns the subject initials from available subject-identifying fields for use in target initials fields.

try {
    return formJson.form.subject.volunteer.initials;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
