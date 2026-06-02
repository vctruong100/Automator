try {
    return formJson.form.subject.volunteer.initials;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}