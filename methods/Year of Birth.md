const dob = formJson.form.subject.volunteer.dateOfBirth;

return getYearOnly(dob);

function getYearOnly(dateString) {
    if (!dateString) return "";
    return dateString.split("-")[0];
}