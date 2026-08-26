/* jshint strict: false */

// Version: v1
var sex = formJson.form.subject.volunteer.sexMale;
var age = formJson.form.subject.volunteer.age;
var dob = formJson.form.subject.volunteer.dateOfBirth;

if (sex) return "Male";
if (!sex) return "Female";

return age;
return dob;