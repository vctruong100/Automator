const formname = formJson.form.name;

if (containsValue(formname, "hypoglycemic event")) {
    return "Hypoglycemic event";
}

const AE = findFirstItemValueByName(formJson, "AE_Adverse event");

if (AE) return AE;

return null;

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}