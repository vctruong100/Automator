/* jshint strict: false */

// Version: v1
// Purpose: QTcF safety edit check with sex-specific thresholds.

// Add item names
var baselineForm = [
    "⚡DAY-1 ECG SINGLE 12 LEAD V1",
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)",

]
var baselineEvent = [
    "Day -1",
    "Visit 2 Week 1 Day 0",
]
var baselineItem = [
    "Fridericia QTc Interval.",
    "Fridericia QTc Interval",
    "Baseline QTcF",
    "QTcF.",
];

var maxRange = 500;
var maleRange = 450;
var femaleRange = 470;
var differenceRange = 60;

var item = itemJson.item;
var sexMale = formJson.form.subject.volunteer.sexMale;
var diff = 0;

function normalizeItemName(name) {
    if (!name) return "";
    return name.toString().replace(/\s+/g, "").toLowerCase();
}

function containsItemName(itemList, itemName) {
    var normalizedName = normalizeItemName(itemName);

    for (var i = 0; i < itemList.length; i++) {
        if (normalizeItemName(itemList[i]) === normalizedName) {
            return true;
        }
    }
    return false;
}
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

  if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null) return item.value;
        }
    }
    return null;
}

function checkBaseline(itemValue) {
    if (baseline !== null && baseline !== undefined) {
        diff = (itemValue - baseline)
        logger("Difference: " + diff)
        if (diff >= differenceRange) {
            return true;
        }
        return false;
    }
}
function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' ||
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            keepers.push(formData);
        } else {

        }
    }
    return keepers;
}

function log(reason, value) {
    if (reason === "500") {
        customErrorMessage("OOR > " + maxRange + ": REPEAT 2 TIMES: " + value);
        return false;
    } else if (reason === "male") {
        customErrorMessage("Male QTcF OOR > " + maleRange + ": " + value);
        return false;
    } else if (reason === "female") {
        customErrorMessage("Female QTcF OOR > " + femaleRange + ": " + value);
        return false;
    } else if (reason === "base") {
        customErrorMessage("QTcF Difference from BaselineOOR > " + differenceRange + ". REPEAT 2 TIMES:" + value);
        return false;
    }
}

try {
    logger("Is it male: " + sexMale);
    logger("Qtcf value: " + item.value);
    if (item.value > maxRange) return log("500", item.value);
    else if (sexMale && item.value > maleRange) return log("male", item.value);
    else if (!sexMale && item.value > femaleRange) return log("female", item.value);

    var form = pullForm(baselineEvent, baselineForm);
    if (!form) return null;
    var baseline = pullItemFromForm(form, baselineItem);
    logger("Baseline: " + baseline);

    if (checkBaseline(itemJson.item.value)) return log("base", diff);

    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
