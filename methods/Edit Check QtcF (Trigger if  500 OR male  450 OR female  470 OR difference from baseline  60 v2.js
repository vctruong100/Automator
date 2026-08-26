/* jshint strict: false */

// Version: v2
// Description: Runs QTcF safety logic using keyword-based QTcF/Fridericia detection instead of exact item names. Triggers when the current QTcF is at least 500 msec, exceeds sex-specific thresholds, or increases by at least 60 msec from baseline.

var baselineForm = [
    "DAY-1 ECG SINGLE 12 LEAD V1",
    "ECG_Predose_Triplicate ECG (baseline) (SPONSOR PROVIDED MACHINE)"
];
var baselineEvent = [
    "Day -1",
    "Visit 2 Week 1 Day 0"
];

var maxRange = 500;
var maleRange = 450;
var femaleRange = 470;
var differenceRange = 60;

var item = itemJson.item;
var sexMale = formJson.form.subject.volunteer.sexMale;
var diff = 0;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function isQtcfItem(itemName) {
    var name = normalizeName(itemName);
    return containsValue(name, "QTCF") ||
        (containsValue(name, "FRIDERICIA") && containsValue(name, "QTC")) ||
        containsValue(name, "QT CORRECTED BY FRIDERICIA");
}

function isBaselineQtcfItem(itemName) {
    return containsValue(itemName, "BASELINE");
}

function addNumericValue(list, sourceItem) {
    if (!sourceItem || sourceItem.canceled || sourceItem.value === null || sourceItem.value === undefined || sourceItem.value === "") return;

    var numericValue = Number(sourceItem.value);
    if (!isNaN(numericValue)) {
        list.push({
            name: sourceItem.name,
            value: numericValue
        });
    }
}

function pullLatestQtcfFromForm(formJsonValue, includeBaselineItems) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;
    var matches = [];

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;
        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (isQtcfItem(groupItem.name) && (includeBaselineItems || !isBaselineQtcfItem(groupItem.name))) addNumericValue(matches, groupItem);
        }
    }

    if (matches.length < 1) return null;

    var latest = matches[matches.length - 1];
    logger("QTcF keyword matches: " + matches.map(function(match) { return match.name + "=" + match.value; }).join(", "));
    logger("Latest QTcF value used: " + latest.value + " from " + latest.name);
    return latest.value;
}

function getCurrentQtcfValue() {
    var numericItemValue = Number(item.value);
    if (item.value !== null && item.value !== undefined && item.value !== "" && !isNaN(numericItemValue)) {
        logger("Using attached item value as current QTcF: " + numericItemValue);
        return numericItemValue;
    }
    logger("Attached item value is blank/non-numeric; pulling current QTcF by keyword");
    return pullLatestQtcfFromForm(formJson, false);
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
    return null;
}

function checkBaseline(itemValue, baseline) {
    if (baseline !== null && baseline !== undefined && !isNaN(baseline)) {
        diff = itemValue - baseline;
        logger("Difference: " + diff);
        if (diff >= differenceRange) {
            return true;
        }
    }
    return false;
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
        customErrorMessage("QTcF Difference from Baseline OOR > " + differenceRange + ". REPEAT 2 TIMES: " + value);
        return false;
    }
}

try {
    var currentQtcf = getCurrentQtcfValue();

    logger("Is it male: " + sexMale);
    logger("QTcF value: " + currentQtcf);

    if (currentQtcf === null || currentQtcf === undefined || isNaN(currentQtcf)) return null;
    if (currentQtcf > maxRange) return log("500", currentQtcf);
    else if (sexMale && currentQtcf > maleRange) return log("male", currentQtcf);
    else if (!sexMale && currentQtcf > femaleRange) return log("female", currentQtcf);

    var form = pullForm(baselineEvent, baselineForm);
    if (!form) return null;
    var baseline = pullLatestQtcfFromForm(form, true);
    logger("Baseline: " + baseline);

    if (checkBaseline(currentQtcf, baseline)) return log("base", diff);

    return true;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
