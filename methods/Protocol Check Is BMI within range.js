/* jshint strict: false */

// Version: v2
// Purpose: Protocol BMI range validation.
// Item-name independent: matches HEIGHT / WEIGHT via standalone keyword checks.

// inclusive
var BMI_lower_range = 18;
var BMI_upper_range = 30;

var currentStudyEvent = formJson.form.studyEventName;
var item = itemJson.item;
var sigfig = 1;

var screeningStudyEvent = ["Screening", "SCREENING"];
var screeningBMI_Form = [
    "📏 BM_SCRN | HEIGHT | WEIGHT | BMI",
    "BM_Height / Weight / BMI",
    "BM_Height/Weight/BMI",
    "📏 SCREEN BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)",
    "📏 BM_BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)"
];

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function containsStandaloneKeyword(input, keyword) {
    var value = normalizeName(input);
    var target = normalizeName(keyword);
    var startIndex = 0;
    var index;
    var before;
    var after;

    while (startIndex < value.length) {
        index = value.indexOf(target, startIndex);
        if (index === -1) return false;

        before = index === 0 ? "" : value.charAt(index - 1);
        after = index + target.length >= value.length ? "" : value.charAt(index + target.length);

        if ((before === "" || !/[A-Z0-9]/.test(before)) && (after === "" || !/[A-Z0-9]/.test(after))) return true;

        startIndex = index + target.length;
    }

    return false;
}

function matchesMetric(itemName, metric) {
    var name = normalizeName(itemName);

    if (metric === "HEIGHT") return containsStandaloneKeyword(name, "HEIGHT") || containsStandaloneKeyword(name, "HT");
    if (metric === "WEIGHT") return containsStandaloneKeyword(name, "WEIGHT") || containsStandaloneKeyword(name, "WT");

    return false;
}

function isAverageItem(itemName) {
    if (containsValue(itemName, "AVERAGE")) return true;
    if (containsValue(itemName, "AVG")) return true;
    if (containsValue(itemName, "MEAN")) return true;

    return false;
}

function addNumericValue(list, value) {
    if (value === null || value === undefined || value === "") return;

    var numericValue = parseFloat(value);
    if (!isNaN(numericValue)) list.push(numericValue);
}

function getFirstMetricValue(formData, metric) {
    var itemGroups = formData.form.itemGroups;
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;
        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem || groupItem.value == null || groupItem.value === "") continue;

            if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                logger(metric + " first value matched item: " + groupItem.name + " | Value: " + groupItem.value);
                return parseFloat(groupItem.value);
            }
        }
    }

    return null;
}

function getMetricMeasurements(formData, metric) {
    var itemGroups = formData.form.itemGroups;
    var result = { values: [], count: 0 };
    var group, items, groupItem, i, j;

    if (!itemGroups || itemGroups.length < 1) return result;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;
        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;

            // Stop collecting raw measurements once we hit an average/mean item for this metric
            if (isAverageItem(groupItem.name) && matchesMetric(groupItem.name, metric)) {
                return result;
            }

            if (matchesMetric(groupItem.name, metric) && !isAverageItem(groupItem.name)) {
                result.count++;
                logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                addNumericValue(result.values, groupItem.value);
            }
        }
    }

    return result;
}

function calculateAverage(values) {
    if (!values || values.length === 0) return null;

    var sum = 0;
    var count = 0;
    for (var i = 0; i < values.length; i++) {
        if (isNaN(values[i])) continue;
        sum += values[i];
        count++;
    }

    if (count === 0) return null;

    return sum / count;
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

try {
    var height = null;

    if (containsValue(currentStudyEvent, "screening")) {
        height = getFirstMetricValue(formJson, "HEIGHT");
    } else {
        var screeningForm = pullForm(screeningStudyEvent, screeningBMI_Form);
        if (screeningForm) {
            height = getFirstMetricValue(screeningForm, "HEIGHT");
        }
    }

    if (!height) {
        height = getFirstMetricValue(formJson, "HEIGHT");
    }

    var weightMeasurements = getMetricMeasurements(formJson, "WEIGHT");
    var weight = null;

    if (weightMeasurements.count > 0 && weightMeasurements.values.length === weightMeasurements.count) {
        weight = calculateAverage(weightMeasurements.values);
    }

    logger("BMI Height: " + height);
    logger("BMI Weight: " + weight);

    if (!weight || weight == 0 || !height || height == 0) return null;

    var heightMtr = height / 100;
    var bmi = weight / (heightMtr * heightMtr);
    var factor = Math.pow(10, sigfig);
    bmi = Math.round(bmi * factor) / factor;
    logger("BMI: " + bmi);

    if (BMI_lower_range <= bmi && bmi <= BMI_upper_range) return item.codeListItems[0].codedValue; // return Yes
    else if (bmi < BMI_lower_range || bmi > BMI_upper_range) return item.codeListItems[1].codedValue; // return No, SF

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
