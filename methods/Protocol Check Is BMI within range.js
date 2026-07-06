/* jshint strict: false */

// Version: v1
// Purpose: Protocol BMI range validation.

// Add item names
var heightitemList = [
    "BM_HT",
    "BM_HT_Visit 2",
    "BMI_HEIGHT"
];
var weightitemList = [
    "BM_WT #1",
    "BM_WT #2",
    "BMI_WEIGHT",

];
var screeningStudyEvent = [
    "Screening",
    "SCREENING"
];
var screeningBMI_Form = [
    "BM_Height/Weight/BMI",
    "📏 SCREEN BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)",
    "📏 BM_BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)"
];

// inclusive
var BMI_lower_range = 18;
var BMI_upper_range = 30;

var currentStudyEvent = formJson.form.studyEventName;
var item = itemJson.item;
var sigfig = itemJson.item.significantDigits;

var weight = 0;
var height = 0;
var bmi = 0;

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
function populateList(form, targetItem, list) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(["Average Weight"], item.name)) return list;
            if (item && containsItemName(targetItem, item.name)) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                }
            }
        }
    }
    return list;
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

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null) return item.value;
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
        } else {

        }
    }
    return keepers;
}

function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    // var factor = Math.pow(10, sigfig);
    return avg;
}

try {
    var form = pullForm(screeningStudyEvent, screeningBMI_Form);
    if (!form) {
        height = pullItemFromForm(formJson, heightitemList);
    }
    else {
        height = pullItemFromForm(form, heightitemList);
    }

    var maxCount = 0;
    var list = [];
    var avg = 0;

    list = populateList(formJson, weightitemList, list);

    avg = calculateAverage(list, sigfig);

    if (list.length === maxCount) {
        weight = avg;
    }

    if (!weight || weight == 0 || !height || height == 0) return null;
    var heightMtr = height / 100;

    var factor = Math.pow(10, sigfig);

    bmi = Math.round((weight / (heightMtr * heightMtr)) * factor) / factor;

    if (BMI_lower_range <= bmi && bmi <= BMI_upper_range) return item.codeListItems[0].codedValue; // return Yes
    else if (bmi < BMI_lower_range || bmi > BMI_upper_range) return item.codeListItems[1].codedValue; // return No, SF

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
