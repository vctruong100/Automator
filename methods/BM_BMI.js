/* jshint strict: false */

// Version: v1
// Purpose: Calculates BMI from height and weight outside of Screening.

// Add item names

var heightitemList = [
    "VS_HEIGHT",
    "BM_HT",
    "BM_HT_Visit 2",
    "BMI_HEIGHT"
];
var weightitemList = [
    "VS_WEIGHT",
    "BM_WT #1",
    "BM_WT #2",
    "BMI_WEIGHT",

];
var screeningStudyEvent = [
    "Screening",
    "SCREENING"
];
var screeningBMI_Form = [
    "BM_Height / Weight / BMI",
    "BM_Height/Weight/BMI",
    "📏 SCREEN BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)",
    "📏 BM_BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)"
];

var currentStudyEvent = formJson.form.studyEventName;
var item = itemJson.item;
var sigfig = itemJson.item.significantDigits;

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
function populateList(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(["Average Weight"], item.name)) return list;
            if (item && containsItemName(targetItem, item.name)) {
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
    var height = pullItemFromForm(formJson, heightitemList);
    var weight = 0;
    if (!height || height == null) {
        var form = pullForm(screeningStudyEvent, screeningBMI_Form);
        if (form) {
        height = pullItemFromForm(form, heightitemList);
        }
    }
    
    var list = populateList(formJson, weightitemList);
    if (list.length < 1) return null;
    if (list.length == 1) weight = list[0];
    else weight = calculateAverage(list, sigfig);

    logger("Weight: " + weight);
    logger("Height: " + height);
    if (!weight || weight == 0 || !height || height == 0) return null;
    var heightMtr = height / 100;

    var factor = Math.pow(10, sigfig);

    var bmi = Math.round((weight / (heightMtr * heightMtr)) * factor) / factor;

    if (bmi) return bmi.toFixed(sigfig);

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
