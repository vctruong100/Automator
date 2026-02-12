const currentStudyEvent = formJson.form.studyEventName;
const item = itemJson.item;
const sigfig = itemJson.item.significantDigits;

const heightitemList = [
    "BM_HT", 
    "BM_HT_Visit 2",
    "BMI_HEIGHT"
];
const weightitemList = [
    "BM_WT #1", 
    "BM_WT #2",
    "BMI_WEIGHT",
    
];
const screeningStudyEvent = [
    "Screening", 
    "SCREENING"
];
const screeningBMI_Form = [
    "BM_Height/Weight/BMI",
    "📏 SCREEN BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)",
    "📏 BM_BODY MEASUREMENTS (HEIGHT / WEIGHT / BMI)"
];

var weight = 0;
var height = 0;
var bmi = 0;

var form = pullForm(screeningStudyEvent, screeningBMI_Form);
if (!form) return null;
height = pullItemFromForm(form, heightitemList);

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

log();

if (bmi) return bmi.toFixed(sigfig);

return null;

function log() {
    logger("List: " + list);
    logger("List length: " + list.length)
    logger("Max count: " + maxCount);
    logger("Average: " + avg)
    logger("Weight: " + weight);
    logger("Height: " + height);
    logger("Height in meter: " + heightMtr)
    logger("Factor: " + factor);
    logger("BMI: " + bmi);
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
            if (["Average Weight"].indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
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