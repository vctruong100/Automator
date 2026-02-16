const sysItems = [
    "SYS",
    "SYS (1 of 3)",
];

const diaItems = [
    "DIA",
    "DIA (1 of 3)",
];

const hrItems = [
    "HR",
    "HR (1 of 3)",
]

const rrItems = [
    "RR",
]

const tempItems = [
    "ORAL TEMP (35.8C - 37.6C)",
    "ORAL_TEMP",
]
const attachedItemCodeList = [
    "⭕Pending Results",
    "✅ Within Protocol Range",
    "🛑 Out of protocol range, SF",
    "❗Out of Normal Range",
    "✅ Within Normal Range",
    " Yes, within normal range",
    "No",
]

// Inclusive
var sys_min_range = 90;
var sys_max_range = 140;

var dia_min_range = 50
var dia_max_range = 90;

var hr_min_range = 60;
var hr_max_range = 100;

var temp_min_range = 35.6;
var temp_max_range = 37.4;

var rr_min_range = 8;
var rr_max_range = 20;

var sys = pullItemFromForm(formJson, sysItems);
var dia = pullItemFromForm(formJson, diaItems);
var hr = pullItemFromForm(formJson, hrItems);
var temp = pullItemFromForm(formJson, tempItems);
var rr = pullItemFromForm(formJson, rrItems);

log();

if (!sys || sys === null || !dia || dia == null || !hr || hr == null || !temp || temp == null || !rr || rr == null) return attachedItemCodeList[0]

// OOR
if (
    sys > sys_max_range ||
    sys < sys_min_range ||
    dia > dia_max_range ||
    dia < dia_min_range ||
    hr > hr_max_range ||
    hr < hr_min_range ||
    temp > temp_max_range ||
    temp < temp_min_range ||
    rr > rr_max_range ||
    rr < rr_min_range
) return attachedItemCodeList[3]; // Out of Normal Range
else if ( // IR
    sys <= sys_max_range ||
    sys >= sys_min_range ||
    dia <= dia_max_range ||
    dia >= dia_min_range ||
    hr <= hr_max_range ||
    hr >= hr_min_range ||
    temp <= temp_max_range ||
    temp >= temp_min_range ||
    rr <= rr_max_range ||
    rr >= rr_min_range
) return attachedItemCodeList[4]; // Within Normal Range

return attachedItemCodeList[0];

function log() {
    logger("sys: " + sys);
    logger("dia:  " + dia);
    logger("hr: " + hr);
    logger("rr: " + rr);
    logger("temp: " + temp);
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
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

function calculateAverage(values) {
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
    return avg;
}