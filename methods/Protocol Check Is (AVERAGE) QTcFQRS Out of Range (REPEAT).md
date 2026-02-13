const QTcFitems = [
    '- QTcF (≤ 450 msec "Males") (≤ 470 msec "Females")', 
    "QTcF", 
    "QTcF_Protocol",
    "QTcF #1",
    "QTcF #2", 
    "QTcF #3"
];

const QRSitems = [
    "- QRS duration ( ≤ 120 ms)", 
    "QRS", 
    "QRS_Protocol"
];

const attachedItemCodeList = [
    "⭕Pending Results",
    "✅ Within Protocol Range",
    "🛑 Out of protocol range, SF"
]

// Inclusive
var QTcF_max_range = 450;
var QRS_max_range = 120;

var QTcFmaxCount = 3; 
var QTcFlist = [];
var QTcFavg = 0;

var QRSmaxCount = 3; 
var QRSlist = [];
var QRSavg = 0;

QTcFlist = populateList(formJson, QTcFitems, QTcFlist, QTcFmaxCount);
QTcFavg = calculateAverage(QTcFlist);

QRSlist = populateList(formJson, QRSitems, QRSlist, QRSmaxCount);
QRSavg = calculateAverage(QRSlist);

log();

if (QTcFlist.length !== QTcFmaxCount || QRSlist.length !== QRSmaxCount) return attachedItemCodeList[0];

if (QTcFavg > QTcF_max_range || QRSavg > QRS_max_range) return attachedItemCodeList[2];
else if (QTcFavg <= QTcF_max_range || QRSavg > QRS_max_range) return attachedItemCodeList[1];

return attachedItemCodeList[0];

function log() {
    logger("List: " + QTcFlist);
    logger("List length: " + QTcFlist.length);
    logger("Max count: " + QTcFmaxCount);
    logger("Average: " + QTcFavg);
    
    logger("List: " + QRSlist);
    logger("List length: " + QRSlist.length);
    logger("Max count: " + QRSmaxCount);
    logger("Average: " + QRSavg);
}

function populateList(form, targetItem, list, maxCount) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                    if (list.length >= maxCount) return list;
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