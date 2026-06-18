/* jshint strict: false */

// Version: v1
// Purpose: Protocol check for averaged QTcF/QRS out-of-range (non-repeat).

// Add Item names
var QTcFitems = [
    '- QTcF (≤ 450 msec "Males") (≤ 470 msec "Females")',
    "QTcF",
    "QTcF_Protocol",
    "QTcF #1",
    "QTcF #2",
    "QTcF #3"
];

var QRSitems = [
    "- QRS duration ( ≤ 120 ms)",
    "QRS",
    "QRS_Protocol"
];

// Inclusive (Edit)
var QTcF_max_range = 450;
var QRS_max_range = 120;
var count = 0;

var item = itemJson.item.id;

function populateList(form, targetItem, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
	if (!itemGroups || itemGroups.length < 1) return null;

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (item && targetItem.indexOf(item.name) !== -1) {
                    count++;
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        if (list.length >= count) return list;
                    }
                }
            }
        }
    }
    else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (item && targetItem.indexOf(item.name) !== -1) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        if (list.length >= count) return list;
                    }
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

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

try {
    var isRepeat = false;
    var itemRaw = getItemDataContextByItemDataId(item.id);
    var context = JSON.parse(itemRaw);
    if (containsValue(context.foundItemGroupName, "repeat")) isRepeat = true;

    var QTcFlist = populateList(formJson, QTcFitems, isRepeat);
    var QTcFavg = calculateAverage(QTcFlist);

    var QRSlist = populateList(formJson, QRSitems, isRepeat);
    var QRSavg = calculateAverage(QRSlist);

    if (QTcFlist.length < count || QRSlist.length < count) return itemJson.item.codeListItems[0].codedValue; // return pending result

    if (QTcFavg > QTcF_max_range || QRSavg > QRS_max_range) return itemJson.item.codeListItems[2].codedValue; // return Out of protocol range
    else if (QTcFavg <= QTcF_max_range || QRSavg > QRS_max_range) return itemJson.item.codeListItems[1].codedValue; // return within protocol range

    return itemJson.item.codeListItems[0].codedValue;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
