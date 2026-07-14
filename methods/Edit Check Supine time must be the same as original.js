/* jshint strict: false */

// Version: v1
// Purpose: Check if the supine time for repeated form is the same as the original
// Return true if they are the same OR if the difference between the collected time and the time from original form is more than 1 hour

var studyEvent = formJson.form.studyEventName;
var formid = formJson.form.id;
var formName = formJson.form.name;
var item = itemJson.item;
var itemDataType = item.dataType;
var isCondition = false;
logger(itemDataType);

function checkForm(studyevent, form) {
    if (!form) {
        return formJson.form;
    } else {
        var arrayForms = findFormData(studyevent, form);
        var completedForm = collectCompleted(arrayForms, true);
        if (!completedForm || completedForm.length === 0) return null;
        return completedForm;
    }
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
    var itemGroups = form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1) return item;
        }
    }
    return null;
}

function formatDateTime(isoString) {
    if (!isoString) return "";

    var parts = isoString.split("T");
    if (parts.length < 2) return "";

    var dateParts = parts[0].split("-");
    var timeParts = parts[1].split(":"); 

    if (dateParts.length < 3 || timeParts.length < 3) return "";

    var year = dateParts[0];
    var month = parseInt(dateParts[1], 10) - 1; 
    var day = dateParts[2];

    var hour = timeParts[0];
    var minute = timeParts[1];
    var second = timeParts[2];

    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    return day + " " + months[month] + " " + year + " "
           + hour + ":" + minute + ":" + second;
}

try {
    var forms = checkForm(studyEvent, formName);
    
    if (!forms || forms.length < 1) {
        return true;
    }

    var ogForm = forms[forms.length - 1].form;
    var ogFormId = ogForm.id;

    if (formid == ogFormId) {
        return true;
    }

    var ogFormItemValue = pullItemFromForm(ogForm, item.name);
    var currentMs = item.dateValueMs;
    var ogMs = ogFormItemValue.dateValueMs;

    logger("Current form id: " + formid + ", Current input date; " + item.value);
    logger("Original form id: " + ogFormId + ", Original Date: " + ogFormItemValue.value);

    var differenceMs = Math.abs(currentMs - ogMs);
    var oneHourMs = 60 * 60 * 1000;

    logger("Difference (minutes): " + (differenceMs / 60000));
    if (itemDataType == "datetime") {
        var parseCurrentItemDate = formatDateTime(item.value);
        var parseOriginalDate = formatDateTime(ogFormItemValue.value);
        if (parseCurrentItemDate === parseOriginalDate || differenceMs > oneHourMs) {
            return true;
        }
        else {
            customErrorMessage("Original resting time: " + parseOriginalDate);
            return false;
        }
    }
    else if (ogFormItemValue === item.value) {
        return true;
    }

    customErrorMessage("Original resting time: " + ogFormItemValue);
    return false;
} catch (e) {
    logger("Error: " + e);
    return false;
}
