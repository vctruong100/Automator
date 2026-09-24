/* jshint strict: false */

// Version: v2
// Description: Validates that a repeat-form supine or semi-recumbent time matches an original source time by comparing the current value against matching source-form values across study events, while allowing first/original forms to pass automatically.

var studyEvent = formJson.form.studyEventName;
var formid = formJson.form.id;
var formName = formJson.form.name;
var item = itemJson.item;
var itemDataType = item.dataType;
logger(itemDataType);

var forms = checkForm(studyEvent, formName);
if ((!forms || forms.length < 1) && studyEvent != "Unscheduled") {
    logger("First form and is not unscheduled")
    return true;
}
var ogForm = forms[forms.length - 1].form;
var ogFormId = ogForm.id;

if (formid == ogFormId && studyEvent != "Unscheduled") {
    logger("Original Form")
    return true;
}

var formData = findFormDataAcrossStudyEvents(formName, false);
logger(formData.length);
logger("Collected value: " + formatDateTime(item.value))
return findMatchingTime(formData, item);

function findMatchingTime(formData, item) {
    for (var i = 0; i < formData.length; i++) {
        var form = formData[i].form;
        if (form.canceled) continue;
        logger("Inspecting form: " + form.name + " | Study Event: " + form.studyEventName);
        
        if (form.itemGroups && form.itemGroups.length > 0 && !form.itemGroups[0].canceled && form.itemGroups[0].items[0] && 
            !form.itemGroups[0].items[0].canceled && form.itemGroups[0].items[0].value !== null) {
                logger("Item name: " + form.itemGroups[0].items[0].name + ", item value: " + formatDateTime(form.itemGroups[0].items[0].value));
                if (form.itemGroups[0].items[0].value === item.value) {
                    logger("Original form ID: " + form.id);
                    logger("Repeat form ID: " + formJson.form.id);
                    return true;
                }
            }
    }
    return false;
}

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
            if (targetItem.indexOf(item.name) !== -1) return item.value;
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
if (!forms || forms.length < 1) {
    return true;
}
