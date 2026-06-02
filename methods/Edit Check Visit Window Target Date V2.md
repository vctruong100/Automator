const studyevents = [
    "Day 1"
]
const formName = [
    "🟡IP_EVOLOCUMAB ADMINISTRATION",
]
const itemName = [
    "IP_StartDate"
]

return true; 
var currentStudyName = formJson.form.studyEventName;
var item = itemJson.item;

try {
    logger("Study event: " + currentStudyName);
    if (currentStudyName == "Screening") return true;
    var day = parseInt(currentStudyName.split(" ")[1]);
    logger("Day: " + day);
    
    if (day < 9) return true;
    
    var form = pullForm(studyevents, formName);
    
    if (!form) return null;
    var val = pullItemFromForm(form, itemName);
    
    if (!val || !item || !item.dateValueMs) return false;
    logger("Start Date: " + val.value);
    logger("Collected Date: " + item.value);
    var baseDate = parseDate(val.value);
    if (!baseDate) return false;
    
    var addDays = day - 1;
    logger("Add days: " + addDays)
    var allowedRange = day <= 15 ? 1 : 2;
    
    logger("Allowed range: " + allowedRange)

    var targetMs = baseDate.getTime() + addDays * 86400000;
    var diffDays = Math.floor(Math.abs(item.dateValueMs - targetMs) / 86400000);
    logger("Differences: " + diffDays);
    return diffDays <= allowedRange;

}
catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
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

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
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
function parseDate(dateStr) {
    if (!dateStr) return null;
    var clean = dateStr.split('T')[0];
    var parts = clean.split('-');
    if (parts.length !== 3) return null;
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var day = Number(parts[2]);
    if (!year || !month || !day) return null;
    return new Date(year, month - 1, day);
}

function formatDateText(dateObj) {
    var MMM = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    var d = dateObj.getDate();
    var m = MMM[dateObj.getMonth()];
    var y = dateObj.getFullYear();
    var hy = "\u2011"; 
    return (d < 10 ? "0"+d : ""+d) + hy + m + hy + y; 
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
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item;
        }
    }
    return null;
}