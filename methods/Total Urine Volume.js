var formName = formJson.form.name;
var studyevent = formJson.form.studyEventName;
var urineJugItem = "Weight of Void Jug+Urine";
var emptyJugItem = "Weight of Empty Void Jug";
var sigfig = itemJson.item.significantDigits;
var urineJug = parseInt(pullItemFromForm(formJson.form, urineJugItem));
var voidJug = parseInt(pullItemFromForm(formJson.form, emptyJugItem));
var diff = (urineJug - voidJug);
var volume = diff / 1.02;
if (volume) return Math.round(volume).toFixed(sigfig);
if (volume == 0) return (0).toFixed(sigfig);
return null;

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.itemGroups;
    var group, items, item, i, j;
    var total = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item && item.value !== null && !item.canceled) {
                logger(item.name);
                logger(item.value);
                total += parseInt(item.value);
            }
        }
    }
    return total;
}
function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm;
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