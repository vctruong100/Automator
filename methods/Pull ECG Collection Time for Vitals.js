/* jshint strict: false */

// Version: v1
// Description: Pulls and formats the ECG collection or supine start time from a completed ECG form in the same study event for use on vital-sign forms.

var formName = formJson.form.name;
var studyEventName = formJson.form.studyEventName;

var confirmationItems = [
    "Remained supine since ECG?"
];

var item = [
    "START SUPINE",
    "ECG_Supine resting time",
];

var ecgFormNames = [
 "⚡ D-1 12-LEAD ECG (SINGLE) V1.0",
  "⚡ 12-LEAD ECG (SINGLE) V1.0",
  "⚡12-LEAD ECG (BASELINE) (SINGLE) 1.0",
  "⚡ECG_12-LEAD ECG (TRIPLICATE)",
];

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
function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function formatDate(dateStr) {
    if (!dateStr) return null;

    dateStr = dateStr.trim();

    var year = dateStr.substring(0, 4);
    var monthNum = dateStr.substring(5, 7);
    var day = dateStr.substring(8, 10);
    var hours = dateStr.substring(11, 13);
    var minutes = dateStr.substring(14, 16);
    var seconds = dateStr.substring(17, 19);

    var months = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
        "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
        "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
    };

    return day + " " + months[monthNum] + " " + year + " " +
           hours + ":" + minutes + ":" + seconds;
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


function checkForm(studyEvent, formName) {
  var forms = findFormData(studyEvent, formName);
  var completed = collectCompleted(forms, true);
  if (!completed || completed.length === 0) return null;
  return completed[completed.length - 1]; // most recent
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
  if (!formDataArray) return [];
  var keepers = [];
  for (var i = formDataArray.length - 1; i >= 0; i--) {
    var f = formDataArray[i];
    if (
      f.form.canceled === false &&
      (
        f.form.dataCollectionStatus === "Complete" ||
        f.form.dataCollectionStatus === "Incomplete" ||
        (INCLUDE_NONCONFORMANT_DATA === true &&
         f.form.dataCollectionStatus === "Nonconformant")
      )
    ) {
      keepers.push(f);
    }
  }
  return keepers;
}

try {
    var confirmation = pullItemFromForm(formJson, confirmationItems)

    if (!confirmation || confirmation == null) return null;
    if (confirmation == "NO") return null;
    logger(confirmation);
    form = pullForm([studyEventName], ecgFormNames);
    if (!form) return null;

    var result = pullItemFromForm(form, item);
    if (!result || result == null) return null;

    return formatDate(result);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}