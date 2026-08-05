/* jshint strict: false */

// Version: v1
// Purpose: Validates time difference between start and collection.

// Add Item Names
var startTimeItem = [
    "VS_Resting time (NEW)", "VS_Resting time", "Same as ECG supine time", "Start Supine time", "ECG_Supine resting time","Start Semirecumbent Time", 
]

var studyevent = formJson.form.studyEventName;
logger("Study event: " + studyevent)

var formName = [
    "⚡ECG_Single 12 - Lead ECG (SCRN/D-1)",
]

var confirmationItemName = ["Did subject remain supine from ECG collection?"]
var difference = 5; // in minutes

// Two Approaches: Method A and Method B
// Let's say Start Time is 10:00:59 and End Time is 10:10:00
// If using Method A, this will calculate the actual difference (9 minutes and 1 second) and round down: 9 minutes.
// If using Method B, this will round down both Time (10:00:00 and 10:10:00) and calculate the difference: 10 minutes.
// Adjust ("A" or "B") below on what method to use.

var methodType = "B";
var form = formJson.form;

function pullItemFromForm(form, targetItem, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    if (!isRepeat) {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled || containsValue(group.name, "repeat")) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1) {
                    logger("Start Time: " + item.name + ", value: " + item.value);
                    if (item.value !== null && !item.canceled && item.value !== "") return item;
                }
            }
        }
    }
    else {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1) {
                    logger("Start Time: " + item.value);
                    if (item.value !== null && !item.canceled && item.value !== "" && item.dataType == "datetime") return item;
                }
            }
        }
    }
    return null;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function getItemGroupName(form) {
    var item = itemJson.item;
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                return group.name;
            }
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function checkForm(studyEvent, formName) {
  var forms = findFormData(studyEvent, formName);
  var completed = collectCompleted(forms, true);
  if (!completed || completed.length === 0) return null;
  return completed[0];
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
    var groupName = getItemGroupName(formJson);
    var isRepeat = false;
    if (containsValue(groupName, "repeat") || containsValue(groupName, "standing")) isRepeat = true;
    var confirmation = pullItemFromForm(formJson, confirmationItemName);
    var startTime = null;
    if (confirmation == "YES") {
        var ecgForm = pullForm([studyevent], formName);
        if (!ecgForm) return null;
        var startTime = pullItemFromForm(ecgForm, startTimeItem, false);
    }
    else {
        startTime = pullItemFromForm(formJson, startTimeItem, isRepeat);
    }
    var endTime = itemJson.item;

    logger("Start time: " + startTime.value)
    logger("Collected Time: " + endTime.value);
    
    if (!startTime || startTime.value == null || startTime.value == "" || !endTime || endTime.value == null || endTime.value == "") return null;

    var startTimeMs = startTime.dateValueMs;
    logger("STart time ms: " + startTimeMs)
    var endTimeMs = endTime.dateValueMs;

    var differenceInMins;

    if (methodType === "A") {

        // Method A: true difference, then floor
        var diffMs = endTimeMs - startTimeMs;
        differenceInMins = Math.floor(diffMs / (1000 * 60));

        logger("Method A used");

    } else {

        // Method B: floor each first, then subtract
        var startMin = Math.floor(startTimeMs / (1000 * 60));
        var endMin = Math.floor(endTimeMs / (1000 * 60));

        differenceInMins = endMin - startMin;

        logger("Method B used");
        logger("Start (min): " + startMin);
        logger("End (min): " + endMin);
    }

    logger("Diff (min): " + differenceInMins);

    if (differenceInMins < 0) {
        return false;
    }

    if (differenceInMins >= difference) {
        return true;
    }

    return false;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}