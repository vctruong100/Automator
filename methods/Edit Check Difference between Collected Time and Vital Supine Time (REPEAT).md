const item = itemJson.item;
const form = formJson.form;
const confirmationItem = [
    "VS_remained in resting position"
];
// If confirmation is YES
const vs_same_as_ecg = [
    "VS_Resting time",
    "vs_same_as_ecg"
];

// If confirmation is NO or doesn't exist
const supinetimeStart = [
    "VS_Resting time (NEW)",
    "VS_Resting time",
    "supine_time_start",
    "vs_same_as_ecg",
];

const difference = 5;

const confirmation = pullItemFromForm(formJson, confirmationItem);

var supineTime = null;
if (confirmation && confirmation == "YES") {
	supineTime = pullItemFromForm(formJson, vs_same_as_ecg);
} else {
	supineTime = pullItemFromForm(formJson, supinetimeStart);
}
if (!supineTime || supineTime == null) return null;

var supineTimeMs = supineTime.dateValueMs;
const collectedTimeMs = item.dateValueMs;

logger("supine Time :"  + formatDateTimeFromMs(supineTimeMs));
logger("collected time: " + formatDateTimeFromMs(collectedTimeMs));

const differenceMs = collectedTimeMs - supineTimeMs;
logger(differenceMs)
if (differenceMs < 0) {
    customErrorMessage("Selected time is less than previous supine time. Previous Supine Time: " + formatDateTimeFromMs(supineTimeMs))
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger(differenceInMins)
if(differenceInMins >= difference ){
    return true;
}

customErrorMessage("Difference is out of range: " + Math.abs(differenceInMins) + ", Previous Supine Time: " + formatDateTimeFromMs(supineTimeMs));
return false;

function formatDateTimeFromMs(ms) {
    if (!ms) return "";

    var d = new Date(ms);

    var day = ("0" + d.getDate()).slice(-2);
    var monthIndex = d.getMonth();
    var year = d.getFullYear();

    var hour = ("0" + d.getHours()).slice(-2);
    var minute = ("0" + d.getMinutes()).slice(-2);
    var second = ("0" + d.getSeconds()).slice(-2);

    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    return day + " " + months[monthIndex] + " " + year + " "
           + hour + ":" + minute + ":" + second;
}

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item;
        }
    }
    return null;
}
