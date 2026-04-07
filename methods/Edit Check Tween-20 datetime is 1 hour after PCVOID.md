
// Add item names
const pcvoidItem = [
    "PCVOID",
    "PCVOID w/edit check",
]


// ======== Don't modify ========
const item = itemJson.item;
const difference = 60;
var groupid = null;

getItemGroupID(formJson.form);
var pcvoid = getItemValueFromSameGroup(formJson.form, pcvoidItem);
if (!pcvoid || pcvoid == null) return null;

var pcvoidMs = pcvoid.dateValueMs;
const collectedTimeMs = item.dateValueMs;

logger("pcvoidMs Time: "  + formatDateTime(pcvoid.value));
logger("collected time: " + formatDateTime(item.value));

const differenceMs = collectedTimeMs - pcvoidMs;
logger(differenceMs)
if (differenceMs < 0) {
    customErrorMessage("Selected time is less than previous PCVoid time. Previous PCVoid Time: " + formatDateTime(pcvoid.value))
    return false;
}
const differenceInMins = Math.abs(Math.floor(differenceMs / (1000 * 60)))

logger(differenceInMins)
if(differenceInMins >= difference){
    customErrorMessage("Selected time is more than previous PCVoid time. Previous PCVoid Time: " + formatDateTime(pcvoid.value))
    return false;
}

customErrorMessage("Difference is out of range: " + Math.abs(differenceInMins) + ", Previous PCVoid Time: " + formatDateTime(pcvoid.value));
return true;

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

function getItemGroupID(form) {
    for (var i = 0; i < form.itemGroups.length; i++) {
    var group = form.itemGroups[i];
    var items = group.items;
    if (!items || items.length < 1) continue;

    for (var j = 0; j < items.length; j++) {
        var it = items[j];
        if (it.id === item.id) {
            groupid = group.id;
            break;
        }
    }
    if (groupid) break;
    }
}

function getItemValueFromSameGroup(form, itemName) {
    var value = null;
    for (var i = 0; i < form.itemGroups.length; i++) {
        var group = form.itemGroups[i];
        if (group.id !== groupid) continue;

        var items = group.items;
        for (var j = 0; j < items.length; j++) {
            var item = items[j];
            if (item && itemName.indexOf(item.name)) {
                value = item.value;
                if (value && value !== null) return item;
            }
        }
    }
    return null; 
}