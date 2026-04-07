/* jshint strict: false */ 
// Add item names here
const SBPitem = [
    "Systolic BP (SCR Protocol Range)", 
    "Systolic BP (Protocol Range)"
];
const DBPitem = [
    "Diastolic BP (SCR_Protocol Range)", 
    "Diastolic BP (Protocol Range)"
];
const HRitem = [
    "Heart Rate (Protocol range) v2.0",
    "Heart Rate (Protocol range)", 
];

const SBPdifferenceItem = ["Difference in SBP"];
const DBPdifferenceItem = ["Difference in DBP"];
const HRdifferenceItem = ["Difference in HR"];

// ======== Don't modify ========

const itemJsn = itemJson.item;
const currentItemName = itemJsn.name;

var supineSBP = pullItemFromForm(formJson, SBPitem, "supine");
var supineDBP = pullItemFromForm(formJson, DBPitem, "supine");
var supineHR = pullItemFromForm(formJson, HRitem, "supine");

var standingSBP = pullItemFromForm(formJson, SBPitem, "standing");
var standingDBP = pullItemFromForm(formJson, DBPitem, "standing");
var standingHR = pullItemFromForm(formJson, HRitem, "standing");

var ortho = null;
logger("Supine SBP: " + supineSBP + ", Standing SBP: " + standingSBP);
logger("Supine DBP: " + supineDBP + ", Standing DBP: " + standingDBP);
logger("Supine HR: " + supineHR + ", Standing HR: " + standingHR);

if (SBPdifferenceItem.indexOf(currentItemName) !== -1) {
    if (!supineSBP || supineSBP == null || !standingSBP || standingSBP == null) return null;
    ortho = String(calculateDifference(supineSBP, standingSBP));
}
if (DBPdifferenceItem.indexOf(currentItemName) !== -1) {
    if (!supineDBP || supineDBP == null || !standingDBP || standingDBP == null) return null;
    ortho = String(calculateDifference(supineDBP, standingDBP));
}
if (HRdifferenceItem.indexOf(currentItemName) !== -1) {
    if (!supineHR || supineHR == null || !standingHR || standingHR == null) return null;
    ortho = String(calculateDifference(supineHR, standingHR));
}

logger("SBP Difference: " + String(calculateDifference(supineSBP, standingSBP)));
logger("DBP Difference: " + String(calculateDifference(supineDBP, standingDBP)));
logger("HR Difference: "  + String(calculateDifference(supineHR, standingHR)));

return ortho;

function pullItemFromForm(form, targetItem, keyword) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        groupName = group.name.toString().toLowerCase();
        if (!group || group.canceled || groupName.indexOf(keyword) == -1) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && item.value !== null && !item.canceled && item.value !== "") {
                if (matchesAny(targetItem, item.name)) return item.value;
            }
        }
    }
    return null;
}

function calculateDifference(supine, standing) {
    return parseInt(supine, 10) - parseInt(standing, 10);
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function normalizeNameStrict(s) {
    if (s === undefined || s === null) {
        return "";
    }
    var t = String(s);
    t = t.normalize ? t.normalize('NFKC') : t;
    t = t.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ');
    t = t.replace(/\s+/g, ' ');
    t = t.trim().toLowerCase();
    return t;
}
function matchesAny(targetNames, name) {
    if (!targetNames || targetNames.length < 1) {
        return false;
    }
    if (!name) {
        return false;
    }
    var normalizedName = normalizeNameStrict(name);
    for (var i = 0; i < targetNames.length; i++) {
        var candidate = normalizeNameStrict(targetNames[i]);
        if (candidate === normalizedName) {
            return true;
        }
    }
    return false;
}
