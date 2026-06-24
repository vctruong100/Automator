/* jshint strict: false */

// Version: v1
// Purpose: Computes differences between supine and standing vitals (different item names).

// Add item names
var sysItem = ["SYS (I: 90-150)", "Systolic Blood Pressure",]
var diaItem = ["DIA (I: 50-99)", "DIA (I: 50-90)"]
var hrItem = ["HR (I: 45-100)"]

var sysStandingItem = ["SYS (I: 90-145) standing", "SYS (I: 90-150) standing",]
var diaStandingItem = ["DIA (I: 50-90) standing","DIA (I: 50-99) standing",]
var hrStandingItem = ["HR (I: 45-90) standing", "HR (I: 45-100) standing",]

var sysAttachedItem = ["Systolic_BP_Diff", "ortho sbp check",]
var diaAttachedItem = ["Diastolic_BP_Diff","Repeat_ortho dbp check",]

var hrAttachedItem = ["HR_Diff", "Repeat_ortho hr check",]

var sysDifferenceRange = 20;
var diaDifferenceRange = 10;
var hrDifferenceRange = 30;

var semi = null;
var standing = null;

var item = itemJson.item;

function checkDifference(range, semi, standing) {
    logger("Semi: " + semi);
    logger("Standing: " + standing);
    var diff = parseInt(semi) - parseInt(standing);
    logger("Difference: " + diff);
    if (diff > range) return "YES, decrease by " + diff + " mmHg.";
    else if (diff <= range && diff >= 0) return "NO, decrease by " + diff + " mmHg.";
    else if (diff < 0) return "NO, increase by " + diff + " mmHg.";
}

function pullItemFromForm(form, targetItem, lastToFirst) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

	if (!itemGroups || itemGroups.length < 1) return null;

    if (lastToFirst) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
            }
        }
    } else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
            }
        }
    }
    return null;
}

try {
    logger("Attached item: " + item.name);
    if (sysAttachedItem.indexOf(item.name) !== -1) {
        semi = pullItemFromForm(formJson, sysItem, true);
        standing = pullItemFromForm(formJson, sysStandingItem, false);
        return checkDifference(sysDifferenceRange, semi, standing);
    } else if (diaAttachedItem.indexOf(item.name) !== -1) {
        semi = pullItemFromForm(formJson, diaItem, true);
        standing = pullItemFromForm(formJson, diaStandingItem, false);
        return checkDifference(diaDifferenceRange, semi, standing);
    } else if (hrAttachedItem.indexOf(item.name) !== -1) {
        semi = pullItemFromForm(formJson, hrItem, true);
        standing = pullItemFromForm(formJson, hrStandingItem, false);
        return checkDifference(hrDifferenceRange, semi, standing);
    }

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
