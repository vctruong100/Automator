const sysItem = [
    "SYS (I: 90-150)",
    "Systolic Blood Pressure",
]
const diaItem = [
    "DIA (I: 50-99)",
    "DIA (I: 50-90)"
]

const hrItem = [
    "HR (I: 45-100)"
]

const sysStandingItem = [
    "SYS (I: 90-145) standing",
    "SYS (I: 90-150) standing",
]

const diaStandingItem = [
    "DIA (I: 50-90) standing",
    "DIA (I: 50-99) standing",
]

const hrStandingItem = [
    "HR (I: 45-90) standing",
    "HR (I: 45-100) standing",
]

const sysAttachedItem = [
    "Systolic_BP_Diff",
    "ortho sbp check",
]

const diaAttachedItem = [
    "Diastolic_BP_Diff",
    "Repeat_ortho dbp check",
]

const hrAttachedItem = [
    "HR_Diff",
    "Repeat_ortho hr check",
]

const sysRange = 20;
const diaRange = 10;
const hrRange = 30;

var sysSemi = null;
var sysStanding = null;

var diaSemi = null;
var diaStanding = null;

var hrSemi = null;
var hrStanding = null;

const item = itemJson.item;

logger("Attached item: " + item.name);
if (sysAttachedItem.indexOf(item.name) !== -1) {
    sysSemi = pullLastItemFromForm(formJson, sysItem);
    sysStanding = pullFirstItemFromForm(formJson, sysStandingItem);
    return checkDifference(sysRange, sysSemi, sysStanding);
} else if (diaAttachedItem.indexOf(item.name) !== -1) {
    diaSemi = pullLastItemFromForm(formJson, diaItem);
    diaStanding = pullFirstItemFromForm(formJson, diaStandingItem);
    return checkDifference(diaRange, diaSemi, diaStanding);
} else if (hrAttachedItem.indexOf(item.name) !== -1) {
    hrSemi = pullLastItemFromForm(formJson, hrItem);
    hrStanding = pullFirstItemFromForm(formJson, hrStandingItem);
    return checkDifference(hrRange, hrSemi, hrStanding);
}

return null;

function checkDifference(range, semi, standing) {
    logger("Semi: " + semi);
    logger("Standing: " + standing);
    var diff = parseInt(semi) - parseInt(standing);
    logger("Difference: " + diff);
    if (diff > range) return "YES, decrease by " + diff + " mmHg.";
    else if (diff <= range && diff >= 0) return "NO, decrease by " + diff + " mmHg.";
    else if (diff < 0) return "NO, increase by " + diff + " mmHg.";
}

function pullFirstItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}

function pullLastItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}