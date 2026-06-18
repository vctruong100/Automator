/* jshint strict: false */

// Version: v1
// Purpose: Comprehensive range check for all standard vital/ECG parameters.


// Add item names
var SYS = [
    "SYS (P: 90-140)", "SYS_3", "SYS_1", "SYS_2"
];
var DIA = [
    "DIA (P: 60-90)", "DIA_3", "DIA_1", "DIA_2"
];
var HR = [
    "HR (P: 40-100)", "HR_3", "HR_1", "HR_2"
];
var RR = [
    "RR (P: 600-1000)", "RR_3", "RR_1", "RR_2"
];
var PR = [
    "PR (P: 120-200)", "PR_3", "PR_1", "PR_2", "PR (I: 120 - 200 ms)",
];
var QRS = [
    "QRS (P: 90-140)", "QRS_3", "QRS_1", "QRS_2"
];
var QTC = [
    "QTC (P: 90-140)", "QTC_3", "QTC_1", "QTC_2"
];
var QtcF = [
    "QtcF (P: 90-140)", "QtcF_3", "QtcF_1", "QtcF_2"
];
var QT = [
    "QT (P: 90-140)", "QT_3", "QT_1", "QT_2"
];

var SysRange = [90, 140]; // Minimum, Maximum
var DiaRange = [60, 90];
var HrRange = [40, 100];
var RrRange = [600, 1000];
var PrRange = [120, 200];
var QrsRange = [100, 130];
var QtcRange = [380, 500];
var QtcFRangeMale = [380, 450];
var QtcFRangeFemale = [380, 470];
var QtRange = [400, 500];

var item = itemJson.item;
var isMale = formJson.form.subject.volunteer.sexMale;
var groupName = getItemGroupName(formJson);
var isRepeat = containsValue(groupName, "repeat");

function checkRange(range, isRepeat) {
    const [min, max] = range;
    var value = item.value;
    logger("Value: " + value);
    logger("Min-Max: " + min + "-" + max)
    if (value >= min && value <= max) {
        return true;
    }
    if (isRepeat) {
        customErrorMessage(RepeatErrorMsg);
        return false;
    }
    customErrorMessage(errorMsg);
    return false;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function getItemGroupName(form) {
    var groupName = "";
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                groupName = group.name;
                return groupName;
            }
        }
        if (groupName) break;
    }
    return groupName;
}

try {
    if (SYS.indexOf(item.name) !== -1) return checkRange(SysRange, isRepeat);
    if (DIA.indexOf(item.name) !== -1) return checkRange(DiaRange, isRepeat);
    if (HR.indexOf(item.name) !== -1) return checkRange(HrRange, isRepeat);
    if (RR.indexOf(item.name) !== -1) return checkRange(RrRange, isRepeat);
    if (PR.indexOf(item.name) !== -1) return checkRange(PrRange, isRepeat);
    if (QRS.indexOf(item.name) !== -1) return checkRange(QrsRange, isRepeat);
    if (QTC.indexOf(item.name) !== -1) return checkRange(QtcRange, isRepeat);
    if (QtcF.indexOf(item.name) !== -1) return checkRange(isMale ? QtcFRangeMale : QtcFRangeFemale, isRepeat);
    if (QT.indexOf(item.name) !== -1) return checkRange(QtRange, isRepeat);

    return false;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
