
// Add item names
const SYS = [
    "SYS (P: 90-140)", "SYS_3", "SYS_1", "SYS_2"
];
const DIA = [
    "DIA (P: 60-90)", "DIA_3", "DIA_1", "DIA_2"
];
const HR = [
    "HR (P: 40-100)", "HR_3", "HR_1", "HR_2"
];
const RR = [
    "RR (P: 600-1000)", "RR_3", "RR_1", "RR_2"
];
const PR = [
    "PR (P: 120-200)", "PR_3", "PR_1", "PR_2", "PR (I: 120 - 200 ms)",
];
const QRS = [
    "QRS (P: 90-140)", "QRS_3", "QRS_1", "QRS_2"
];
const QTC = [
    "QTC (P: 90-140)", "QTC_3", "QTC_1", "QTC_2"
];
const QtcF = [
    "QtcF (P: 90-140)", "QtcF_3", "QtcF_1", "QtcF_2"
];
const QT = [
    "QT (P: 90-140)", "QT_3", "QT_1", "QT_2"
];

const SysRange = [90, 140]; // Minimum, Maximum
const DiaRange = [60, 90];
const HrRange = [40, 100];
const RrRange = [600, 1000];
const PrRange = [120, 200];
const QrsRange = [100, 130];
const QtcRange = [380, 500];
const QtcFRangeMale = [380, 450];
const QtcFRangeFemale = [380, 470];
const QtRange = [400, 500];

// ======== Don't modify ========
const item = itemJson.item;
const isMale = formJson.form.subject.volunteer.sexMale;
const groupName = getItemGroupName(formJson);
const isRepeat = containsValue(groupName, "repeat");

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

function checkRange(range, isRepeat) {
    const [min, max] = range;
    const value = item.value;
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