/* jshint strict: false */

// Version: v2
// Description: Applies comprehensive range checks for standard vital-sign and ECG parameters using keyword-based metric detection instead of exact item names. Detects HR, SYS, DIA, PR, RR, QTcF, QTC, QRS, and QT while avoiding ECG keyword collisions such as QT matching QTC/QTCF.

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

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;

    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function containsStandaloneKeyword(input, keyword) {
    var value = normalizeName(input);
    var target = normalizeName(keyword);
    var startIndex = 0;
    var index;
    var before;
    var after;

    while (startIndex < value.length) {
        index = value.indexOf(target, startIndex);
        if (index === -1) return false;

        before = index === 0 ? "" : value.charAt(index - 1);
        after = index + target.length >= value.length ? "" : value.charAt(index + target.length);

        if ((before === "" || !/[A-Z0-9]/.test(before)) && (after === "" || !/[A-Z0-9]/.test(after))) return true;

        startIndex = index + target.length;
    }

    return false;
}

function isQtcfItem(itemName) {
    var name = normalizeName(itemName);
    return containsValue(name, "QTCF") ||
        (containsValue(name, "FRIDERICIA") && containsValue(name, "QTC")) ||
        containsValue(name, "QT CORRECTED BY FRIDERICIA");
}

function getMetricFromItemName(itemName) {
    var name = normalizeName(itemName);

    if (isQtcfItem(name)) return "QTCF";
    if (containsValue(name, "QTC") && !containsValue(name, "QTCF")) return "QTC";
    if (containsStandaloneKeyword(name, "QRS")) return "QRS";
    if (containsStandaloneKeyword(name, "PR")) return "PR";
    if (containsStandaloneKeyword(name, "QT") && !containsValue(name, "QTC")) return "QT";
    if (containsStandaloneKeyword(name, "SYS") || containsValue(name, "SYSTOLIC")) return "SYS";
    if (containsStandaloneKeyword(name, "DIA") || containsValue(name, "DIASTOLIC")) return "DIA";
    if (containsStandaloneKeyword(name, "HR") || containsValue(name, "HEART RATE") || containsValue(name, "PULSE")) return "HR";
    if (containsStandaloneKeyword(name, "RR") || containsValue(name, "RESPIRATORY RATE")) return "RR";

    return null;
}

function getRangeForMetric(metric) {
    if (metric === "SYS") return SysRange;
    if (metric === "DIA") return DiaRange;
    if (metric === "HR") return HrRange;
    if (metric === "RR") return RrRange;
    if (metric === "PR") return PrRange;
    if (metric === "QRS") return QrsRange;
    if (metric === "QTC") return QtcRange;
    if (metric === "QTCF") return isMale ? QtcFRangeMale : QtcFRangeFemale;
    if (metric === "QT") return QtRange;

    return null;
}

function checkRange(range, isRepeat, metric) {
    var min = range[0];
    var max = range[1];
    var value = Number(item.value);

    logger("Metric: " + metric);
    logger("Item name: " + item.name);
    logger("Value: " + item.value);
    logger("Min-Max: " + min + "-" + max);

    if (item.value === null || item.value === undefined || item.value === "" || isNaN(value)) return null;

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
    var metric = getMetricFromItemName(item.name);
    var range = getRangeForMetric(metric);

    logger("Matched metric: " + metric);
    if (!metric || !range) return false;

    return checkRange(range, isRepeat, metric);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
