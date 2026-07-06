/* jshint strict: false */

// Version: v1
// Purpose: Computes median values for full ECG/vitals panel (non-repeat).
var HRitem = ["HR (P: <60) 🟢", "RATE V4", "HR (P: >60) 🟢", "HR (P: <60)", "HR (P: <60)", "HR (P: >60) ms", "HR (P: <60) ms"];
var PRitem = ["PR (P: <220) ms 🟢", "PR V4", "PR (P: <220)  ms 🟢", "PR (P: <220)  ms", "PR (P: <220)  ms 🟢"];
var QTcFitem = ['QTc (≤ 450 msec "Males") (≤ 470 msec "Females") 🟢', "QTcF V4", 'QTc (≤ 450 msec "Males") (≤ 470 msec "Females")  🟢', 'QTc (≤ 450 msec "Males") (≤ 470 msec "Females")'];

var PRattach = ["PR AVERAGE (100-250)", "PR AVERAGE (<220)", "PR AVERAGE (100-250)", "PR AVERAGE (<220) ms"];
var HRattach = ["HR AVERAGE (30- 200)", "Avg_HR", "HR AVERAGE (<60)", "HR AVERAGE (<60) ", "HR AVERAGE (>60) ms", "HR AVERAGE (>60)"];
var QTcFattach = ["QTcF AVERAGE (250-500)", 'QTc AVERAGE (≤ 450 msec "Males") (≤ 470 msec "Females") 🟢', "QTcF AVERAGE  (250-500)", 'QTc AVERAGE (≤ 450 msec "Males") (≤ 470 msec "Females")'];

var item = itemJson.item;
var sigfig = item.significantDigits;
var isMale = formJson.form.subject.volunteer.sexMale;

var QTcFlist = [];

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
function calculateMedian(values, sigfig) {
    if (values.length === 0) return null;
    values.sort(function(a, b) { return a - b; });

    var mid = Math.floor(values.length / 2);
    var median;

    if (values.length % 2 === 0) {
        median = (values[mid - 1] + values[mid]) / 2;
    } else {
        median = values[mid];
    }

    var factor = Math.pow(10, sigfig);
    return Math.round(median * factor) / factor;
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}


function populateList(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    var list = [];
	if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            if (containsItemName(attachedItem, item.name) && list.length > 1) return list;
            if (item && containsItemName(targetItem, item.name)) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                }
            }
        }
    }

    return list;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

try {
    logger("Attached item: " + item.name);
    var HRList = populateList(formJson, HRitem, HRattach);
    var PRList = populateList(formJson, PRitem, PRattach);
    var QtcFList = populateList(formJson, QTcFitem, QTcFattach);

    var hrMedian = calculateMedian(HRList, sigfig);
    var prMedian = calculateMedian(PRList, sigfig);
    var qtcfAverage = calculateAverage(QtcFList);
    
    logger("HR list: " + HRList);
    logger("PR list: " + PRList);
    logger("QTcF list: " + QtcFList);
    logger("HR median: " + hrMedian);
    logger("PR median: " + prMedian);
    logger("qtcfAverage: " + qtcfAverage);
    
    if (containsValue(item.name, "inc001")) {
        if (prMedian < 220 && hrMedian > 60) return "YES";
        return "NO, SF";
    }
    else if (containsValue(item.name, "exc022")) {
        if ((isMale && qtcfAverage > 470) || (!isMale && qtcfAverage > 480)) return "YES, SF";
        else return "NO";
    }
    return null;

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}