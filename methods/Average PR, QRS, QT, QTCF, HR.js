/* jshint strict: false */

// Version: v1
// Purpose: Extended average calculator with wider item name coverage.

// Add item names
// items to pull for calculation
var PRitem = ["PR", "PR V4", "🟢 PR (100-250) ms",];
var QRSitem = ["QRS", "QRS V4", "🟢 QRS (50-150)",];
var Qtitem = ["QT", "QT V4", "🟢 QT (250-500) ms",];
var QTcFitem = ["QTcF", "QTcF V4", "🟢 QTcF (250 - 500) ms",];
var HRitem = ["RATE", "RATE V4", "🟢 HR (30 - 200)",];
var QtcBitem = ["🟢 QTcB (I: >470)",]
var RRitem = ["🟢 RR (I: 120-200)"];

// items to attach
var PRattach = ["avg_PR", "PR AVERAGE (100-250)",];
var QRSattach = ["avg_QRS", "QRS AVERAGE (50-150)",];
var QTcFattach = ["avg_QTcf", "QTcF AVERAGE (250-500)",];
var QTattach = ["Avg_QT", "QT AVERAGE (250-500)",];
var HRattach = ["HR", "Avg_HR", "HR AVERAGE (30- 200)",];
var QtcBattach = ["QTcB AVERAGE (I: >470)"]
var RRattach = ["RR AVERAGE (I: 120-200)"];

var form = formJson.form;
var item = itemJson.item;

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
function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function calculateAverage(values) {
    if (values.length === 0) return null;
    var sum = 0;
    for (var i = 0; i < values.length; i++) {
        sum += values[i];
    }
    return Math.round(sum / values.length).toString().split('.')[0];
}

function getItemGroupName(form) {
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

function populateList(form, targetItem, attachedItem, repeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;

    if (repeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = group.items.length - 1; j >= 0; j--) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name) && list.length > 1) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    logger(item.value);
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
                }
            }
        }
    }
    else {
        for (i = 0; i < itemGroups.length; i++) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = 0; j < group.items.length; j++) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name)) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseInt(item.value));
                    }
                }
            }
        }
    }
    return list;
}

try {
    var isRepeat = false;
    var groupName = getItemGroupName(formJson);
    logger("Group Name: " + groupName)
    if (groupName) {
        isRepeat = containsValue(groupName, "repeat")
    }

    var PRlist = populateList(formJson, PRitem, PRattach, isRepeat);
    var QRSlist = populateList(formJson, QRSitem, QRSattach, isRepeat);
    var Qtlist = populateList(formJson, Qtitem, QTattach, isRepeat);
    var QTcFlist = populateList(formJson, QTcFitem, QTcFattach, isRepeat);
    var HRlist = populateList(formJson, HRitem, HRattach, isRepeat);
    var QTcBlist = populateList(formJson, QtcBitem, QtcBattach, isRepeat);
    var RRlist = populateList(formJson, RRitem, RRattach, isRepeat);

    logger("Is it a repeat? " + isRepeat);
    logger("PRlist: " + PRlist);
    logger("QRSlist: " + QRSlist);
    logger("Qtlist: " + Qtlist);
    logger("QTcFlist: " + QTcFlist);
    logger("HRlist: " + HRlist);
    logger("QTcBlist: " + QTcBlist);
    logger("RRlist: " + RRlist);

    var avgPR = calculateAverage(PRlist);
    var avgQRS = calculateAverage(QRSlist);
    var avgQT = calculateAverage(Qtlist);
    var avgQTcF = calculateAverage(QTcFlist);
    var avgHR = calculateAverage(HRlist);
    var avgQtcB = calculateAverage(QTcBlist);
    var avgRR = calculateAverage(RRlist);

    logger("Average PR: " + avgPR);
    logger("AVerage QRS: " + avgQRS);
    logger("Average QT: " + avgQT);
    logger("Average QtcF: " + avgQTcF);
    logger("Average HR: " + avgHR);
    logger("Average QtcB: " + avgQtcB);
    logger("Average RR: " + avgRR);

    if (containsItemName(PRattach, item.name)) return avgPR;
    if (containsItemName(QRSattach, item.name)) return avgQRS;
    if (containsItemName(QTcFattach, item.name)) return avgQTcF;
    if (containsItemName(QTattach, item.name)) return avgQT;
    if (containsItemName(HRattach, item.name)) return avgHR;
    if (containsItemName(QtcBattach, item.name)) return avgQtcB;
    if (containsItemName(RRattach, item.name)) return avgRR;

    return null;
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
