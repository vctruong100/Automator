/* jshint strict: false */\

// Add item names
// items to pull for calculation
const PRitem = ["PR", "PR V4"];
const QRSitem = ["QRS", "QRS V4"];
const Qtitem = ["QT", "QT V4"];
const QTcFitem = ["QTcF", "QTcF V4"];
const HRitem = ["RATE", "RATE V4"];

// items to attach
const PRattach = ["avg_PR"];
const QRSattach = ["avg_QRS"];
const QTcFattach = ["avg_QTcf"];
const QTattach = ["Avg_QT"];
const HRattach = ["HR", "Avg_HR"];

// ======== Don't modify ========
var form = formJson.form;
var item = itemJson.item;

var PRlist = [];
var QRSlist = [];
var Qtlist = [];
var QTcFlist = [];
var HRlist = [];

var PRcount = 0;
var QRScount = 0;
var QtCount = 0;
var QTcFcount = 0;
var HRcount = 0;

var isRepeat = false;
var groupName = getItemGroupName(formJson);
logger("Group Name: " + groupName)
if (groupName) {
    isRepeat = containsValue(groupName, "repeat")
}

if (isRepeat) {
    PRlist = populateListLastToFirst(formJson, PRitem, PRattach);
    QRSlist = populateListLastToFirst(formJson, QRSitem, QRSattach);
    Qtlist = populateListLastToFirst(formJson, Qtitem, QTattach);
    QTcFlist = populateListLastToFirst(formJson, QTcFitem, QTcFattach);
    HRlist = populateListLastToFirst(formJson, HRitem, HRattach);
}
else {
    PRlist = populateList(formJson, PRitem, PRattach);
    QRSlist = populateList(formJson, QRSitem, QRSattach);
    Qtlist = populateList(formJson, Qtitem, QTattach);
    QTcFlist = populateList(formJson, QTcFitem, QTcFattach);
    HRlist = populateList(formJson, HRitem, HRattach);
}


logger("Is it a repeat? " + isRepeat);
logger("PRlist: " + PRlist);
logger("QRSlist: " + QRSlist);
logger("Qtlist: " + Qtlist);
logger("QTcFlist: " + QTcFlist);
logger("HRlist: " + HRlist);

var avgPR = calculateAverage(PRlist);
var avgQRS = calculateAverage(QRSlist);
var avgQT = calculateAverage(Qtlist);
var avgQTcF = calculateAverage(QTcFlist);
var avgHR = calculateAverage(HRlist);

logger("Average PR: " + avgPR);
logger("AVerage QRS: " + avgQRS);
logger("Average QT: " + avgQT);
logger("Average QtcF: " + avgQTcF);
logger("Average HR: " + avgHR);

if (PRattach.indexOf(item.name) !== -1) {
    return avgPR;
}
if (QRSattach.indexOf(item.name) !== -1) {
    return avgQRS;
}
if (QTcFattach.indexOf(item.name) !== -1) {
    return avgQTcF;
}
if (QTattach.indexOf(item.name) !== -1) {
    return avgQT;
}
if (HRattach.indexOf(item.name) !== -1) {
    return avgHR;
}

return null;

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

function populateList(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                }
            }
        }
    }
    return list;
}

function populateListLastToFirst(form, targetItem, attachedItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = []
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = group.items.length - 1; j >= 0; j--) {
            item = group.items[j];
            if (attachedItem.indexOf(item.name) !== -1 && list.length > 1) return list;
            if (item && targetItem.indexOf(item.name) !== -1) {
                logger(item.value);
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseInt(item.value));
                }
            }
        }
    }
    return list;
}