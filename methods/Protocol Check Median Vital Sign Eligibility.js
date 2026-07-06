/* jshint strict: false */

// Version: v1
// Purpose: Computes median values for full ECG/vitals panel (non-repeat).
// items to pull for calculation
var sysItem = ["SYS (60 - 200) mmHg", "SYS  (I: 90 - 150)", "SYS (60 - 200)", "SYS (P: 91 - 150)",];
var diaItem = ["DIA (40 - 110) mmHg", "DIA (I: 50 - 100)", "DIA (40 - 110)", "DIA (P:  51 - 100)", "DIA (P: 51 - 100)"];
var hrItem = ["HR (50 - 100) bpm", "HR (30 - 100)", "HR (P: 61 - 100)"];

// items to attach
var sysAttach = ["SYS MEAN AVERAGE", "SYS MEAN AVERAGE REPEAT", "SYS MEAN (P: 91 - 150)"];
var diaAttach = ["DIA MEAN AVERAGE", "DIA MEAN AVERAGE REPEAT", "DIA MEAN (P: 51 - 100)"];
var hrAttach = ["HR MEAN AVERAGE", "HR MEAN AVERAGE REPEAT", "HR MEAN (P: 61 - 100)"];

var EXC04Item = ["EXC004"]
var EXC05Item = ["EXC005"]

var isRepeatRequiredItem = ["3️⃣ ▶VS_Required repeat (Triplicate)"]

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

    return Math.round(median);
}

function populateList(form, targetItem, attachedItem, isRepeat) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j;
    var list = [];
	if (!itemGroups || itemGroups.length < 1) return null;

    if (isRepeat) {
        for (i = itemGroups.length - 1; i >= 0; i--) {
            group = itemGroups[i];
            if (!group || group.canceled) continue;
            for (j = group.items.length -1; j >= 0; j--) {
                item = group.items[j];
                if (containsItemName(attachedItem, item.name.trim()) && list.length > 1) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        logger("Item name: " + item.name + ", value: " + item.value);
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
                if (containsItemName(attachedItem, item.name.trim())) return list;
                if (item && containsItemName(targetItem, item.name)) {
                    if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                        list.push(parseFloat(item.value));
                        logger("Item name: " + item.name + ", value: " + item.value);
                    }
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

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "") {
                return item;
            }
        }
    }
    return null;
}

try {
    var isRepeat = false;
    var list = [];

    var repeatItem = pullItemFromForm(formJson, isRepeatRequiredItem);
    if (!repeatItem || repeatItem.value == null) return null;

    if (repeatItem.value == repeatItem.codeListItems[1].codedValue) isRepeat = true;
    
    logger("Attached item: " + item.name);
    var sysList = populateList(formJson, sysItem, sysAttach, isRepeat);
    var diaList = populateList(formJson, diaItem, diaAttach, isRepeat);
    var hrList = populateList(formJson, hrItem, hrAttach, isRepeat);

    logger("Sys list: "  + sysList);
    logger("Dia list: " + diaList);
    logger("Hr list: " + diaList);


    var sysMedian = calculateMedian(sysList);
    var diaMedian = calculateMedian(diaList);
    var hrMedian = calculateMedian(hrList);
    
    logger("Sys Median: " + sysMedian);
    logger("Dia Median: " + diaMedian);
    logger("Hr Median: " + hrMedian);
    
    if (containsItemName(EXC04Item, item.name)) {
        if (sysMedian  > 150 || diaMedian >= 100) return "YES, SF";
        return "NO";
    }
    if (containsItemName(EXC05Item, item.name)) {
        if (hrMedian > 100) return "YES, SF";
        return "NO";
    }
    
    if (sysMedian > 90 && hrMedian > 60 && diaMedian > 50) return "YES";
    else if (sysMedian <= 90 || hrMedian <= 60 || diaMedian <= 50) return "NO, SF";
    return null;

} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}