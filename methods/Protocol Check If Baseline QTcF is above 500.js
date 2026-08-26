/* jshint strict: false */

// Version: v3
// Description: Calculates median ECG or vital values from the attached median item, auto-detecting HR, RR, PR, QRS, QT, QTC, QTcF, SYS, or DIA source values across repeat and non-repeat item groups and logging collected values.

var form = formJson.form;
var attachedItem = itemJson.item;
var sigfig = attachedItem.significantDigits;

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

function matchesMetric(itemName, metric) {
    var name = normalizeName(itemName);
    if (metric === "QTCF") return name.indexOf("QTCF") !== -1 || containsStandaloneKeyword(name, "QTCF");

    return false;
}

function pullItemOnKeyword(formJsonValue, metric) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;

        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (matchesMetric(groupItem.name, metric)) {
                logger(metric + " matched item: " + groupItem.name + " | Value: " + groupItem.value);
                return groupItem.value;
            }
        }
    }

    return null;
}

var collected = Number(pullItemOnKeyword(formJson, "QTCF"));
if (collected && collected > 500) return "Y"
else if (collected && collected <= 500) return "N";
return null;