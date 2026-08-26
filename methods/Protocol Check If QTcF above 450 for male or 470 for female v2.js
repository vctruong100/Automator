/* jshint strict: false */

// Version: v2
// Description: Applies sex-specific QTcF protocol thresholds using keyword-based QTcF detection instead of exact item names. Pulls the latest matching QTcF/Fridericia value from the current form and returns the configured Yes/No codelist value.

var item = itemJson.item;

var maleRange = 450;
var femaleRange = 470;

var sexMale = formJson.form.subject.volunteer.sexMale;

function normalizeName(value) {
    if (value == null) return "";
    return value.toString().toUpperCase().replace(/\s+/g, " ");
}

function containsValue(input, keyword) {
    if (input == null) return false;
    return input.toString().toLowerCase().indexOf(keyword.toLowerCase()) !== -1;
}

function isQtcfItem(itemName) {
    var name = normalizeName(itemName);
    return containsValue(name, "QTCF") ||
        (containsValue(name, "FRIDERICIA") && containsValue(name, "QTC")) ||
        containsValue(name, "QT CORRECTED BY FRIDERICIA");
}

function addNumericValue(list, sourceItem) {
    if (!sourceItem || sourceItem.canceled || sourceItem.value === null || sourceItem.value === undefined || sourceItem.value === "") return;

    var numericValue = Number(sourceItem.value);
    if (!isNaN(numericValue)) {
        list.push({
            name: sourceItem.name,
            value: numericValue
        });
    }
}

function pullLatestQtcfFromForm(formJsonValue) {
    var itemGroups = formJsonValue.form.itemGroups;
    var group, items, groupItem, i, j;
    var matches = [];

    if (!itemGroups || itemGroups.length < 1) return null;

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !group.items) continue;

        items = group.items;
        for (j = 0; j < items.length; j++) {
            groupItem = items[j];
            if (!groupItem) continue;
            if (isQtcfItem(groupItem.name)) addNumericValue(matches, groupItem);
        }
    }

    if (matches.length < 1) return null;

    var latest = matches[matches.length - 1];
    logger("QTcF keyword matches: " + matches.map(function(match) { return match.name + "=" + match.value; }).join(", "));
    logger("Latest QTcF value used: " + latest.value + " from " + latest.name);
    return latest.value;
}

try {
    var qtcf = pullLatestQtcfFromForm(formJson);

    logger("Is it male: " + sexMale);
    logger("QTcF value: " + qtcf);

    if (qtcf === null || qtcf === undefined || isNaN(qtcf)) return null;
    if ((sexMale && qtcf > maleRange) || (!sexMale && qtcf > femaleRange)) return item.codeListItems[1].codedValue; // yes

    return item.codeListItems[0].codedValue; // no
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
