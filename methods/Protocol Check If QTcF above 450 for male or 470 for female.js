/* jshint strict: false */

// Version: v1
// Purpose: Sex-specific QTcF protocol threshold check.

var item = itemJson.item;

var qtcfItems = [
    "QTcF (P: < 450 (male), < 470 (female))",
]

var maleRange = 450;
var femaleRange = 470;

var sexMale = formJson.form.subject.volunteer.sexMale;

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;

  if (!itemGroups || itemGroups.length < 1) return null;

    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null) return item.value;
        }
    }
    return null;
}

try {
    var qtcf = pullItemFromForm(formJson, qtcfItems);

    logger("Is it male: " + sexMale);
    logger("Qtcf value: " + item.value);
    if ((sexMale && item.value > maleRange) || (!sexMale && item.value > femaleRange)) return item.codeListItems[1].codedValue; // yes

    return item.codeListItems[0].codedValue; // no
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}
