const item = itemJson.item;

const qtcfItems = [
    "QTcF (P: < 450 (male), < 470 (female))",
]

const yesCodeList = item.codeListItems[1].codedValue;
const noCodeList = item.codeListItems[0].codedValue;
const maleRange = 450;
const femaleRange = 470;

const sexMale = formJson.form.subject.volunteer.sexMale;
var qtcf = pullItemFromForm(formJson, qtcfItems);

logger("Is it male: " + sexMale);
logger("Qtcf value: " + item.value);
if ((sexMale && item.value > maleRange) || (!sexMale && item.value > femaleRange)) return yesCodeList;

return noCodeList;

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