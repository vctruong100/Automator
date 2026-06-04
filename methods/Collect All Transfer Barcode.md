const formsList = [
    "(*)LAB_🩸Plasma Acetaminophen PREDOSE - 1 TUBE",
    "(*)LAB_🩸D-1_15min, 30min, 60min, 2h, 3h, 4h, 6h, 10h_Plasma Acetaminophen - 1 TUBE",
    "(*)LAB_🩸D84 | 12h MK-4082 PK | 10h PL Acetaminophen - 2 TUBES ",
    "(*)LAB_🩸D84 PL Acetaminophen (0.5h)  - 1 TUBE",
];
const itemNames = [
    "🟣Plasma Acetaminophen (PREDOSE)",
    "🟣Plama Acetaminophen (PREDOSE)",
    "🟣Plasma Acetaminophen (0.5H)",
    "🟣Plasma Acetaminophen (10H)",
];

var allMappings = [];

for (var i = 0; i < formsList.length; i++) {
    var formData = findFormDataAcrossStudyEvents(formsList[i], false);
    var result = getItemMapping(formData);

    for (var j = 0; j < result.length; j++) {
        allMappings.push(result[j]);
    }
}

return allMappings.join(' | ');

function getItemMapping(formData) {
    var mapping = [];
    logger('Form data length: ' + formData.length);

    for (var i = 0; i < formData.length; i++) {
        var form = formData[i].form;
        var studyevent = form.studyEventName;
        logger('Inspecting form: ' + form.name);
        if (form.canceled) continue;
        if (form.itemGroups && form.itemGroups.length > 0) {
            for (var j = 0; j < form.itemGroups.length; j++) {
                var itemGroup = form.itemGroups[j];

                if (itemGroup.items && itemGroup.items.length > 0) {
                    for (var k = 0; k < itemGroup.items.length; k++) {
                        var item = itemGroup.items[k];

                        if (itemNames.indexOf(item.name) !== -1) {
                            var contextRaw = getItemDataContextByItemDataId(item.id);
                            var context = JSON.parse(contextRaw);
                            if (typeof context === 'string') {
                                logger('Barcode lookup failed for item.id ' + item.id + ': ' + context);
                                continue;
                            }
                            logger(JSON.stringify(context, null, 2));
                            var transferBarcode = context.transferBarcodes || '';
                            var formatted = studyevent + ' - ' + item.name + ': ' + transferBarcode;
                            mapping.push(formatted);
                            logger('Pushed value: ' + formatted);
                        }
                    }
                }
            }
        } else {
            logger('No item groups found for form: ' + form.name);
        }
    }

    return mapping;
}