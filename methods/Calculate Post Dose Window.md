/* jshint strict: false */

// Version: v1
// Purpose: Calculates and returns the valid assessment window (30–45 minutes post-dose) based on the DateTime of IP administration within the same item group.

const itemName = [
    "DateTime of bronchodilator administration"    
]

var item = itemJson.item;

function pullItemFromForm(form, targetItem, groupName) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || group.name !== groupName) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}

function getTimeRange(datetimeStr) {
    if (!datetimeStr) {
        return "";
    }

    var parts = datetimeStr.split("T");
    if (parts.length < 2) {
        return "";
    }

    var timePart = parts[1]; 

    var timePieces = timePart.split(":");
    var hour = parseInt(timePieces[0], 10);
    var minute = parseInt(timePieces[1], 10);
    var second = parseInt(timePieces[2], 10);

    var totalSeconds = (hour * 3600) + (minute * 60) + second;

    // Add 30 minutes
    var minSeconds = totalSeconds + (30 * 60);

    // Add 45 minutes
    var maxSeconds = totalSeconds + (45 * 60);

    function formatTime(seconds) {
        var h = Math.floor(seconds / 3600);
        var m = Math.floor((seconds % 3600) / 60);
        var s = seconds % 60;

        if (h >= 24) {
            h = h % 24;
        }

        return (
            (h < 10 ? "0" : "") + h + ":" +
            (m < 10 ? "0" : "") + m + ":" +
            (s < 10 ? "0" : "") + s
        );
    }

    return formatTime(minSeconds) + " to " + formatTime(maxSeconds);
}

try {
    var rawItem = getItemDataContextByItemDataId(item.id);
    var context = JSON.parse(rawItem);
    var groupName = context.foundItemGroupName;
    logger("Group Name: " + groupName)
    var datetime = pullItemFromForm(formJson, itemName, groupName);
    logger("Datetime: " + datetime)
    if (!datetime || datetime == null || datetime == "") return null;
    return getTimeRange(datetime);
} catch (e) {
    customErrorMessage("Error: " + e);
    return null;
}

