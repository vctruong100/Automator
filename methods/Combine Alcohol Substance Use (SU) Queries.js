var item = itemJson.item;

function pullItemFromForm(form, targetItem, itemid) {
    var itemGroups = form.form.itemGroups;
    var group, item, i, j;
    var list = [];
    if (!itemGroups || itemGroups.length < 1) return null;
    logger(targetItem)
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled || !containsValue(group.name, "alcohol")) continue;

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (containsItemName(targetItem, item.name) && item.value !== null && !item.canceled && item.value !== "" && item.id !== itemid) {
                logger("Item value:" + item.value)
                list.push(item.value);
            }
        }
    }
    return list;
}

function getAmountFrequencyPairs(form, itemid) {
    var itemGroups = form.form.itemGroups;
    var pairs = [];
    var group;
    var item;
    var i;
    var j;

    if (!itemGroups || itemGroups.length < 1) {
        return pairs;
    }

    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];

        if (!group || group.canceled || !containsValue(group.name, "alcohol")) {
            continue;
        }

        var amount = null;
        var frequency = null;

        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];

            if (item.canceled || item.id == itemid) {
                continue;
            }

            if (item.name == "SU_Amount" && item.value != null && item.value != "-") {
                amount = Number(item.value);
            }

            if (item.name == "SU_Frequency" && item.value != null && item.value != "-") {
                frequency = item.value;
            }
        }

        if (amount != null && frequency != null) {
            pairs.push({
                amount: amount,
                frequency: frequency
            });
        }
    }

    return pairs;
}

function getFrequencyMultiplier(fromRank, toRank) {
    if (fromRank == toRank) {
        return 1;
    }

    if (fromRank == 1 && toRank == 2) {
        return 7;
    }

    if (fromRank == 2 && toRank == 1) {
        return 1 / 7;
    }

    if (fromRank == 1 && toRank == 3) {
        return 30;
    }

    if (fromRank == 3 && toRank == 1) {
        return 1 / 30;
    }

    if (fromRank == 2 && toRank == 3) {
        return 4;
    }

    if (fromRank == 3 && toRank == 2) {
        return 1 / 4;
    }

    return 1;
}

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

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

function getFrequencyRank(value) {
    if (value == null) return 999;

    value = value.toString().toLowerCase();

    // Day
    if (value.indexOf("day") !== -1 ||
        value === "d") {
        return 1;
    }

    // Week
    if (value.indexOf("week") !== -1 ||
        value.indexOf("wk") !== -1 ||
        value === "w") {
        return 2;
    }

    // Month
    if (value.indexOf("month") !== -1 ||
        value.indexOf("mon") !== -1 ||
        value === "m") {
        return 3;
    }

    return 999;
}

function getFrequencyName(rank) {
    if (rank == 1) return "Day";
    if (rank == 2) return "Week";
    if (rank == 3) return "Month";
    return null;
}

function getFrequencyName(rank) {
    if (rank == 1) return "Day";
    if (rank == 2) return "Week";
    if (rank == 3) return "Month";
    return null;
}

function getMonthNumber(month) {
    month = month.toLowerCase();

    if (month == "jan") return "01";
    if (month == "feb") return "02";
    if (month == "mar") return "03";
    if (month == "apr") return "04";
    if (month == "may") return "05";
    if (month == "jun") return "06";
    if (month == "jul") return "07";
    if (month == "aug") return "08";
    if (month == "sep") return "09";
    if (month == "oct") return "10";
    if (month == "nov") return "11";
    if (month == "dec") return "12";

    return null;
}

function getSortableDate(dateString) {
    if (!dateString) return null;

    var value = dateString.toString().replace(/-/g, "");

    // Year only
    if (value.length == 4) {
        return Number(value + "0101");
    }

    // DDMonYYYY
    if (value.length != 9) return null;

    var day = value.substring(0, 2);
    var month = getMonthNumber(value.substring(2, 5));
    var year = value.substring(5, 9);

    if (month == null) return null;

    return Number(year + month + day);
}

try {
    var amount = null;
    var values = [];

    if (item.name == "SU_Amount") {

        var pairs = getAmountFrequencyPairs(formJson, item.id);
        if (pairs.length == 0) {
            return "-";
        }
        
        var targetRank = 0;
        
        for (var i = 0; i < pairs.length; i++) {
            var rank = getFrequencyRank(pairs[i].frequency);
        
            if (rank > targetRank) {
                targetRank = rank;
            }
        }

        var totalAmount = 0;
        for (var i = 0; i < pairs.length; i++) {
            var sourceRank = getFrequencyRank(pairs[i].frequency);
            totalAmount += pairs[i].amount * getFrequencyMultiplier(sourceRank, targetRank);
            logger("Total amount: " + totalAmount)
        }

        return totalAmount.toFixed(0);
    }
    if (item.name == "SU_Frequency") {
        values = pullItemFromForm(formJson, ["SU_Frequency"], item.id);
        var bestRank = 0;
        
        for (var i = 0; i < values.length; i++) {
            var rank = getFrequencyRank(values[i]);
        
            if (rank > bestRank) {
                bestRank = rank;
            }
        }
        logger(bestRank)
        var frequency = getFrequencyName(bestRank);
        logger(frequency)
        if (frequency == null) return "-";
        return frequency;
    }
    if (["SU_Unit TEST", "SU_Unit"].indexOf(item.name) !== -1) {
        values = pullItemFromForm(formJson, ["SU_Occurrence"], item.id);
        var occurred = false;
        if (containsValue(values, "true")) return "DRINK";
        return "-";
        // values = pullItemFromForm(formJson, ["SU_Unit TEST", "SU_Unit"])
        // var hasDrink = false;

        // for (var i = 0; i < values.length; i++) {
        //     if (values[i] != "-") {
        //         hasDrink = true;
        //         break;
        //     }
        // }
        
        // return hasDrink ? "DRINK" : "-";
    }
    if (item.name == "SU_Occurrence") {
        values = pullItemFromForm(formJson, ["SU_Occurrence"], item.id);
        var occurred = false;
        if (containsValue(values, "true")) return "True";
        return "False";
    }
    if (item.name == "SU_Start Date") {
        values = pullItemFromForm(formJson, ["SU_Start Date"], item.id);
    
        var earliestValue = null;
        var earliestSortable = null;
    
        for (var i = 0; i < values.length; i++) {
    
            var sortable = getSortableDate(values[i]);
    
            if (sortable == null) continue;
    
            if (earliestSortable == null || sortable < earliestSortable) {
                earliestSortable = sortable;
                earliestValue = values[i];
            }
        }
    
        if (earliestValue == null) return "-";
        return earliestValue;
    }
    if (item.name == "SU_End Date") {
        values = pullItemFromForm(formJson, ["SU_End Date"], item.id);
    
        var latestValue = null;
        var latestSortable = null;
    
        for (var i = 0; i < values.length; i++) {
    
            var sortable = getSortableDate(values[i]);
    
            if (sortable == null) continue;
    
            if (latestSortable == null || sortable > latestSortable) {
                latestSortable = sortable;
                latestValue = values[i];
            }
        }
    
        if (latestValue == null) return "-";
        return latestValue;
    }
    return null;
} catch (e) {
    logger("Error: " + e);
    return null;
}