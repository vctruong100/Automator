var item = itemJson.item;

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var normalizedInput = input.toString().toLowerCase();
    return normalizedInput.indexOf(keyword) !== -1;
}

function getFrequencyTokens(value) {
    if (value == null) return [];
    return value.toString().toLowerCase().split(/[^a-z0-9]+/).filter(function(t) { return t.length > 0; });
}

function tokenMatches(tokens, keyword, exact) {
    for (var i = 0; i < tokens.length; i++) {
        if (exact) {
            if (tokens[i] === keyword) return true;
        } else {
            if (tokens[i].indexOf(keyword) === 0) return true;
        }
    }
    return false;
}

function getFrequencyRank(value) {
    if (value == null) return 999;

    var tokens = getFrequencyTokens(value);

    if (tokenMatches(tokens, "day") || tokenMatches(tokens, "daily") || tokenMatches(tokens, "d", true)) return 1;
    if (tokenMatches(tokens, "week") || tokenMatches(tokens, "weekly") || tokenMatches(tokens, "wk") || tokenMatches(tokens, "w", true)) return 2;
    if (tokenMatches(tokens, "month") || tokenMatches(tokens, "monthly") || tokenMatches(tokens, "mon", true) || tokenMatches(tokens, "m", true)) return 3;
    if (tokenMatches(tokens, "year") || tokenMatches(tokens, "yearly") || tokenMatches(tokens, "yr") || tokenMatches(tokens, "annu") || tokenMatches(tokens, "y", true)) return 4;

    return 999;
}

function getFrequencyName(rank) {
    if (rank == 1) return "Day";
    if (rank == 2) return "Week";
    if (rank == 3) return "Month";
    if (rank == 4) return "Year";
    return "-";
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

function padTwo(n) {
    var s = n.toString();
    return s.length < 2 ? "0" + s : s;
}

function getSortableDate(dateString) {
    if (!dateString) return null;

    var s = dateString.toString().replace(/[\s\-\/]/g, "").toLowerCase();

    // Year only
    if (s.length === 4 && /^\d{4}$/.test(s)) {
        return Number(s + "0101");
    }

    // DDMonYYYY, e.g. 18jan1976 or 18-Jan-1976
    var m1 = s.match(/^(\d{1,2})([a-z]{3})(\d{4})$/);
    if (m1) {
        var month = getMonthNumber(m1[2]);
        if (month) {
            return Number(m1[3] + month + padTwo(parseInt(m1[1], 10)));
        }
    }

    // YYYYMMDD or YYYY-MM-DD or YYYY/MM/DD
    var m2 = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m2) {
        return Number(m2[1] + m2[2] + m2[3]);
    }

    return null;
}

function getAlcoholEntries(formJsonValue, attachedItemId, groupName) {
    var itemGroups = formJsonValue.form.itemGroups;
    var entries = [];
    if (!itemGroups || itemGroups.length < 1) return entries;

    for (var i = 0; i < itemGroups.length; i++) {
        var group = itemGroups[i];
        if (!group || group.canceled || !containsValue(group.name, "alcohol") || group.name == groupName) continue;

        var entry = {
            amount: null,
            frequency: null,
            frequencyRank: 999,
            startDate: null,
            endDate: null,
            occurrence: false
        };
        var hasData = false;

        for (var j = 0; j < group.items.length; j++) {
            var it = group.items[j];
            if (!it || it.canceled || it.id == attachedItemId) continue;

            var val = it.value;
            if (val == null || val === "" || val === "-") continue;
            logger("Item name: " + it.name + ", item value: " + it.value)
            if (it.name === "SU_Amount") {
                var n = Number(val);
                if (!isNaN(n)) {
                    entry.amount = n;
                    hasData = true;
                }
            } else if (it.name === "SU_Frequency") {
                entry.frequency = val;
                entry.frequencyRank = getFrequencyRank(val);
                hasData = true;
            } else if (it.name === "SU_Start Date") {
                entry.startDate = val;
                hasData = true;
            } else if (it.name === "SU_End Date") {
                entry.endDate = val;
                hasData = true;
            } else if (it.name === "SU_Occurrence") {
                if (val.toString().toLowerCase() === "true") {
                    entry.occurrence = true;
                }
                hasData = true;
            }
        }

        if (hasData) {
            entries.push(entry);
        }
    }

    return entries;
}

function selectEntries(entries) {
    if (!entries || entries.length === 0) {
        return { selected: [], rank: 999, source: null };
    }

    var ongoing = [];
    var ended = [];
    var i;

    for (i = 0; i < entries.length; i++) {
        if (entries[i].endDate == null || entries[i].endDate === "") {
            ongoing.push(entries[i]);
        } else {
            ended.push(entries[i]);
        }
    }

    logger(ongoing.length)
    var pool = (ongoing.length > 0) ? ongoing : ended;
    var source = (ongoing.length > 0) ? "ongoing" : "ended";

    // If all entries have ended, narrow to the most recent end date
    logger(source)
    if (source === "ended") {
        var maxEndSortable = null;
        var hasValidEnd = false;

        for (i = 0; i < ended.length; i++) {
            var endSortable = getSortableDate(ended[i].endDate);
            if (endSortable != null && (maxEndSortable == null || endSortable > maxEndSortable)) {
                maxEndSortable = endSortable;
                hasValidEnd = true;
            }
        }
        logger("Has Valid End: " + hasValidEnd)
        if (hasValidEnd) {
            var recentEnded = [];
            for (i = 0; i < ended.length; i++) {
                var endSortable2 = getSortableDate(ended[i].endDate);
                if (endSortable2 != null && endSortable2 === maxEndSortable) {
                    recentEnded.push(ended[i]);
                }
            }
            if (recentEnded.length > 0) {
                pool = recentEnded;
            }
        }
    }

    // Choose most common frequency: Day > Week > Month > Year
    var bestRank = 999;
    for (i = 0; i < pool.length; i++) {
        if (pool[i].frequencyRank < bestRank) {
            bestRank = pool[i].frequencyRank;
        }
    }

    var selected = [];
    for (i = 0; i < pool.length; i++) {
        if (pool[i].frequencyRank === bestRank) {
            selected.push(pool[i]);
        }
    }

    return { selected: selected, rank: bestRank, source: source };
}

function combineAmount(entries) {
    var total = 0;
    var hasAmount = false;

    for (var i = 0; i < entries.length; i++) {
        if (entries[i].amount != null && !isNaN(entries[i].amount)) {
            total += entries[i].amount;
            hasAmount = true;
        }
    }

    if (!hasAmount) return "-";
    return total.toFixed(0);
}

function getEarliestStart(entries) {
    var earliestValue = null;
    var earliestSortable = null;

    for (var i = 0; i < entries.length; i++) {
        var sortable = getSortableDate(entries[i].startDate);
        if (sortable == null) continue;

        if (earliestSortable == null || sortable < earliestSortable) {
            earliestSortable = sortable;
            earliestValue = entries[i].startDate;
        }
    }

    if (earliestValue == null) return "-";
    return earliestValue;
}

function getLatestEnd(entries) {
    var latestValue = null;
    var latestSortable = null;

    for (var i = 0; i < entries.length; i++) {
        var sortable = getSortableDate(entries[i].endDate);
        if (sortable == null) continue;

        if (latestSortable == null || sortable > latestSortable) {
            latestSortable = sortable;
            latestValue = entries[i].endDate;
        }
    }

    if (latestValue == null) return "-";
    return latestValue;
}

function anyOccurrence(entries) {
    for (var i = 0; i < entries.length; i++) {
        if (entries[i].occurrence) return true;
    }
    return false;
}

function getFrequencyOutput(entries, rank) {
    if (rank <= 4) return getFrequencyName(rank);

    if (entries.length === 1 && entries[0].frequency != null) {
        return entries[0].frequency;
    }

    return "-";
}

try {
    var rawgroupName = getItemDataContextByItemDataId(item.id);
    var parsedGroupName = JSON.parse(rawgroupName).foundItemGroupName;
    logger("Group name: " + parsedGroupName);

    var entries = getAlcoholEntries(formJson, item.id, parsedGroupName);
    var selection = selectEntries(entries);
    var selected = selection.selected;
    var rank = selection.rank;

    if (item.name === "SU_Amount") {
        return combineAmount(selected);
    }

    if (item.name === "SU_Frequency") {
        return getFrequencyOutput(selected, rank);
    }

    if (item.name === "SU_Start Date") {
        return getEarliestStart(selected);
    }

    if (item.name === "SU_End Date") {
        if (selection.source === "ongoing") return "-";
        return getLatestEnd(selected);
    }

    if (item.name === "SU_Unit TEST" || item.name === "SU_Unit") {
        return anyOccurrence(selected) ? "DRINK" : "-";
    }

    if (item.name === "SU_Occurrence") {
        return anyOccurrence(selected) ? "True" : "False";
    }

    return null;
} catch (e) {
    logger("Error: " + e);
    return null;
}
