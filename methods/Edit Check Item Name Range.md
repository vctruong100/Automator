const errorMsg = "OOR, Repeat"; // Custom Error Message
const RepeatErrorMsg = "OOR, INVESTIGATOR ASSESS"; // Custom Repeat Error Message

// =======================
var item = itemJson.item;
logger("Item name: " + item.name);
logger("Value: " + item.value);
var groupName = getItemGroupName(formJson);
var isRepeat = containsValue(groupName, "repeat");
var isQtcOrQTcF = containsValue(item.name, "qtcf") || containsValue(item.name, "qtc");
var isMale = formJson.form.subject.volunteer.sexMale;

logger("Is repeat: " + isRepeat + ", Is Male: " + isMale + ", is isQtcOrQTcF: " + isQtcOrQTcF);

if (isQtcOrQTcF) {   
    var qtcfDict = parseMultiRange(item.name);
    logger(JSON.stringify(qtcfDict, null, 2))
    var range;
    
    if (isMale && qtcfDict.male) {
        range = qtcfDict.male;
    } else if (!isMale && qtcfDict.female) {
        range = qtcfDict.female;
    } else {
        range = qtcfDict.default;
    }
    
    if (!range) return true;
    
    var min = range[0];
    var max = range[1];
    
    if (!checkRange(isRepeat, min, max)) return false;
} else {
    var minMax = parseRange(item.name);
    var min = minMax[0];
    var max = minMax[1];
    logger("Range: " + min + " - " + max);
    if (!checkRange(isRepeat, min, max)) return false;
}

return true;

function checkRange(repeat, min, max) {
    if (repeat && (item.value < min || item.value > max)) {
        customErrorMessage(RepeatErrorMsg);
        return false;
    }
    else if (!repeat && (item.value < min || item.value > max)) {
        customErrorMessage(errorMsg);
        return false;
    }
    return true;
}

function containsValue(input, keyword) {
    if (input == null) {
        return false;
    }

    var value = input.toString().toLowerCase();
    return value.indexOf(keyword) !== -1;
}

function getItemGroupName(form) {
    var groupName = "";
    for (var i = 0; i < form.form.itemGroups.length; i++) {
        var group = form.form.itemGroups[i];
        var items = group.items;
        if (!items || items.length < 1) continue;

        for (var j = 0; j < items.length; j++) {
            var it = items[j];
            if (it.id === item.id) {
                groupName = group.name;
                return groupName;
            }
        }
        if (groupName) break;
    }
    return groupName;
}

function parseRange(input) {
    if (!input) return null;

    var content = input;

    var matches = input.match(/\(([^()]*)\)/g);
    if (matches && matches.length > 0) {
        content = matches[matches.length - 1];
        content = content.replace(/[()]/g, "");
    }

    content = content.replace(/&nbsp;/gi, " ");
    content = content.replace(/–/g, "-");
    content = content.replace(/to/gi, "-");
    content = content.replace(/≤/g, "<=");
    content = content.replace(/≥/g, ">=");
    content = content.replace(/p\s*:/gi, "P:");
    content = content.replace(/i\s*:/gi, "I:");

    content = content.replace(/,/g, "");

    var pMatch = content.match(/P:\s*([^,]+)/);
    var iMatch = content.match(/I:\s*([^,]+)/);

    var min = null;
    var max = null;

    function adjustValue(value, operator) {
        var num = parseFloat(value);

        if (operator === ">") return num + 1;
        if (operator === "<") return num - 1;

        return num;
    }

    function extractOpNum(str) {
        var opMatch = str.match(/(<=|>=|<|>)/);
        var numMatch = str.match(/-?[\d.]+/);

        if (!numMatch) return null;

        return {
            op: opMatch ? opMatch[1] : null,
            val: numMatch[0]
        };
    }

    if (pMatch && iMatch) {
        var pObj = extractOpNum(pMatch[1]);
        var iObj = extractOpNum(iMatch[1]);

        if (pObj && iObj) {
            min = adjustValue(iObj.val, iObj.op);
            max = adjustValue(pObj.val, pObj.op);

            if (min > max) {
                var temp = min;
                min = max;
                max = temp;
            }
            return [min, max];
        }
    }

    if (iMatch) {
        content = iMatch[1];
    } else if (pMatch) {
        content = pMatch[1];
    }

    content = content.replace(/[a-zA-Z°%]/g, "");
    content = content.replace(/[^\d<>=\-. ]/g, "");
    content = content.replace(/\s+/g, " ");
    content = content.replace(/^\s+|\s+$/g, "");

    var rangeMatch = content.match(/(-?[\d.]+)\s*-\s*(-?[\d.]+)/);
    if (rangeMatch) {
        min = parseFloat(rangeMatch[1]);
        max = parseFloat(rangeMatch[2]);

        if (min > max) {
            var temp2 = min;
            min = max;
            max = temp2;
        }

        return [min, max];
    }

    var lessMatch = content.match(/(<=|<)\s*(-?[\d.]+)/);
    if (lessMatch) {
        var op = lessMatch[1];
        var val = parseFloat(lessMatch[2]);

        min = 0;
        max = (op === "<") ? val - 1 : val;

        return [min, max];
    }

    var greaterMatch = content.match(/(>=|>)\s*(-?[\d.]+)/);
    if (greaterMatch) {
        var op2 = greaterMatch[1];
        var val2 = parseFloat(greaterMatch[2]);

        min = (op2 === ">") ? val2 + 1 : val2;
        max = 999;

        return [min, max];
    }

    var singleMatch = content.match(/-?[\d.]+/);
    if (singleMatch) {
        var v = parseFloat(singleMatch[0]);
        return [v, v];
    }

    return null;
}

function parseMultiRange(input) {
    logger("parseMultiRange input: " + input);

    if (!input) return null;

    var result = {};

    var match = input.match(/\((.*)\)/);
    if (!match) return null;

    var content = match[1].toLowerCase();

    content = content
        .replace(/&nbsp;/gi, " ")
        .replace(/≤/g, "<=")
        .replace(/≥/g, ">=")
        .replace(/–/g, "-")
        .replace(/to/gi, "-")
        .replace(/\s+/g, " ")
        .trim();

    var parts = content.split(/[,|;/\-]+/);

    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        if (!part) continue;

        var label = null;

        if (/\b(female|f)\b/.test(part)) label = "female";
        else if (/\b(male|m)\b/.test(part)) label = "male";

        var rangeMatch = part.match(/(<=|>=|<|>)?\s*-?[\d.]+(\s*-\s*-?[\d.]+)?/);

        if (!rangeMatch) continue;

        var rule = rangeMatch[0];

        var parsed = parseRange("(" + rule + ")");

        if (!parsed) continue;

        if (label) {
            result[label] = parsed;
        } else {
            result.default = parsed;
        }
    }

    logger("Final parsed ranges: " + JSON.stringify(result));
    return Object.keys(result).length ? result : null;
}