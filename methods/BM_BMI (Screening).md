
// Add item names
const heightitemList = [
    "BM_HT", 
    "BM_HT_Visit 2",
    "BMI_HEIGHT",
    "BMI_HEIGHT #1", 
    "BMI_HEIGHT #2"
];
const weightitemList = [
    "BM_WT #1", 
    "BM_WT #2",
    "BMI_WEIGHT",
    "BMI_WEIGHT #1", 
    "BMI_WEIGHT #2"
];

// ======== Don't modify ========
const currentStudyEvent = formJson.form.studyEventName;
const item = itemJson.item;
const sigfig = itemJson.item.significantDigits;

var bmi = 0;

var htmaxCount = 2; 
var htlist = [];
var htavg = 0;

var wtmaxCount = 2;
var wtlist = [];
var wtavg = 0;

htlist = populateList(formJson, heightitemList, htmaxCount);
htavg = calculateAverage(htlist, sigfig);

wtlist = populateList(formJson, weightitemList, wtmaxCount);
wtavg = calculateAverage(wtlist, sigfig);

logger(htavg);
logger(wtavg);

if (htlist.length !== htmaxCount && wtlist.length !== wtmaxCount) return null;

var heightMtr = htavg / 100;

var factor = Math.pow(10, sigfig);
bmi = Math.round((wtavg / (heightMtr * heightMtr)) * factor) / factor;

log();

if (bmi) return bmi.toFixed(sigfig);

return null;

function log() {
    logger("Height in meter: " + heightMtr)
    logger("Factor: " + factor);
    logger("BMI: " + bmi);
}

function populateList(form, targetItem, maxCount) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    var list = [];
    var count = 0;
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = itemGroups.length - 1; i >= 0; i--) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (item && targetItem.indexOf(item.name) !== -1) {
                maxCount++;
                if (item.value !== null && !isNaN(item.value) && item.value !== "") {
                    list.push(parseFloat(item.value));
                }
            }
        }
    }
    return list;
}

function calculateAverage(values, sigfig) {
    if (values.length === 0) return null;
    var sum = 0;
    var count = 0;

    for (var i = 0; i < values.length; i++) {
        if (isNaN(values[i])) continue;
        else sum += values[i];
        count++;
    }
    if (count === 0) return null;

    var avg = sum / count;
    var factor = Math.pow(10, sigfig);
    return Math.round(avg * factor) / factor;
}