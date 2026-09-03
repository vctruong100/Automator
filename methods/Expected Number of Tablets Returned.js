/* jshint strict: false */
// Description: Calculates the expected number of tablets returned from prior visit date, dispensed quantity, prescribed daily dose, and current visit timing.

const formNames = [
    "REV_Review Questions (w/ Glucose monitoring) v3.0",
    "REV_Review Questions (w/ Glucose monitoring) (D36, D43, D50, D57, D64, D71, D78, D80, D81, D98, Early discontinuation) v3.0"
]

const itemName = [
    "Visit Date",
]

const studyeventMap = {
    "Wk7 Day 43": "Wk6 Day 38",
    "Wk8 Day 50": "Wk7 Day 43",
    "Wk9 Day 57": "Wk8 Day 50",
    "Wk10 Day 64": "Wk9 Day 57",
    "Wk11 Day 71": "Wk10 Day 64",
    "Wk12 Day 78": "Wk11 Day 71",
    "Wk12 Day 84": "Wk12 Day 79",
};

const dispensedItem = [
    "Number of tablets dispensed:",    
]

const prescribedItem = [
    "Number of tablets prescribed per day:",    
]

const studyevent = formJson.form.studyEventName;

const screeningNumber = formJson.form.subject.screeningNumber;
logger("Subject ID: " + screeningNumber);

var item = itemJson.item;
/*if (item && item.value !== null) {//Overwrite protection
    if (item.value <= 0) return (0).toFixed(0);
    logger('Current item is already set: ' + item.value + '. Skipping.');
    return item.value;
} else {
    logger('No value set for the current item. Proceeding with calculation.');
}*/

if (studyevent === "Unscheduled") {
    if (screeningNumber === "000500008") return 14;
    else if (screeningNumber === "000500007") return 2;
} 

function pullItemFromForm(form, targetItem) {
    var itemGroups = form.form.itemGroups;
    var group, items, item, i, j, value;
    
	if (!itemGroups || itemGroups.length < 1) return null;
    
    for (i = 0; i < itemGroups.length; i++) {
        group = itemGroups[i];
        if (!group || group.canceled) continue;
        for (j = 0; j < group.items.length; j++) {
            item = group.items[j];
            if (targetItem.indexOf(item.name) !== -1 && item.value !== null && !item.canceled && item.value !== "") return item.value;
        }
    }
    return null;
}

function pullForm(studyeventList, formNameList) {
    for (var i = 0; i < studyeventList.length; i++) {
        for (var j = 0; j < formNameList.length; j++) {
            var temp = checkForm(studyeventList[i], formNameList[j]);
            if (temp) return temp;
        }
    }
}

function checkForm(studyevent, form) {
    var arrayForms = findFormData(studyevent, form);
    var completedForm = collectCompleted(arrayForms, true);
    if (!completedForm || completedForm.length === 0) return null;
    return completedForm[0];
}

function collectCompleted(formDataArray, INCLUDE_NONCONFORMANT_DATA) {
    if (formDataArray == null) { return []; }
    var keepers = [];
    for (var i = formDataArray.length - 1; i >= 0; i--) {
        var formData = formDataArray[i];
        if (formData.form.canceled == false && formData.form.itemGroups[0].canceled == false && (formData.form.dataCollectionStatus == 'Complete' || 
                (INCLUDE_NONCONFORMANT_DATA == true && formData.form.dataCollectionStatus == 'Nonconformant') || formData.form.dataCollectionStatus == "Incomplete")) {
            keepers.push(formData);
        } else {

        }
    }
    return keepers;
}

try {
    var expecteddays = 0;
    var day = parseInt(studyevent.split(" ")[2]); //Wk7 Day 43 -> 43
    
    if (day == 84) expecteddays = 5;
    
    var prevDay = studyeventMap[studyevent];
    logger("studyevent: " + studyevent + ", previous visit: " + prevDay)
    var pastVisitform = pullForm([prevDay], formNames);
    var currentVisitform = pullForm([studyevent], formNames);
    if (!pastVisitform && expecteddays == 0) {
        if (day > 43 && day <= 84) expecteddays = 7;
        else if (day <= 43) expecteddays = 4;
    }
    else if (expecteddays == 0) {
        var pastdate = pullItemFromForm(pastVisitform, itemName);
        var currentdate = pullItemFromForm(currentVisitform, itemName);

        logger("Past Date: " + pastdate);
        logger("Current Date: " + currentdate);

        var pastParts = pastdate.split("-");
        var currentParts = currentdate.split("-");
        
        var pastDateObj = new Date(
            parseInt(pastParts[0], 10),
            parseInt(pastParts[1], 10) - 1,
            parseInt(pastParts[2], 10)
        );
        
        var currentDateObj = new Date(
            parseInt(currentParts[0], 10),
            parseInt(currentParts[1], 10) - 1,
            parseInt(currentParts[2], 10)
        );
        
        var diffMilliseconds = currentDateObj.getTime() - pastDateObj.getTime();
        var diffDays = Math.floor(diffMilliseconds / (1000 * 60 * 60 * 24));

        expecteddays = diffDays - 1;

        logger("Date Difference: " + diffDays);
        logger("Expected Days: " + expecteddays);
    }
    var dispensed = parseFloat(pullItemFromForm(formJson, dispensedItem));
    var prescribed = parseFloat(pullItemFromForm(formJson, prescribedItem));
    if (!dispensed || dispensed == null || dispensed == "" || !prescribed || prescribed == null || prescribed == "") return null;
    
    logger("Current Day: " + day);
    logger("expectedDays: " + expecteddays);
    
    var totalPrescribed = expecteddays * prescribed;
    logger("Tablets Dispensed: " + dispensed)
    logger("Tablet Prescribed: " + totalPrescribed)
    var expectedReturn = dispensed - totalPrescribed;
    if (expectedReturn <= 0) expectedReturn = 0;
    return String(expectedReturn.toFixed(0));

} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
}

