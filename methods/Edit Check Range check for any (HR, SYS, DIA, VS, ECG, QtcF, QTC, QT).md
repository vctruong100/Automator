const value = itemJson.item.value;
const minimum_range = 60;
const maximum_range = 100;
const errorMsg = "OOR Repeat"; // Update if you want custom Error message

if (value >= minimum_range && value <= maximum_range) {
    return true;
}

customErrorMessage(errorMsg);
return false;