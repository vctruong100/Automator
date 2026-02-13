const itemName = "Dispensed kit number #4";
var item = itemJson.item;
var value = findFirstItemValueByName(formJson, itemName);
if (value && value !== null) {
    logger(value);
    if (item && item.value !== null && item.value !== "") return true;
    else {
        customErrorMessage("Please complete when kits are returned")
        return false;
    }
}

return true;