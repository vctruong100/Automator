const value = itemJson.item.value;
const minimum_range = 60;
const maximum_range = 100;

try {
    if (value >= minimum_range && value <= maximum_range) {
        return true;
    }

    return false;
} catch (e) {
    logger("Error in main execution logic: " + e.message);
    return null;
}