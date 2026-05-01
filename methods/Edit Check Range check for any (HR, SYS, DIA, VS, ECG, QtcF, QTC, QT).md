const value = itemJson.item.value;
const minimum_range = 60;
const maximum_range = 100;

if (value >= minimum_range && value <= maximum_range) {
    return true;
}

return false;