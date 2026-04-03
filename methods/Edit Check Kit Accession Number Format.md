const value = itemJson.item.value;

return isValidFormat(value);

function isValidFormat(value)
{
    var pattern = /^[A-Za-z]{2}[0-9]{5}$/;

    if (typeof value !== "string")
    {
        return false;
    }

    if (pattern.test(value))
    {
        return true;
    }
    else
    {
        return false;
    }
}
