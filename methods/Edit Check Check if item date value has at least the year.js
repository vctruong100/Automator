var datePieces, year, month, day;
const value = itemJson.item.value;
var date = value.split("T")[0];

if (date) {
    datePieces = date.split("-");
    if (datePieces.length >= 3) {
        year = datePieces[0];
        month = parseInt(datePieces[1], 10) - 1;
        day = datePieces[2];
    }
}
logger(date);
logger(year);
if (year && year !== null) return true;
customErrorMessage("Cannot be blank. Must enter at least the Year");
return false;