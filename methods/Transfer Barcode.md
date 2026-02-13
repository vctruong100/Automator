const itemNames = [
    "BE_Date of sample 0 (PK)",
    "BE_Date of sample 1 (ADA, nAb)",
    "BE_Date of sample 2",
    "BE_Date of sample 3",
    "BE_Date of sample 4 (Serum Biomarker)",
    "BE_Date of sample 5 (Plasma biomarker)",
    "BE_Date of sample 6 (WB DNA)",
    
]

return parseItemContext(itemNames);

function parseItemContext(item) {
    for (var i = 0; i < item.length; i++) {
        var context = JSON.parse(getRelatedItemDataContext(item[i]))
        var collectedBarcodes = context.transferBarcodes;
        if (collectedBarcodes && collectedBarcodes != null) return collectedBarcodes;
    }
    return null;
}