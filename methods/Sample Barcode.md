// Version: v1
// Purpose: Manages sample barcode validation/formatting.

const itemNames = [
    "BE_Date of sample 0 (PK)",
    "BE_Date of sample 1 (ADA, nAb)",
    "BE_Date of sample 2",
    "BE_Date of sample 3",
    "BE_Date of sample 4 (Serum Biomarker)",
    "BE_Date of sample 5 (Plasma biomarker)",
    "BE_Date of sample 6 (WB DNA)",
    
]

try {
    return parseItemContext(itemNames);
} catch (e) {
    logger("Error in main execution logic: " + e);
    return null;
}

// Parses related item data contexts to find collected barcodes across multiple item names.
function parseItemContext(item) {
    for (var i = 0; i < item.length; i++) {
        var context = JSON.parse(getRelatedItemDataContext(item[i]))
        var collectedBarcodes = context.collectedBarcodes;
        if (collectedBarcodes && collectedBarcodes != null) return collectedBarcodes;
    }
    return null;
}