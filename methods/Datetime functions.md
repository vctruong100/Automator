/* jshint strict: false */

// Version: v1
// Purpose: Provides reusable datetime parsing and formatting utilities for ClinSpark automation scripts.


// Add time to input time. Uses datetime datatype.
function getTimeRange(datetimeStr) {
    if (!datetimeStr) {
        return "";
    }

    var parts = datetimeStr.split("T");
    if (parts.length < 2) {
        return "";
    }

    var timePart = parts[1]; 

    var timePieces = timePart.split(":");
    var hour = parseInt(timePieces[0], 10);
    var minute = parseInt(timePieces[1], 10);
    var second = parseInt(timePieces[2], 10);

    var totalSeconds = (hour * 3600) + (minute * 60) + second;

    // Add 30 minutes
    var minSeconds = totalSeconds + (30 * 60);

    // Add 45 minutes
    var maxSeconds = totalSeconds + (45 * 60);

    // Transforms data using: formatTime.
    function formatTime(seconds) {
        var h = Math.floor(seconds / 3600);
        var m = Math.floor((seconds % 3600) / 60);
        var s = seconds % 60;

        if (h >= 24) {
            h = h % 24;
        }

        return (
            (h < 10 ? "0" : "") + h + ":" +
            (m < 10 ? "0" : "") + m + ":" +
            (s < 10 ? "0" : "") + s
        );
    }

    return formatTime(minSeconds) + " to " + formatTime(maxSeconds);
}

// Convert military time to standard time
function militaryToStandard(timeStr) {
    if (!timeStr) {
        return "";
    }

    var parts = timeStr.split(":");

    var hour = parseInt(parts[0], 10);
    var minute = parts[1];
    var second = parts[2] || "00";

    var ampm = "AM";

    if (hour >= 12) {
        ampm = "PM";
    }

    if (hour === 0) {
        hour = 12;
    } else if (hour > 12) {
        hour = hour - 12;
    }

    return hour + ":" + minute + ":" + second + " " + ampm;
}