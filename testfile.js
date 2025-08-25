function calculateDaysBetweenDates(begin, end) {
    const oneDay = 24 * 60 * 60 * 1000; // hours * minutes * seconds * milliseconds
    const diffInMs = Math.abs(end - begin);
    return Math.round(diffInMs / oneDay);
}