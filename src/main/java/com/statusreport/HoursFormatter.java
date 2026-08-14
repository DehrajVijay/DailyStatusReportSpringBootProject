package com.statusreport;

/**
 * Formats hour values for display in Excel and email.
 *
 * <p>Examples: {@code 1 hr.}, {@code 2 hrs.}, {@code 0.5 hrs.}
 * Uses "hr." only when the value is exactly 1.
 */
public final class HoursFormatter {

    private HoursFormatter() {
    }

    public static String format(double hours) {
        String suffix = Math.abs(hours - 1.0) < 0.001 ? " hr." : " hrs.";
        if (hours == Math.rint(hours)) {
            return ((long) hours) + suffix;
        }
        return hours + suffix;
    }
}
