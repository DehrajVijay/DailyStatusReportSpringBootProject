package com.statusreport;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * Loads table header colors for Excel and HTML email.
 *
 * <p>Reads RGB values from {@code report-style.properties} on the classpath.
 * Excel uses byte arrays ({@link #getHeaderBackgroundRgb}); HTML uses hex ({@link #getHeaderBackgroundHex}).
 */
public final class ReportStyleConfig {

    private static final int DEFAULT_BG_R = 31;
    private static final int DEFAULT_BG_G = 78;
    private static final int DEFAULT_BG_B = 120;
    private static final int DEFAULT_FONT_R = 255;
    private static final int DEFAULT_FONT_G = 255;
    private static final int DEFAULT_FONT_B = 255;

    private final int headerBackgroundR;
    private final int headerBackgroundG;
    private final int headerBackgroundB;
    private final int headerFontR;
    private final int headerFontG;
    private final int headerFontB;

    private ReportStyleConfig(int headerBackgroundR, int headerBackgroundG, int headerBackgroundB,
                              int headerFontR, int headerFontG, int headerFontB) {
        this.headerBackgroundR = headerBackgroundR;
        this.headerBackgroundG = headerBackgroundG;
        this.headerBackgroundB = headerBackgroundB;
        this.headerFontR = headerFontR;
        this.headerFontG = headerFontG;
        this.headerFontB = headerFontB;
    }

    /** Loads colors from report-style.properties, or uses blue header / white text defaults */
    public static ReportStyleConfig load() {
        Properties properties = new Properties();
        try (InputStream input = ReportStyleConfig.class.getResourceAsStream("/report-style.properties")) {
            if (input != null) {
                properties.load(input);
            }
        } catch (IOException ex) {
            // Use defaults when the config file is missing or unreadable.
        }

        return new ReportStyleConfig(
                readColor(properties, "header.background.r", DEFAULT_BG_R),
                readColor(properties, "header.background.g", DEFAULT_BG_G),
                readColor(properties, "header.background.b", DEFAULT_BG_B),
                readColor(properties, "header.font.r", DEFAULT_FONT_R),
                readColor(properties, "header.font.g", DEFAULT_FONT_G),
                readColor(properties, "header.font.b", DEFAULT_FONT_B)
        );
    }

    /** For Apache POI Excel cell styling */
    public byte[] getHeaderBackgroundRgb() {
        return rgbBytes(headerBackgroundR, headerBackgroundG, headerBackgroundB);
    }

    public byte[] getHeaderFontRgb() {
        return rgbBytes(headerFontR, headerFontG, headerFontB);
    }

    /** For inline CSS in HTML email */
    public String getHeaderBackgroundHex() {
        return toHex(headerBackgroundR, headerBackgroundG, headerBackgroundB);
    }

    public String getHeaderFontHex() {
        return toHex(headerFontR, headerFontG, headerFontB);
    }

    private static String toHex(int red, int green, int blue) {
        return String.format("#%02X%02X%02X", red, green, blue);
    }

    private static int readColor(Properties properties, String key, int defaultValue) {
        String value = properties.getProperty(key);
        if (value == null || value.trim().isEmpty()) {
            return defaultValue;
        }
        try {
            int color = Integer.parseInt(value.trim());
            if (color < 0 || color > 255) {
                return defaultValue;
            }
            return color;
        } catch (NumberFormatException ex) {
            return defaultValue;
        }
    }

    private static byte[] rgbBytes(int red, int green, int blue) {
        return new byte[]{(byte) red, (byte) green, (byte) blue};
    }
}
