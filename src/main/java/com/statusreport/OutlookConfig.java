package com.statusreport;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Properties;

/**
 * Loads Outlook email settings (To, CC, subject, greeting, signature).
 *
 * <p>Two-layer config (user overrides win):
 * <ol>
 *   <li>Defaults from classpath: {@code src/main/resources/outlook-config.properties}</li>
 *   <li>User file: {@code Documents/DailyStatusReports/outlook-config.properties}</li>
 * </ol>
 *
 * <p>On first run, copies defaults to the Documents folder if the user file does not exist.
 */
public final class OutlookConfig {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("MM/dd/yyyy");
    private static final Path USER_CONFIG = Paths.get(
            System.getProperty("user.home"),
            "Documents",
            "DailyStatusReports",
            "outlook-config.properties"
    );

    private final String to;
    private final String cc;
    private final String subjectPrefix;
    private final String greeting;
    private final String signature;

    private OutlookConfig(String to, String cc, String subjectPrefix, String greeting, String signature) {
        this.to = to;
        this.cc = cc;
        this.subjectPrefix = subjectPrefix;
        this.greeting = greeting;
        this.signature = signature;
    }

    /**
     * Loads merged config: classpath defaults + user overrides from Documents folder.
     */
    public static OutlookConfig load() throws IOException {
        ensureUserConfigExists();

        Properties properties = new Properties();
        try (InputStream classpath = OutlookConfig.class.getResourceAsStream("/outlook-config.properties")) {
            if (classpath != null) {
                properties.load(classpath);
            }
        }

        if (Files.exists(USER_CONFIG)) {
            try (InputStream userConfig = Files.newInputStream(USER_CONFIG)) {
                Properties userProperties = new Properties();
                userProperties.load(userConfig);
                properties.putAll(userProperties); // user values override defaults
            }
        }

        String to = properties.getProperty("mail.to", "").trim();
        String cc = properties.getProperty("mail.cc", "").trim();
        String subjectPrefix = properties.getProperty("mail.subject.prefix", "Daily Status Report").trim();
        String greeting = properties.getProperty("mail.greeting", "Hi,").trim();
        String signature = properties.getProperty("mail.signature", "").trim();
        if (signature.isEmpty()) {
            signature = properties.getProperty("mail.closing", "Thanks.").trim();
        }

        return new OutlookConfig(to, cc, subjectPrefix, greeting, signature);
    }

    public String getTo() {
        return to;
    }

    public String getCc() {
        return cc;
    }

    /** Example: "Daily Status Report - 07/09/2026" */
    public String buildSubject() {
        return subjectPrefix + " - " + LocalDate.now().format(DATE_FORMAT);
    }

    public String getGreeting() {
        return greeting;
    }

    public String getSignature() {
        return signature;
    }

    public boolean hasRecipient() {
        return to != null && !to.isEmpty();
    }

    public boolean hasCc() {
        return cc != null && !cc.isEmpty();
    }

    /** Path shown on the web page so user knows where to edit mail.to */
    public static Path getUserConfigPath() {
        return USER_CONFIG;
    }

    /** Creates Documents/outlook-config.properties from defaults if missing */
    private static void ensureUserConfigExists() throws IOException {
        if (Files.exists(USER_CONFIG)) {
            return;
        }

        Files.createDirectories(USER_CONFIG.getParent());
        try (InputStream defaults = OutlookConfig.class.getResourceAsStream("/outlook-config.properties")) {
            if (defaults != null) {
                Files.copy(defaults, USER_CONFIG);
            } else {
                Properties properties = new Properties();
                properties.setProperty("mail.to", "");
                properties.setProperty("mail.cc", "");
                properties.setProperty("mail.subject.prefix", "Daily Status Report");
                properties.setProperty("mail.greeting", "Hi,");
                properties.setProperty("mail.signature", "Thanks and Regards,<br>Your Name");
                try (OutputStream out = Files.newOutputStream(USER_CONFIG)) {
                    properties.store(out, "Edit mail.to and mail.cc with recipient email addresses");
                }
            }
        }
    }
}
