package com.statusreport;

import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

/**
 * Opens an Outlook email draft with the status report as an HTML table in the body.
 *
 * <p>Flow:
 * <ol>
 *   <li>Load email settings from {@link OutlookConfig}</li>
 *   <li>Build HTML via {@link StatusReportHtmlBuilder}</li>
 *   <li>Write HTML to a temp file</li>
 *   <li>Run PowerShell script (send-outlook-draft.ps1) which uses Outlook COM API</li>
 *   <li>Outlook opens draft — user reviews and clicks Send manually</li>
 * </ol>
 *
 * <p>Requires Windows with desktop Outlook installed.
 */
@Service
public class OutlookDraftService {

    private static final Path OUTPUT_DIR = Paths.get(
            System.getProperty("user.home"),
            "Documents",
            "DailyStatusReports"
    );
    private static final Path SCRIPT_PATH = OUTPUT_DIR.resolve("scripts").resolve("send-outlook-draft.ps1");

    private final ReportStyleConfig styleConfig;

    public OutlookDraftService() {
        this(ReportStyleConfig.load());
    }

    public OutlookDraftService(ReportStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * Creates and displays an Outlook draft email (does not send automatically).
     *
     * @param tasks validated task list (same data as Excel)
     * @throws IllegalStateException if mail.to is not configured
     * @throws IOException if PowerShell script or Outlook fails
     */
    public void openDraft(List<TaskEntry> tasks) throws IOException {
        OutlookConfig outlookConfig = OutlookConfig.load();
        if (!outlookConfig.hasRecipient()) {
            throw new IllegalStateException(
                    "Recipient email is not configured. Edit:\n" + OutlookConfig.getUserConfigPath()
            );
        }

        StatusReportHtmlBuilder htmlBuilder = new StatusReportHtmlBuilder(styleConfig, outlookConfig);
        String html = htmlBuilder.buildHtml(tasks);

        Files.createDirectories(OUTPUT_DIR);
        Path htmlFile = Files.createTempFile(OUTPUT_DIR, "DailyStatus_", ".html");
        Files.write(htmlFile, html.getBytes(StandardCharsets.UTF_8));

        Path scriptPath = ensurePowerShellScript();

        List<String> command = new ArrayList<>();
        command.add("powershell.exe");
        command.add("-ExecutionPolicy");
        command.add("Bypass");
        command.add("-File");
        command.add(scriptPath.toString());
        command.add("-HtmlFile");
        command.add(htmlFile.toAbsolutePath().toString());
        command.add("-To");
        command.add(outlookConfig.getTo());
        command.add("-Subject");
        command.add(outlookConfig.buildSubject());
        if (outlookConfig.hasCc()) {
            command.add("-Cc");
            command.add(outlookConfig.getCc());
        }

        // SECURITY-REVIEW: launches PowerShell with fixed script path and user-supplied HTML file path
        ProcessBuilder processBuilder = new ProcessBuilder(command);
        processBuilder.redirectErrorStream(true);
        Process process = processBuilder.start();

        try {
            int exitCode = process.waitFor();
            if (exitCode != 0) {
                throw new IOException("Outlook draft script failed with exit code " + exitCode);
            }
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
            throw new IOException("Outlook draft script was interrupted.", ex);
        }
    }

    /**
     * Copies send-outlook-draft.ps1 from classpath to Documents folder on first use.
     * Outlook automation must run via an external script (Java cannot talk to Outlook directly).
     */
    private static Path ensurePowerShellScript() throws IOException {
        Files.createDirectories(SCRIPT_PATH.getParent());

        try (InputStream input = OutlookDraftService.class.getResourceAsStream("/send-outlook-draft.ps1")) {
            if (input == null) {
                throw new IOException("Missing send-outlook-draft.ps1 in application resources.");
            }
            Files.copy(input, SCRIPT_PATH, StandardCopyOption.REPLACE_EXISTING);
        }
        return SCRIPT_PATH;
    }
}
