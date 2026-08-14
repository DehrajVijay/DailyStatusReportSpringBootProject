package com.statusreport;

import java.util.List;

/**
 * Builds an HTML table for the Outlook email body.
 *
 * <p>Same columns and data as {@link StatusReportExcelWriter}, but output is HTML
 * with inline CSS (required for Outlook). Includes greeting and signature from
 * {@link OutlookConfig}.
 *
 * <p>Called by {@link OutlookDraftService#openDraft} before the PowerShell script runs.
 */
public class StatusReportHtmlBuilder {

    private static final String[] HEADERS = {
            "Sr. no.",
            "Task Name",
            "Completion Status",
            "Estimated Hours",
            "Hours Spent( HRS)",
            "Remaining efforts",
            "Remarks",
            "Completion Date(MM/DD/YYYY)"
    };

    private final ReportStyleConfig styleConfig;
    private final OutlookConfig outlookConfig;

    public StatusReportHtmlBuilder(ReportStyleConfig styleConfig, OutlookConfig outlookConfig) {
        this.styleConfig = styleConfig;
        this.outlookConfig = outlookConfig;
    }

    /**
     * Produces a complete HTML document string for Outlook's HTMLBody property.
     */
    public String buildHtml(List<TaskEntry> tasks) {
        String headerBg = styleConfig.getHeaderBackgroundHex();
        String headerFont = styleConfig.getHeaderFontHex();
        String headerStyle = "background-color:" + headerBg + ";color:" + headerFont
                + ";border:1px solid #000000;padding:6px;text-align:center;font-weight:bold;";
        String cellCenter = "border:1px solid #000000;padding:6px;text-align:center;";
        String cellLeft = "border:1px solid #000000;padding:6px;text-align:left;";
        String totalStyle = cellCenter + "font-weight:bold;";

        StringBuilder html = new StringBuilder();
        html.append("<html><body style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt;\">");
        html.append("<p>").append(formatMultiline(outlookConfig.getGreeting())).append("</p>");

        html.append("<table style=\"border-collapse:collapse;width:100%;max-width:1100px;\">");

        html.append("<tr>");
        for (String header : HEADERS) {
            html.append("<th style=\"").append(headerStyle).append("\">")
                    .append(escapeHtml(header))
                    .append("</th>");
        }
        html.append("</tr>");

        double totalSpent = 0;
        for (int i = 0; i < tasks.size(); i++) {
            TaskEntry task = tasks.get(i);
            totalSpent += task.getHoursSpent();

            html.append("<tr>");
            html.append("<td style=\"").append(cellCenter).append("\">").append(i + 1).append("</td>");
            html.append("<td style=\"").append(cellLeft).append("\">")
                    .append(escapeHtml(task.getTaskName()).replace("\n", "<br>"))
                    .append("</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(task.getCompletionPercent()).append(" %</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(HoursFormatter.format(task.getEstimatedHours())).append("</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(HoursFormatter.format(task.getHoursSpent())).append("</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(HoursFormatter.format(task.getRemainingHours())).append("</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(escapeHtml(task.getRemarks())).append("</td>");
            html.append("<td style=\"").append(cellCenter).append("\">")
                    .append(escapeHtml(task.getCompletionDate())).append("</td>");
            html.append("</tr>");
        }

        html.append("<tr>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("<td style=\"").append(totalStyle).append("\">Total</td>");
        html.append("<td style=\"").append(totalStyle).append("\">")
                .append(HoursFormatter.format(totalSpent)).append("</td>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("<td style=\"").append(cellCenter).append("\"></td>");
        html.append("</tr>");

        html.append("</table>");
        html.append("<p>").append(formatMultiline(outlookConfig.getSignature())).append("</p>");
        html.append("</body></html>");

        return html.toString();
    }

    /** Converts mail.greeting / mail.signature with &lt;br&gt; tags into safe HTML paragraphs */
    private static String formatMultiline(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }

        String[] lines = text.split("(?i)<br\\s*/?>|\\n");
        StringBuilder formatted = new StringBuilder();
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                formatted.append("<br>");
            }
            formatted.append(escapeHtml(lines[i].trim()));
        }
        return formatted.toString();
    }

    /** Prevents XSS in email body — user task names are treated as plain text */
    private static String escapeHtml(String value) {
        if (value == null) {
            return "";
        }
        return value
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;");
    }
}
