package com.statusreport;

import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.List;

/**
 * Orchestrates the full "Generate" workflow — the heart of the business logic.
 *
 * <p>Called by {@link ReportController#generate}. Steps:
 * <ol>
 *   <li>{@link TaskParser#parse} — validate and clean task rows</li>
 *   <li>{@link StatusReportExcelWriter#write} — save .xlsx file</li>
 *   <li>{@link OutlookDraftService#openDraft} — open Outlook email draft</li>
 * </ol>
 *
 * <p>{@code @Service} marks this as a Spring bean (business logic layer).
 */
@Service
public class ReportService {

    /** Where Excel files and Outlook config are stored on the user's PC */
    public static final Path OUTPUT_DIR = Paths.get(
            System.getProperty("user.home"),
            "Documents",
            "DailyStatusReports"
    );

    private final StatusReportExcelWriter excelWriter;
    private final OutlookDraftService outlookDraftService;
    private final TaskParser taskParser;

    public ReportService(StatusReportExcelWriter excelWriter,
                         OutlookDraftService outlookDraftService,
                         TaskParser taskParser) {
        this.excelWriter = excelWriter;
        this.outlookDraftService = outlookDraftService;
        this.taskParser = taskParser;
    }

    /**
     * Runs the complete generate pipeline: validate → Excel → Outlook draft.
     *
     * @param rawTasks tasks submitted from the HTML form (may have empty rows)
     * @return path to the created Excel file
     */
    public ReportResult generate(List<TaskEntry> rawTasks) throws IOException {
        List<TaskEntry> tasks = taskParser.parse(rawTasks);

        String fileName = "DailyStatus_" + LocalDate.now() + ".xlsx";
        Path outputFile = OUTPUT_DIR.resolve(fileName);

        excelWriter.write(outputFile, tasks);
        outlookDraftService.openDraft(tasks);

        return new ReportResult(outputFile);
    }

    /** Simple holder returned to the controller so it can show the Excel path on screen */
    public static class ReportResult {
        private final Path excelFile;

        public ReportResult(Path excelFile) {
            this.excelFile = excelFile;
        }

        public Path getExcelFile() {
            return excelFile;
        }
    }
}
