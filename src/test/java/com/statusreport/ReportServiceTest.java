package com.statusreport;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InOrder;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.inOrder;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * Unit tests for {@link ReportService} — no Spring context, no disk, no
 * Outlook.
 * Collaborators are mocked so we test orchestration only.
 */
@ExtendWith(MockitoExtension.class)
class ReportServiceTest {

    @Mock
    private StatusReportExcelWriter excelWriter;

    @Mock
    private OutlookDraftService outlookDraftService;

    @Mock
    private TaskParser taskParser;

    @InjectMocks
    private ReportService reportService;

    private List<TaskEntry> rawTasks;
    private List<TaskEntry> parsedTasks;
    private Path expectedOutputFile;

    @BeforeEach
    void setUp() {
        rawTasks = List.of(sampleTask("Fix login bug", 80, 4.0, 3.0));
        parsedTasks = List.of(sampleTask("Fix login bug", 80, 4.0, 3.0));
        expectedOutputFile = ReportService.OUTPUT_DIR.resolve(
                "DailyStatus_" + LocalDate.now() + ".xlsx");
    }

    @Test
    void generate_runsValidateThenExcelThenOutlook_inOrder() throws IOException {
        when(taskParser.parse(rawTasks)).thenReturn(parsedTasks);
        doNothing().when(excelWriter).write(any(Path.class), eq(parsedTasks));
        doNothing().when(outlookDraftService).openDraft(parsedTasks);

        ReportService.ReportResult result = reportService.generate(rawTasks);

        assertEquals(expectedOutputFile, result.getExcelFile());

        InOrder order = inOrder(taskParser, excelWriter, outlookDraftService);
        order.verify(taskParser).parse(rawTasks);
        order.verify(excelWriter).write(expectedOutputFile, parsedTasks);
        order.verify(outlookDraftService).openDraft(parsedTasks);
    }

    @Test
    void generate_whenValidationFails_doesNotWriteExcelOrOpenOutlook() throws Exception {
        when(taskParser.parse(rawTasks))
                .thenThrow(new IllegalArgumentException("Row 2: Completion % must be between 0 and 100."));

        assertThrows(IllegalArgumentException.class, () -> reportService.generate(rawTasks));

        verify(excelWriter, never()).write(any(), any());
        verify(outlookDraftService, never()).openDraft(any());
    }

    @Test
    void generate_whenExcelWriteFails_doesNotOpenOutlook() throws Exception {
        when(taskParser.parse(rawTasks)).thenReturn(parsedTasks);
        doThrow(new IOException("disk full"))
                .when(excelWriter).write(any(Path.class), eq(parsedTasks));

        assertThrows(IOException.class, () -> reportService.generate(rawTasks));

        verify(outlookDraftService, never()).openDraft(any());
    }

    @Test
    void generate_whenOutlookFails_afterExcelWasWritten() throws IOException {
        when(taskParser.parse(rawTasks)).thenReturn(parsedTasks);
        doNothing().when(excelWriter).write(any(Path.class), eq(parsedTasks));
        doThrow(new IOException("Outlook not installed"))
                .when(outlookDraftService).openDraft(parsedTasks);

        assertThrows(IOException.class, () -> reportService.generate(rawTasks));

        verify(excelWriter).write(expectedOutputFile, parsedTasks);
        verify(outlookDraftService).openDraft(parsedTasks);
    }

    private static TaskEntry sampleTask(String name, int percent, double estimated, double spent) {
        TaskEntry task = new TaskEntry();
        task.setTaskName(name);
        task.setCompletionPercent(percent);
        task.setEstimatedHours(estimated);
        task.setHoursSpent(spent);
        task.setRemarks("In Progress");
        task.setCompletionDate("09/10/2026");
        return task;
    }
}
