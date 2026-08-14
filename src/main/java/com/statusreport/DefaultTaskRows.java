package com.statusreport;

import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * Builds the initial task rows shown when the form first loads.
 *
 * <p>Same defaults as the old Swing app:
 * <ol>
 *   <li>SOD: Performance Review (fixed row)</li>
 *   <li>5 empty rows for your daily tasks</li>
 *   <li>Break (fixed row at the bottom)</li>
 * </ol>
 *
 * <p>JavaScript in index.html also prevents removing SOD/Break rows.
 */
@Component
public class DefaultTaskRows {

    public static final String TASK_SOD = "SOD: Performance Review";
    public static final String TASK_BREAK = "Break";

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("MM/dd/yyyy");

    /**
     * Creates the starter rows for a new daily report form.
     * Called from {@link ReportController#index} on first page load.
     */
    public List<TaskEntry> createDefaultRows() {
        List<TaskEntry> tasks = new ArrayList<>();
        tasks.add(createRow(TASK_SOD, 100, 0.5, 0, "Complete", today()));
        for (int i = 0; i < 5; i++) {
            tasks.add(createRow("", 0, 0, 0, "In Progress", today()));
        }
        tasks.add(createRow(TASK_BREAK, 100, 1, 1, "Complete", today()));
        return tasks;
    }

    /** Used by UI logic to know which rows cannot be deleted */
    public boolean isFixedTask(String taskName) {
        if (taskName == null) {
            return false;
        }
        String trimmed = taskName.trim();
        return TASK_SOD.equals(trimmed) || TASK_BREAK.equals(trimmed);
    }

    private static TaskEntry createRow(String name, int completion, double estimated,
                                       double spent, String remarks, String date) {
        TaskEntry task = new TaskEntry();
        task.setTaskName(name);
        task.setCompletionPercent(completion);
        task.setEstimatedHours(estimated);
        task.setHoursSpent(spent);
        task.setRemarks(remarks);
        task.setCompletionDate(date);
        return task;
    }

    private static String today() {
        return LocalDate.now().format(DATE_FORMAT);
    }
}
