package com.statusreport;

import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * Validates and cleans task rows from the form before Excel/email generation.
 *
 * <p>Rules applied:
 * <ul>
 *   <li>Skip rows with empty task name</li>
 *   <li>Completion % must be 0–100</li>
 *   <li>Hours cannot be negative</li>
 *   <li>Default remarks = "In Progress", default date = today</li>
 * </ul>
 *
 * <p>Throws {@link IllegalArgumentException} with row number if validation fails.
 */
@Component
public class TaskParser {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("MM/dd/yyyy");

    /**
     * Converts raw form data into a clean list ready for Excel and Outlook.
     *
     * @param rawTasks tasks as submitted from the browser (may include blank rows)
     * @return validated tasks (at least one required)
     */
    public List<TaskEntry> parse(List<TaskEntry> rawTasks) {
        if (rawTasks == null || rawTasks.isEmpty()) {
            throw new IllegalArgumentException("Add at least one task with a task name.");
        }

        List<TaskEntry> tasks = new ArrayList<>();
        String today = LocalDate.now().format(DATE_FORMAT);

        for (int row = 0; row < rawTasks.size(); row++) {
            TaskEntry raw = rawTasks.get(row);
            if (raw == null) {
                continue;
            }

            String taskName = raw.getTaskName() != null ? raw.getTaskName().trim() : "";
            if (taskName.isEmpty()) {
                continue; // blank row — user left task name empty
            }

            TaskEntry task = new TaskEntry();
            task.setTaskName(taskName);
            task.setCompletionPercent(parseCompletionPercent(raw.getCompletionPercent(), row));
            task.setEstimatedHours(parseHours(raw.getEstimatedHours(), row, "Estimated Hours"));
            task.setHoursSpent(parseHours(raw.getHoursSpent(), row, "Spent Hours"));

            String remarks = raw.getRemarks() != null ? raw.getRemarks().trim() : "";
            task.setRemarks(remarks.isEmpty() ? "In Progress" : remarks);

            String completionDate = raw.getCompletionDate() != null ? raw.getCompletionDate().trim() : "";
            task.setCompletionDate(completionDate.isEmpty() ? today : completionDate);

            tasks.add(task);
        }

        if (tasks.isEmpty()) {
            throw new IllegalArgumentException("Add at least one task with a task name.");
        }

        return tasks;
    }

    private int parseCompletionPercent(int percent, int row) {
        if (percent < 0 || percent > 100) {
            throw new IllegalArgumentException("Row " + (row + 1) + ": Completion % must be between 0 and 100.");
        }
        return percent;
    }

    private double parseHours(double hours, int row, String fieldName) {
        if (hours < 0) {
            throw new IllegalArgumentException("Row " + (row + 1) + ": " + fieldName + " cannot be negative.");
        }
        return hours;
    }
}
