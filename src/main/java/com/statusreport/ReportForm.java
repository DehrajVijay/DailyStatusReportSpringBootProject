package com.statusreport;

import java.util.ArrayList;
import java.util.List;

/**
 * Form backing object — binds HTML form fields to Java objects.
 *
 * <p>When the user submits the form, Spring maps fields like:
 * {@code tasks[0].taskName}, {@code tasks[0].completionPercent}, etc.
 * into this object's {@code tasks} list.
 *
 * <p>Used in index.html as {@code th:object="${form}"} and {@code *{tasks[0].taskName}}.
 * Getters/setters are required for Spring's data binding to work.
 */
public class ReportForm {

    private List<TaskEntry> tasks = new ArrayList<>();

    public List<TaskEntry> getTasks() {
        return tasks;
    }

    public void setTasks(List<TaskEntry> tasks) {
        this.tasks = tasks != null ? tasks : new ArrayList<>();
    }
}
