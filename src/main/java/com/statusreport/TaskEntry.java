package com.statusreport;

/**
 * One row in the daily status report — a single task or activity.
 *
 * <p>Used in three places:
 * <ul>
 *   <li>HTML form binding ({@link ReportForm})</li>
 *   <li>Excel output ({@link StatusReportExcelWriter})</li>
 *   <li>Outlook HTML email ({@link StatusReportHtmlBuilder})</li>
 * </ul>
 *
 * <p>Plain POJO (Plain Old Java Object) with getters/setters — no Spring annotations needed.
 */
public class TaskEntry {

    private String taskName;
    private int completionPercent;
    private double estimatedHours;
    private double hoursSpent;
    private String remarks;
    private String completionDate;

    public String getTaskName() {
        return taskName;
    }

    public void setTaskName(String taskName) {
        this.taskName = taskName;
    }

    public int getCompletionPercent() {
        return completionPercent;
    }

    public void setCompletionPercent(int completionPercent) {
        this.completionPercent = completionPercent;
    }

    public double getEstimatedHours() {
        return estimatedHours;
    }

    public void setEstimatedHours(double estimatedHours) {
        this.estimatedHours = estimatedHours;
    }

    public double getHoursSpent() {
        return hoursSpent;
    }

    public void setHoursSpent(double hoursSpent) {
        this.hoursSpent = hoursSpent;
    }

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }

    public String getCompletionDate() {
        return completionDate;
    }

    public void setCompletionDate(String completionDate) {
        this.completionDate = completionDate;
    }

    /**
     * Calculated field: estimated minus spent, never negative.
     * Shown in Excel and email but not entered by the user.
     */
    public double getRemainingHours() {
        return Math.max(estimatedHours - hoursSpent, 0);
    }
}
