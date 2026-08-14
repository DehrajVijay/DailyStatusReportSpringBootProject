package com.statusreport;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Writes the daily status report to an Excel (.xlsx) file using Apache POI.
 *
 * <p>Called by {@link ReportService#generate}. Output example:
 * {@code Documents/DailyStatusReports/DailyStatus_2026-07-09.xlsx}
 *
 * <p>Creates one sheet with header row, task rows, and a Total row for hours spent.
 * Header colors come from {@link ReportStyleConfig}.
 */
@Component
public class StatusReportExcelWriter {

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

    public StatusReportExcelWriter() {
        this(ReportStyleConfig.load());
    }

    public StatusReportExcelWriter(ReportStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * Creates the Excel file at the given path (overwrites if exists).
     *
     * @param outputFile full path including filename
     * @param tasks      validated tasks from {@link TaskParser}
     */
    public void write(Path outputFile, List<TaskEntry> tasks) throws IOException {
        Files.createDirectories(outputFile.getParent());

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Daily Status");

            CellStyle headerStyle = createHeaderStyle((XSSFWorkbook) workbook);
            CellStyle centerStyle = createCenterStyle(workbook);
            CellStyle wrapStyle = createWrapStyle(workbook);
            CellStyle totalLabelStyle = createTotalLabelStyle(workbook);

            // Row 0: column headers
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerStyle);
            }

            // Rows 1..n: one row per task
            int rowNum = 1;
            double totalSpent = 0;

            for (int i = 0; i < tasks.size(); i++) {
                TaskEntry task = tasks.get(i);
                Row row = sheet.createRow(rowNum++);

                setCell(row, 0, i + 1, centerStyle);
                setCell(row, 1, task.getTaskName(), wrapStyle);
                setCell(row, 2, task.getCompletionPercent() + " %", centerStyle);
                setCell(row, 3, HoursFormatter.format(task.getEstimatedHours()), centerStyle);
                setCell(row, 4, HoursFormatter.format(task.getHoursSpent()), centerStyle);
                setCell(row, 5, HoursFormatter.format(task.getRemainingHours()), centerStyle);
                setCell(row, 6, task.getRemarks(), centerStyle);
                setCell(row, 7, task.getCompletionDate(), centerStyle);

                totalSpent += task.getHoursSpent();
            }

            // Last row: Total hours spent
            Row totalRow = sheet.createRow(rowNum);
            Cell totalLabelCell = totalRow.createCell(3);
            totalLabelCell.setCellValue("Total");
            totalLabelCell.setCellStyle(totalLabelStyle);

            Cell totalValueCell = totalRow.createCell(4);
            totalValueCell.setCellValue(HoursFormatter.format(totalSpent));
            totalValueCell.setCellStyle(centerStyle);

            sheet.setColumnWidth(0, 2500);
            sheet.setColumnWidth(1, 14000);
            for (int col = 2; col <= 7; col++) {
                sheet.setColumnWidth(col, 5500);
            }

            try (OutputStream out = Files.newOutputStream(outputFile)) {
                workbook.write(out);
            }
        }
    }

    private static void setCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void setCell(Row row, int column, int value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(styleConfig.getHeaderFontRgb(), null));
        style.setFont(font);
        style.setFillForegroundColor(new XSSFColor(styleConfig.getHeaderBackgroundRgb(), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        applyBorders(style);
        return style;
    }

    private static CellStyle createCenterStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        applyBorders(style);
        return style;
    }

    private static CellStyle createWrapStyle(Workbook workbook) {
        CellStyle style = createCenterStyle(workbook);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        return style;
    }

    private static CellStyle createTotalLabelStyle(Workbook workbook) {
        CellStyle style = createCenterStyle(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static void applyBorders(CellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }
}
