package com.statusreport;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;

import java.io.IOException;

/**
 * Web layer — handles HTTP requests from the browser.
 *
 * <p>Flow:
 * <pre>
 *   GET  /          → show the task form (index.html)
 *   POST /generate   → validate, create Excel, open Outlook draft
 * </pre>
 *
 * <p>{@code @Controller} tells Spring this class handles web pages (not REST JSON).
 * Return value {@code "index"} maps to {@code templates/index.html} via Thymeleaf.
 */
@Controller
public class ReportController {

    private final ReportService reportService;
    private final DefaultTaskRows defaultTaskRows;

    /**
     * Constructor injection: Spring creates ReportService and DefaultTaskRows,
     * then passes them here automatically. No {@code new ReportService()} needed.
     */
    public ReportController(ReportService reportService, DefaultTaskRows defaultTaskRows) {
        this.reportService = reportService;
        this.defaultTaskRows = defaultTaskRows;
    }

    /**
     * Shows the daily status form when user opens http://localhost:8081/
     *
     * <p>Prepares a {@link ReportForm} with default rows (SOD, 5 blank rows, Break)
     * and passes it to the Thymeleaf template as {@code ${form}}.
     */
    @GetMapping("/")
    public String index(Model model) {
        // After POST /generate we keep the submitted form; only create defaults on first visit
        if (!model.containsAttribute("form")) {
            ReportForm form = new ReportForm();
            form.setTasks(defaultTaskRows.createDefaultRows());
            model.addAttribute("form", form);
        }
        model.addAttribute("outputDir", ReportService.OUTPUT_DIR.toString());
        model.addAttribute("outlookConfigPath", OutlookConfig.getUserConfigPath().toString());
        return "index";
    }

    /**
     * Triggered when user clicks "Generate" on the form.
     *
     * <p>{@code @ModelAttribute("form")} binds submitted form fields (tasks[0].taskName, etc.)
     * into a ReportForm object automatically.
     *
     * <p>Delegates business logic to {@link ReportService#generate}, then shows success
     * or error message on the same page.
     */
    @PostMapping("/generate")
    public String generate(@ModelAttribute("form") ReportForm form, Model model) {
        try {
            ReportService.ReportResult result = reportService.generate(form.getTasks());
            model.addAttribute("successMessage",
                    "Excel saved and Outlook draft opened. Review the email and click Send when ready.");
            model.addAttribute("excelPath", result.getExcelFile().toString());
            model.addAttribute("form", form);
        } catch (IllegalArgumentException ex) {
            // Validation errors from TaskParser (empty tasks, bad %, negative hours)
            model.addAttribute("errorMessage", ex.getMessage());
            model.addAttribute("form", form);
        } catch (IllegalStateException ex) {
            // Outlook not configured (missing mail.to)
            model.addAttribute("errorMessage", ex.getMessage());
            model.addAttribute("form", form);
        } catch (IOException ex) {
            // Excel write failed or Outlook/PowerShell script failed
            model.addAttribute("errorMessage",
                    "Could not generate report. Make sure Outlook is installed and try again.");
            model.addAttribute("form", form);
        }

        model.addAttribute("outputDir", ReportService.OUTPUT_DIR.toString());
        model.addAttribute("outlookConfigPath", OutlookConfig.getUserConfigPath().toString());
        return "index";
    }
}
