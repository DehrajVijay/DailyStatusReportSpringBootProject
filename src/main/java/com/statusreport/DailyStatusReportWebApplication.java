package com.statusreport;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * Entry point for the entire application.
 *
 * <p>When you run this class, Spring Boot:
 * <ol>
 *   <li>Starts an embedded Tomcat web server (port from application.properties)</li>
 *   <li>Scans this package and creates beans (@Controller, @Service, @Component)</li>
 *   <li>Wires dependencies together (constructor injection)</li>
 *   <li>Calls {@link BrowserLauncher} when the app is ready</li>
 * </ol>
 *
 * <p>End-to-end flow starts here, then goes to {@link ReportController}.
 */
@SpringBootApplication
public class DailyStatusReportWebApplication {

    /**
     * Launches Spring Boot. Equivalent to running {@code java -jar daily-status-report-web.jar}.
     */
    public static void main(String[] args) {
        SpringApplication.run(DailyStatusReportWebApplication.class, args);
    }
}
