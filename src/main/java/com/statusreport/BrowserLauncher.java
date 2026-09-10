package com.statusreport;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

import java.awt.Desktop;
import java.net.URI;

/**
 * Opens the default browser automatically when the app finishes starting.
 *
 * <p>
 * Listens for {@link ApplicationReadyEvent} — fired after Tomcat is up and
 * all beans are ready. Reads port from {@code server.port} in
 * application.properties.
 *
 * <p>
 * Disable with {@code app.browser.auto-open=false} in application.properties
 * or when running tests.
 */
@Component
public class BrowserLauncher {

    @Value("${app.browser.auto-open:true}")
    private boolean autoOpen;

    @Value("${server.port:8081}")
    private int serverPort;

    /**
     * Opens http://localhost:{port}/ in the user's default browser.
     * Fails silently if Desktop API is not available (e.g. headless server).
     */
    @EventListener(ApplicationReadyEvent.class)
    public void openBrowser() {
        if (!autoOpen || !Desktop.isDesktopSupported()) {
            return;
        }

        try {
            Desktop desktop = Desktop.getDesktop();
            if (desktop.isSupported(Desktop.Action.BROWSE)) {
                desktop.browse(new URI("http://localhost:" + serverPort + "/"));
            }
        } catch (Exception ex) {
            // Browser open is optional; app still works if user navigates manually.
        }
    }
}
