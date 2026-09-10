package com.statusreport.config;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.boot.context.event.ApplicationStartedEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

@Component
public class ApplicationStartupLogger {

    private static final Logger log = LoggerFactory.getLogger(ApplicationStartupLogger.class);

    @EventListener
    public void onApplicationStarted(ApplicationStartedEvent event) {
        log.info("Application started successfully");
    }

    @EventListener
    public void onApplicationReady(ApplicationReadyEvent event) {
        log.info("Application is ready to accept requests");
    }
}