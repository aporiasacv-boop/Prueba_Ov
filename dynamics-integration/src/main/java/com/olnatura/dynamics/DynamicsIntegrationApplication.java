package com.olnatura.dynamics;

import com.olnatura.dynamics.properties.AzureProperties;
import com.olnatura.dynamics.properties.DynamicsProperties;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
@EnableConfigurationProperties({AzureProperties.class, DynamicsProperties.class})
public class DynamicsIntegrationApplication {

    public static void main(String[] args) {
        SpringApplication.run(DynamicsIntegrationApplication.class, args);
    }
}
