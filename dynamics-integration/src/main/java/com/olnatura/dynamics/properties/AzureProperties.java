package com.olnatura.dynamics.properties;

import jakarta.annotation.PostConstruct;
import jakarta.validation.constraints.NotBlank;
import lombok.Getter;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.validation.annotation.Validated;

@Getter
@Setter
@Validated
@ConfigurationProperties(prefix = "azure")
public class AzureProperties {

    @NotBlank
    private String tenantId;

    @NotBlank
    private String clientId;

    @NotBlank
    private String clientSecret;

    @NotBlank
    private String tokenVersion = "v1";

    @NotBlank
    private String resource;

    private String scope;

    @PostConstruct
    void normalizeSecret() {
        if (clientSecret != null) {
            clientSecret = clientSecret.trim();
            if (clientSecret.startsWith("\"") && clientSecret.endsWith("\"")) {
                clientSecret = clientSecret.substring(1, clientSecret.length() - 1);
            }
        }
    }

    public boolean isTokenV1() {
        return "v1".equalsIgnoreCase(tokenVersion);
    }

    public String tokenEndpoint() {
        if (isTokenV1()) {
            return "https://login.microsoftonline.com/" + tenantId + "/oauth2/token";
        }
        return "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";
    }
}
