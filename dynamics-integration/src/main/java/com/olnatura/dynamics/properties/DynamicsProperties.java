package com.olnatura.dynamics.properties;

import jakarta.validation.constraints.NotBlank;
import lombok.Getter;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.validation.annotation.Validated;

@Getter
@Setter
@Validated
@ConfigurationProperties(prefix = "dynamics")
public class DynamicsProperties {

    @NotBlank
    private String baseUrl;

    @NotBlank
    private String dataAreaId = "olna";

    @NotBlank
    private String defaultCurrency = "MXN";

    private LineDefaults line = new LineDefaults();

    public String odataUrl(String entityPath) {
        String base = baseUrl.endsWith("/") ? baseUrl.substring(0, baseUrl.length() - 1) : baseUrl;
        String path = entityPath.startsWith("/") ? entityPath : "/" + entityPath;
        return base + path;
    }

    @Getter
    @Setter
    public static class LineDefaults {

        private String salesUnitSymbol = "PZA";

        @NotBlank
        private String shippingSiteId = "OLNATURA";

        @NotBlank
        private String shippingWarehouseId = "PTM";
    }
}
