package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonProperty;

public record SalesOrderLinePriceUpdateRequest(
        @JsonProperty("SalesPrice") double precioUnitario,
        @JsonProperty("SalesPriceQuantity") double salesPriceQuantity) {

    public static SalesOrderLinePriceUpdateRequest of(double precioUnitario) {
        return new SalesOrderLinePriceUpdateRequest(precioUnitario, 1.0);
    }
}
