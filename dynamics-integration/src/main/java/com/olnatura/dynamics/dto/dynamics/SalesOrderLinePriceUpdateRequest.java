package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record SalesOrderLinePriceUpdateRequest(
        @JsonProperty("SalesPrice") double precioUnitario,
        @JsonProperty("SalesPriceQuantity") double salesPriceQuantity,
        @JsonProperty("LineAmount") Double lineAmount) {

    public static SalesOrderLinePriceUpdateRequest of(double precioUnitario, double cantidad) {
        return new SalesOrderLinePriceUpdateRequest(precioUnitario, 1.0, precioUnitario * cantidad);
    }
}
