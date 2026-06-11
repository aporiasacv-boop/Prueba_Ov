package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonProperty;

public record SalesOrderLineCreateRequest(
        @JsonProperty("dataAreaId") String dataAreaId,
        @JsonProperty("SalesOrderNumber") String salesOrderNumber,
        @JsonProperty("ProductNumber") String productNumber,
        @JsonProperty("OrderedSalesQuantity") double orderedSalesQuantity,
        @JsonProperty("SalesUnitSymbol") String salesUnitSymbol,
        @JsonProperty("ShippingSiteId") String shippingSiteId,
        @JsonProperty("ShippingWarehouseId") String shippingWarehouseId) {
}
