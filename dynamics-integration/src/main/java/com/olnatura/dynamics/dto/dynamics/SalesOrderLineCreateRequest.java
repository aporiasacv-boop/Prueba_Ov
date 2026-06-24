package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record SalesOrderLineCreateRequest(
        @JsonProperty("dataAreaId") String dataAreaId,
        @JsonProperty("SalesOrderNumber") String salesOrderNumber,
        @JsonProperty("ProductNumber") String productNumber,
        @JsonProperty("OrderedSalesQuantity") double orderedSalesQuantity,
        @JsonProperty("SalesUnitSymbol") String salesUnitSymbol,
        @JsonProperty("ShippingSiteId") String shippingSiteId,
        @JsonProperty("ShippingWarehouseId") String shippingWarehouseId,
        @JsonProperty("DeliveryModeCode") String deliveryModeCode,
        @JsonProperty("SalesPrice") Double precioUnitario,
        @JsonProperty("RequestedShippingDate") String requestedShippingDate) {
}
