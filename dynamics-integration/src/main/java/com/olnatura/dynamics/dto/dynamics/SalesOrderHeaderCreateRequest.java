package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record SalesOrderHeaderCreateRequest(
        @JsonProperty("dataAreaId") String dataAreaId,
        @JsonProperty("OrderingCustomerAccountNumber") String orderingCustomerAccountNumber,
        @JsonProperty("InvoiceCustomerAccountNumber") String invoiceCustomerAccountNumber,
        @JsonProperty("CurrencyCode") String currencyCode,
        @JsonProperty("SalesOrderName") String salesOrderName,
        @JsonProperty("CustomersOrderReference") String customersOrderReference,
        @JsonProperty("DeliveryModeCode") String deliveryModeCode,
        @JsonProperty("RequestedShippingDate") String requestedShippingDate) {
}
