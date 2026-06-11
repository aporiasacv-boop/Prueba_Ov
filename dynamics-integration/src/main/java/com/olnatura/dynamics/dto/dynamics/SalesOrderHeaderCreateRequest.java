package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonProperty;

public record SalesOrderHeaderCreateRequest(
        @JsonProperty("dataAreaId") String dataAreaId,
        @JsonProperty("OrderingCustomerAccountNumber") String orderingCustomerAccountNumber,
        @JsonProperty("InvoiceCustomerAccountNumber") String invoiceCustomerAccountNumber,
        @JsonProperty("CurrencyCode") String currencyCode,
        @JsonProperty("SalesOrderName") String salesOrderName) {
}
