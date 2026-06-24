package com.olnatura.dynamics.dto.dynamics;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record SalesOrderHeaderPatchRequest(
        @JsonProperty("CustomerRequisitionNumber") String customerRequisitionNumber,
        @JsonProperty("CustomersOrderReference") String customersOrderReference,
        @JsonProperty("RequestedShippingDate") String requestedShippingDate,
        @JsonProperty("RequestedReceiptDate") String requestedReceiptDate) {

    public static SalesOrderHeaderPatchRequest ordenCliente(String numeroOc) {
        return postCreacion(numeroOc, null);
    }

    public static SalesOrderHeaderPatchRequest postCreacion(String numeroOc, String fechaRecepcion) {
        String oc = blankToNull(numeroOc);
        String recepcion = blankToNull(fechaRecepcion);
        return new SalesOrderHeaderPatchRequest(oc, oc, null, recepcion);
    }

    private static String blankToNull(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        return value.trim();
    }
}
