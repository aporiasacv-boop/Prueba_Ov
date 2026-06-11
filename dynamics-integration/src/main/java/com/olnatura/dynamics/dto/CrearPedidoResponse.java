package com.olnatura.dynamics.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CrearPedidoResponse {

    private boolean success;
    private String salesOrderNumber;
}
