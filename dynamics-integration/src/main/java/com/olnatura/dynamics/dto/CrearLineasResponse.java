package com.olnatura.dynamics.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CrearLineasResponse {

    private boolean success;
    private String salesOrderNumber;
    private int lineasCreadas;
    private String mensaje;
}
