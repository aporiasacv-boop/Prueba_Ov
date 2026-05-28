package com.olnatura.dynamics.dto;

import jakarta.validation.Valid;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotEmpty;
import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class CrearLineasRequest {

    @NotBlank
    private String salesOrderNumber;

    @NotEmpty
    @Valid
    private List<LineaPedidoRequest> lineas;
}
