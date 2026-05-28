package com.olnatura.dynamics.dto;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Positive;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class LineaPedidoRequest {

    @NotBlank
    private String codigoArticulo;

    @Positive
    private double cantidad;

    private double precioUnitario;
    private double porcentajeDescuento;
    private String fechaEnvio;
    private String comentario;
}
