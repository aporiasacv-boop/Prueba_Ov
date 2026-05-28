package com.olnatura.dynamics.dto;

import jakarta.validation.constraints.NotBlank;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CrearPedidoRequest {

    @NotBlank
    private String cliente;

    @NotBlank
    private String referenciaCliente;

    @NotBlank
    private String descripcionPedido;

    private String fechaEnvioSolicitada;
    private String fechaRecepcionSolicitada;
}
