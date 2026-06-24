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

    /** Nombre comercial del cliente (ej. COSTCO DE MEXICO). Si viene vacío, se consulta en Dynamics. */
    private String nombreCliente;

    private String fechaEnvioSolicitada;
    private String fechaRecepcionSolicitada;
}
