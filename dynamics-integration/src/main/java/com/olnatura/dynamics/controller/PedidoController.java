package com.olnatura.dynamics.controller;

import com.olnatura.dynamics.dto.CrearLineasRequest;
import com.olnatura.dynamics.dto.CrearLineasResponse;
import com.olnatura.dynamics.dto.CrearPedidoRequest;
import com.olnatura.dynamics.dto.CrearPedidoResponse;
import com.olnatura.dynamics.service.PedidoService;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/pedidos")
@RequiredArgsConstructor
public class PedidoController {

    private final PedidoService pedidoService;

    @PostMapping("/crear")
    @ResponseStatus(HttpStatus.CREATED)
    public CrearPedidoResponse crearPedido(@Valid @RequestBody CrearPedidoRequest request) {
        return pedidoService.crearPedido(request);
    }

    @PostMapping("/lineas")
    @ResponseStatus(HttpStatus.CREATED)
    public CrearLineasResponse crearLineas(@Valid @RequestBody CrearLineasRequest request) {
        return pedidoService.crearLineas(request);
    }
}
