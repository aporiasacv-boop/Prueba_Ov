package com.olnatura.dynamics.service;

import com.olnatura.dynamics.client.DynamicsClient;
import com.olnatura.dynamics.dto.CrearLineasRequest;
import com.olnatura.dynamics.dto.CrearLineasResponse;
import com.olnatura.dynamics.dto.CrearPedidoRequest;
import com.olnatura.dynamics.dto.CrearPedidoResponse;
import com.olnatura.dynamics.dto.LineaPedidoRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderHeaderCreateRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderLineCreateRequest;
import com.olnatura.dynamics.properties.DynamicsProperties;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

@Slf4j
@Service
@RequiredArgsConstructor
public class PedidoService {

    private final OAuthTokenService oAuthTokenService;
    private final DynamicsClient dynamicsClient;
    private final DynamicsProperties dynamicsProperties;

    public CrearPedidoResponse crearPedido(CrearPedidoRequest request) {
        oAuthTokenService.getAccessToken();

        SalesOrderHeaderCreateRequest header = new SalesOrderHeaderCreateRequest(
                dynamicsProperties.getDataAreaId(),
                request.getCliente(),
                request.getCliente(),
                dynamicsProperties.getDefaultCurrency(),
                request.getDescripcionPedido());

        String headerResponse = dynamicsClient.createSalesOrderHeader(header);
        String salesOrderNumber = dynamicsClient.extractSalesOrderNumber(headerResponse);

        return CrearPedidoResponse.builder()
                .success(true)
                .salesOrderNumber(salesOrderNumber)
                .build();
    }

    public CrearLineasResponse crearLineas(CrearLineasRequest request) {
        if (!StringUtils.hasText(request.getSalesOrderNumber())) {
            throw new IllegalArgumentException("Debe crear primero la orden de venta");
        }

        String salesOrderNumber = request.getSalesOrderNumber().trim();
        oAuthTokenService.getAccessToken();

        DynamicsProperties.LineDefaults lineDefaults = dynamicsProperties.getLine();
        int creadas = 0;

        for (LineaPedidoRequest linea : request.getLineas()) {
            double precio = linea.getPrecioUnitario();
            Double precioUnitario = precio > 0 ? precio : null;

            log.info(
                    "Linea OV {} articulo {} cantidad {} precioUnitario {}",
                    salesOrderNumber,
                    linea.getCodigoArticulo(),
                    linea.getCantidad(),
                    precioUnitario);

            SalesOrderLineCreateRequest linePayload = new SalesOrderLineCreateRequest(
                    dynamicsProperties.getDataAreaId(),
                    salesOrderNumber,
                    linea.getCodigoArticulo(),
                    linea.getCantidad(),
                    lineDefaults.getSalesUnitSymbol(),
                    lineDefaults.getShippingSiteId(),
                    lineDefaults.getShippingWarehouseId(),
                    precioUnitario);

            String lineResponse = dynamicsClient.createSalesOrderLine(linePayload);

            if (precioUnitario != null) {
                String inventoryLotId = dynamicsClient.resolveInventoryLotId(
                        lineResponse,
                        dynamicsProperties.getDataAreaId(),
                        salesOrderNumber,
                        linea.getCodigoArticulo());
                log.info(
                        "Actualizando SalesPrice (Precio unitario) {} en lote {}",
                        precioUnitario,
                        inventoryLotId);
                dynamicsClient.updateSalesOrderLinePrice(
                        dynamicsProperties.getDataAreaId(),
                        inventoryLotId,
                        precioUnitario,
                        linea.getCantidad());
            } else {
                log.warn(
                        "Linea OV {} articulo {} sin precio (>0). Dynamics dejara Precio unitario en 0.",
                        salesOrderNumber,
                        linea.getCodigoArticulo());
            }

            creadas++;
        }

        String mensaje = "Se crearon " + creadas + " línea(s) en OV " + salesOrderNumber;

        return CrearLineasResponse.builder()
                .success(true)
                .salesOrderNumber(salesOrderNumber)
                .lineasCreadas(creadas)
                .mensaje(mensaje)
                .build();
    }
}
