package com.olnatura.dynamics.service;

import com.olnatura.dynamics.client.DynamicsClient;
import com.olnatura.dynamics.dto.CrearLineasRequest;
import com.olnatura.dynamics.dto.CrearLineasResponse;
import com.olnatura.dynamics.dto.CrearPedidoRequest;
import com.olnatura.dynamics.dto.CrearPedidoResponse;
import com.olnatura.dynamics.dto.LineaPedidoRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderHeaderCreateRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderLineCreateRequest;
import com.olnatura.dynamics.exception.DynamicsApiException;
import com.olnatura.dynamics.properties.DynamicsProperties;
import java.util.ArrayList;
import java.util.List;
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
                request.getDescripcionPedido(),
                toODataDateTime(request.getFechaEnvioSolicitada()));

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
        List<String> advertencias = new ArrayList<>();

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
                    precioUnitario,
                    toODataDateTime(linea.getFechaEnvio()));

            var lineResult = dynamicsClient.createSalesOrderLineWithMeta(linePayload);
            String lineResponse = lineResult.body();

            if (precioUnitario != null) {
                try {
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
                            salesOrderNumber,
                            inventoryLotId,
                            precioUnitario,
                            linea.getCantidad(),
                            lineResult.entityPath());
                } catch (DynamicsApiException ex) {
                    log.warn(
                            "Linea OV {} articulo {} creada, pero PATCH de precio fallo: {}",
                            salesOrderNumber,
                            linea.getCodigoArticulo(),
                            ex.getMessage());
                    advertencias.add(
                            "Linea " + linea.getCodigoArticulo()
                                    + ": creada en Dynamics; precio puede requerir revision (PATCH "
                                    + ex.getHttpStatus() + ").");
                }
            } else {
                log.warn(
                        "Linea OV {} articulo {} sin precio (>0). Dynamics dejara Precio unitario en 0.",
                        salesOrderNumber,
                        linea.getCodigoArticulo());
            }

            creadas++;
        }

        String mensaje = "Se crearon " + creadas + " línea(s) en OV " + salesOrderNumber;
        if (!advertencias.isEmpty()) {
            mensaje += ". Avisos: " + String.join(" ", advertencias);
        }

        return CrearLineasResponse.builder()
                .success(true)
                .salesOrderNumber(salesOrderNumber)
                .lineasCreadas(creadas)
                .mensaje(mensaje)
                .build();
    }

    private String toODataDateTime(String fecha) {
        if (!StringUtils.hasText(fecha)) {
            return null;
        }
        String trimmed = fecha.trim();
        if (trimmed.contains("T")) {
            return trimmed;
        }
        return trimmed + "T12:00:00Z";
    }
}
