package com.olnatura.dynamics.client;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.olnatura.dynamics.dto.dynamics.SalesOrderHeaderCreateRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderLineCreateRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderLinePriceUpdateRequest;
import com.olnatura.dynamics.exception.DynamicsApiException;
import com.olnatura.dynamics.properties.DynamicsProperties;
import com.olnatura.dynamics.service.OAuthTokenService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Component
@RequiredArgsConstructor
public class DynamicsClient {

    private static final String CUSTOMERS_PATH = "/data/Customers";
    private static final String D365_SALES_ORDER_HEADERS = "/data/D365SalesOrderHeaders";
    private static final String D365_SALES_ORDER_LINES = "/data/D365SalesOrderLines";
    private static final Pattern SALES_ORDER_NUMBER_PATTERN = Pattern.compile(
            "SalesOrderNumber['\"]?(?:=|%3D)['\"]?([A-Za-z0-9]+)",
            Pattern.CASE_INSENSITIVE);
    private static final Pattern OV_NUMBER_PATTERN = Pattern.compile(
            "(OV\\d{4,})",
            Pattern.CASE_INSENSITIVE);

    private static final Pattern INVENTORY_LOT_ID_PATTERN = Pattern.compile(
            "InventoryLotId['\"]?(?:=|%3D)['\"]?([^'\"&\\s,)]+)",
            Pattern.CASE_INSENSITIVE);

    private final RestTemplate restTemplate;
    private final OAuthTokenService oAuthTokenService;
    private final DynamicsProperties dynamicsProperties;
    private final ObjectMapper objectMapper;

    public String getCustomers() {
        return getODataEntity(CUSTOMERS_PATH);
    }

    public String createSalesOrderHeader(SalesOrderHeaderCreateRequest request) {
        return postODataEntity(D365_SALES_ORDER_HEADERS, request);
    }

    public String createSalesOrderLine(SalesOrderLineCreateRequest request) {
        return postODataEntity(D365_SALES_ORDER_LINES, request);
    }

    public void updateSalesOrderLinePrice(
            String dataAreaId,
            String inventoryLotId,
            double precioUnitario,
            double cantidad) {
        String entityKey = D365_SALES_ORDER_LINES
                + "(dataAreaId='" + dataAreaId + "',InventoryLotId='" + inventoryLotId + "')";
        patchODataEntity(entityKey, SalesOrderLinePriceUpdateRequest.of(precioUnitario, cantidad));
    }

    public String resolveInventoryLotId(
            String createLineResponseJson,
            String dataAreaId,
            String salesOrderNumber,
            String productNumber) {
        try {
            return extractInventoryLotId(createLineResponseJson);
        } catch (DynamicsApiException ex) {
            String fromQuery = findInventoryLotId(dataAreaId, salesOrderNumber, productNumber);
            if (fromQuery != null) {
                return fromQuery;
            }
            throw ex;
        }
    }

    public String findInventoryLotId(String dataAreaId, String salesOrderNumber, String productNumber) {
        String filter = String.format(
                "$filter=SalesOrderNumber eq '%s' and ProductNumber eq '%s' and dataAreaId eq '%s'"
                        + "&$select=InventoryLotId&$top=1",
                salesOrderNumber.replace("'", "''"),
                productNumber.replace("'", "''"),
                dataAreaId.replace("'", "''"));
        try {
            String response = getODataEntity(D365_SALES_ORDER_LINES + "?" + filter);
            JsonNode root = objectMapper.readTree(response);
            JsonNode value = root.get("value");
            if (value == null || !value.isArray() || value.isEmpty()) {
                return null;
            }
            JsonNode lotNode = value.get(0).get("InventoryLotId");
            if (lotNode == null || lotNode.isNull() || lotNode.asText().isBlank()) {
                return null;
            }
            return lotNode.asText();
        } catch (Exception ex) {
            throw new DynamicsApiException(
                    "No se pudo consultar InventoryLotId para OV " + salesOrderNumber
                            + " producto " + productNumber + ": " + ex.getMessage(),
                    ex);
        }
    }

    public String extractInventoryLotId(String createLineResponseJson) {
        try {
            JsonNode root = objectMapper.readTree(createLineResponseJson);
            JsonNode lotNode = root.get("InventoryLotId");
            if (lotNode != null && !lotNode.isNull() && !lotNode.asText().isBlank()) {
                return lotNode.asText();
            }

            String fromText = extractInventoryLotIdFromText(createLineResponseJson);
            if (fromText != null) {
                return fromText;
            }

            throw new DynamicsApiException(
                    "InventoryLotId no viene en la respuesta de Dynamics",
                    502,
                    createLineResponseJson);
        } catch (DynamicsApiException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new DynamicsApiException(
                    "No se pudo leer InventoryLotId de la respuesta OData: " + ex.getMessage(),
                    ex);
        }
    }

    public String extractSalesOrderNumber(String createHeaderResponseJson) {
        try {
            JsonNode root = objectMapper.readTree(createHeaderResponseJson);
            JsonNode numberNode = root.get("SalesOrderNumber");
            if (numberNode != null && !numberNode.isNull() && !numberNode.asText().isBlank()) {
                return numberNode.asText();
            }

            String fromText = extractSalesOrderNumberFromText(createHeaderResponseJson);
            if (fromText != null) {
                return fromText;
            }

            throw new DynamicsApiException(
                    "SalesOrderNumber no viene en la respuesta de Dynamics",
                    502,
                    createHeaderResponseJson);
        } catch (DynamicsApiException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new DynamicsApiException(
                    "No se pudo leer SalesOrderNumber de la respuesta OData: " + ex.getMessage(),
                    ex);
        }
    }

    private String getODataEntity(String entityPath) {
        String url = dynamicsProperties.odataUrl(entityPath);
        HttpHeaders headers = buildJsonHeaders();
        HttpEntity<Void> request = new HttpEntity<>(headers);

        try {
            ResponseEntity<String> response = restTemplate.exchange(
                    url, HttpMethod.GET, request, String.class);
            return requireBody(response.getBody(), url);
        } catch (HttpStatusCodeException ex) {
            throw DynamicsApiException.fromHttp(url, ex);
        } catch (RestClientException ex) {
            throw new DynamicsApiException("Dynamics connection error for " + url + ": " + ex.getMessage(), ex);
        }
    }

    private String postODataEntity(String entityPath, Object body) {
        String url = dynamicsProperties.odataUrl(entityPath);

        try {
            String jsonBody = objectMapper.writeValueAsString(body);
            HttpHeaders headers = buildJsonHeaders();
            HttpEntity<String> request = new HttpEntity<>(jsonBody, headers);

            ResponseEntity<String> response = restTemplate.exchange(
                    url, HttpMethod.POST, request, String.class);

            String responseBody = response.getBody();
            if (responseBody == null || responseBody.isBlank()) {
                return buildJsonFromEntityHeaders(response.getHeaders());
            }

            return mergeResponseWithEntityHeaders(responseBody, response.getHeaders());

        } catch (HttpStatusCodeException ex) {
            throw DynamicsApiException.fromHttp(url, ex);
        } catch (RestClientException ex) {
            throw new DynamicsApiException("Dynamics connection error for " + url + ": " + ex.getMessage(), ex);
        } catch (Exception ex) {
            throw new DynamicsApiException("Dynamics POST error for " + url + ": " + ex.getMessage(), ex);
        }
    }

    private void patchODataEntity(String entityPathWithKey, Object body) {
        String url = dynamicsProperties.odataUrl(entityPathWithKey);

        try {
            String jsonBody = objectMapper.writeValueAsString(body);
            HttpHeaders headers = buildJsonHeaders();
            headers.add("If-Match", "*");
            HttpEntity<String> request = new HttpEntity<>(jsonBody, headers);

            restTemplate.exchange(url, HttpMethod.PATCH, request, String.class);

        } catch (HttpStatusCodeException ex) {
            throw DynamicsApiException.fromHttp(url, ex);
        } catch (RestClientException ex) {
            throw new DynamicsApiException("Dynamics connection error for " + url + ": " + ex.getMessage(), ex);
        } catch (Exception ex) {
            throw new DynamicsApiException("Dynamics PATCH error for " + url + ": " + ex.getMessage(), ex);
        }
    }

    private HttpHeaders buildJsonHeaders() {
        HttpHeaders headers = new HttpHeaders();
        headers.setBearerAuth(oAuthTokenService.getAccessToken());
        headers.setContentType(MediaType.APPLICATION_JSON);
        headers.setAccept(List.of(MediaType.APPLICATION_JSON));
        headers.add("OData-Version", "4.0");
        headers.add("OData-MaxVersion", "4.0");
        return headers;
    }

    private String requireBody(String body, String url) {
        if (body == null || body.isBlank()) {
            throw new DynamicsApiException("Dynamics response body is empty for " + url);
        }
        return body;
    }

    private String extractSalesOrderNumberFromHeaders(HttpHeaders headers) {
        String entityId = entityIdFromHeaders(headers);
        if (entityId == null) {
            return null;
        }
        return extractSalesOrderNumberFromText(entityId);
    }

    private String extractInventoryLotIdFromHeaders(HttpHeaders headers) {
        String entityId = entityIdFromHeaders(headers);
        if (entityId == null) {
            return null;
        }
        return extractInventoryLotIdFromText(entityId);
    }

    private String entityIdFromHeaders(HttpHeaders headers) {
        String entityId = headers.getFirst("OData-EntityId");
        if (entityId == null || entityId.isBlank()) {
            entityId = headers.getFirst("Location");
        }
        if (entityId == null || entityId.isBlank()) {
            return null;
        }
        return entityId;
    }

    private String mergeResponseWithEntityHeaders(String responseBody, HttpHeaders headers) {
        try {
            JsonNode root = objectMapper.readTree(responseBody);
            if (!root.isObject()) {
                return responseBody;
            }

            com.fasterxml.jackson.databind.node.ObjectNode merged = (com.fasterxml.jackson.databind.node.ObjectNode) root;
            boolean changed = false;

            String salesOrderNumber = extractSalesOrderNumberFromHeaders(headers);
            if (salesOrderNumber != null && !merged.hasNonNull("SalesOrderNumber")) {
                merged.put("SalesOrderNumber", salesOrderNumber);
                changed = true;
            }

            String inventoryLotId = extractInventoryLotIdFromHeaders(headers);
            if (inventoryLotId != null && !merged.hasNonNull("InventoryLotId")) {
                merged.put("InventoryLotId", inventoryLotId);
                changed = true;
            }

            return changed ? objectMapper.writeValueAsString(merged) : responseBody;
        } catch (Exception ex) {
            String fromBody = extractSalesOrderNumberFromText(responseBody);
            if (fromBody != null && !responseBody.contains("\"SalesOrderNumber\"")) {
                return "{\"SalesOrderNumber\":\"" + fromBody + "\"}";
            }
            return responseBody;
        }
    }

    private String buildJsonFromEntityHeaders(HttpHeaders headers) {
        String salesOrderNumber = extractSalesOrderNumberFromHeaders(headers);
        String inventoryLotId = extractInventoryLotIdFromHeaders(headers);

        StringBuilder json = new StringBuilder("{");
        boolean hasField = false;

        if (salesOrderNumber != null) {
            json.append("\"SalesOrderNumber\":\"").append(salesOrderNumber).append("\"");
            hasField = true;
        }
        if (inventoryLotId != null) {
            if (hasField) {
                json.append(",");
            }
            json.append("\"InventoryLotId\":\"").append(inventoryLotId).append("\"");
            hasField = true;
        }

        json.append("}");
        return hasField ? json.toString() : "{}";
    }

    private String extractSalesOrderNumberFromText(String text) {
        if (text == null || text.isBlank()) {
            return null;
        }
        Matcher matcher = SALES_ORDER_NUMBER_PATTERN.matcher(text);
        if (matcher.find()) {
            return matcher.group(1).toUpperCase();
        }
        Matcher ovMatcher = OV_NUMBER_PATTERN.matcher(text);
        if (ovMatcher.find()) {
            return ovMatcher.group(1).toUpperCase();
        }
        return null;
    }

    private String extractInventoryLotIdFromText(String text) {
        if (text == null || text.isBlank()) {
            return null;
        }
        Matcher matcher = INVENTORY_LOT_ID_PATTERN.matcher(text);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return null;
    }
}
