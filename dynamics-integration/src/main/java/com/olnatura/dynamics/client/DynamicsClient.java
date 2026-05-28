package com.olnatura.dynamics.client;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.olnatura.dynamics.dto.dynamics.SalesOrderHeaderCreateRequest;
import com.olnatura.dynamics.dto.dynamics.SalesOrderLineCreateRequest;
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

@Component
@RequiredArgsConstructor
public class DynamicsClient {

    private static final String CUSTOMERS_PATH = "/data/Customers";
    private static final String D365_SALES_ORDER_HEADERS = "/data/D365SalesOrderHeaders";
    private static final String D365_SALES_ORDER_LINES = "/data/D365SalesOrderLines";

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

    public String extractSalesOrderNumber(String createHeaderResponseJson) {
        try {
            JsonNode root = objectMapper.readTree(createHeaderResponseJson);
            JsonNode numberNode = root.get("SalesOrderNumber");
            if (numberNode == null || numberNode.isNull()) {
                throw new DynamicsApiException(
                        "SalesOrderNumber no viene en la respuesta de Dynamics",
                        502,
                        createHeaderResponseJson);
            }
            return numberNode.asText();
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
                return "{}";
            }
            return responseBody;

        } catch (HttpStatusCodeException ex) {
            throw DynamicsApiException.fromHttp(url, ex);
        } catch (RestClientException ex) {
            throw new DynamicsApiException("Dynamics connection error for " + url + ": " + ex.getMessage(), ex);
        } catch (Exception ex) {
            throw new DynamicsApiException("Dynamics POST error for " + url + ": " + ex.getMessage(), ex);
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
}
