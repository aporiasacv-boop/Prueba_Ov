package com.olnatura.dynamics.controller;

import com.olnatura.dynamics.client.DynamicsClient;
import com.olnatura.dynamics.dto.DynamicsHealthResponse;
import com.olnatura.dynamics.service.OAuthTokenService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/health")
@RequiredArgsConstructor
public class HealthController {

    private final OAuthTokenService oAuthTokenService;
    private final DynamicsClient dynamicsClient;

    @GetMapping("/ping")
    public DynamicsHealthResponse ping() {
        return DynamicsHealthResponse.ok("Servidor Spring Boot accesible.");
    }

    @GetMapping("/dynamics")
    public DynamicsHealthResponse dynamics() {
        try {
            oAuthTokenService.getAccessToken();
            dynamicsClient.getCustomers();
            return DynamicsHealthResponse.ok("Conexion OK con Dynamics FO y token Azure valido.");
        } catch (Exception ex) {
            return DynamicsHealthResponse.fail(ex.getMessage());
        }
    }
}
