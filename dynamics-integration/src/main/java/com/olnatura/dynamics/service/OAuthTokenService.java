package com.olnatura.dynamics.service;

import com.olnatura.dynamics.dto.AzureTokenResponse;
import com.olnatura.dynamics.dto.TokenStatusResponse;
import com.olnatura.dynamics.exception.OAuthTokenException;
import com.olnatura.dynamics.properties.AzureProperties;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.util.StringUtils;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

import java.time.Duration;
import java.time.Instant;

@Service
@RequiredArgsConstructor
public class OAuthTokenService {

    private static final Duration REFRESH_BUFFER = Duration.ofMinutes(5);

    private final RestTemplate restTemplate;
    private final AzureProperties azureProperties;

    private String accessToken;
    private Instant expirationTime;

    public synchronized String getAccessToken() {
        if (!StringUtils.hasText(accessToken) || expirationTime == null) {
            obtainNewToken();
            return accessToken;
        }

        if (isExpiringWithinBuffer()) {
            refreshToken();
            return accessToken;
        }

        return accessToken;
    }

    public synchronized TokenStatusResponse getTokenStatus() {
        boolean hasToken = StringUtils.hasText(accessToken) && expirationTime != null;
        if (!hasToken) {
            return TokenStatusResponse.builder()
                    .hasToken(false)
                    .expiresAt(null)
                    .secondsRemaining(0)
                    .build();
        }

        long secondsRemaining = Duration.between(Instant.now(), expirationTime).getSeconds();
        if (secondsRemaining < 0) {
            secondsRemaining = 0;
        }

        return TokenStatusResponse.builder()
                .hasToken(true)
                .expiresAt(expirationTime.toString())
                .secondsRemaining(secondsRemaining)
                .build();
    }

    public synchronized void forceRefresh() {
        this.accessToken = null;
        this.expirationTime = null;
        obtainNewToken();
    }

    @Scheduled(fixedRate = 300_000)
    public synchronized void scheduledTokenRefresh() {
        if (!StringUtils.hasText(accessToken) || expirationTime == null) {
            return;
        }
        if (isExpiringWithinBuffer()) {
            refreshToken();
        }
    }

    private void obtainNewToken() {
        fetchAndStoreToken();
    }

    private void refreshToken() {
        fetchAndStoreToken();
    }

    private void fetchAndStoreToken() {
        try {
            String clientSecret = azureProperties.getClientSecret();
            if (!StringUtils.hasText(clientSecret)) {
                throw new OAuthTokenException(
                        "client-secret vacío. Usa comillas en application.yml o AZURE_CLIENT_SECRET.");
            }

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_FORM_URLENCODED);

            MultiValueMap<String, String> body = new LinkedMultiValueMap<>();
            body.add("grant_type", "client_credentials");
            body.add("client_id", azureProperties.getClientId());
            body.add("client_secret", clientSecret);

            if (azureProperties.isTokenV1()) {
                body.add("resource", azureProperties.getResource());
            } else {
                body.add("scope", azureProperties.getScope());
            }

            HttpEntity<MultiValueMap<String, String>> request = new HttpEntity<>(body, headers);

            ResponseEntity<AzureTokenResponse> response = restTemplate.postForEntity(
                    azureProperties.tokenEndpoint(),
                    request,
                    AzureTokenResponse.class);

            AzureTokenResponse tokenResponse = response.getBody();
            validateTokenResponse(tokenResponse);

            this.accessToken = tokenResponse.getAccessToken().trim();
            this.expirationTime = Instant.now().plusSeconds(tokenResponse.getExpiresIn());

        } catch (HttpStatusCodeException ex) {
            throw new OAuthTokenException(
                    "OAuth HTTP error: " + ex.getStatusCode() + " - " + ex.getResponseBodyAsString(),
                    ex);
        } catch (RestClientException ex) {
            throw new OAuthTokenException("OAuth connection error: " + ex.getMessage(), ex);
        } catch (Exception ex) {
            throw new OAuthTokenException("OAuth unexpected error: " + ex.getMessage(), ex);
        }
    }

    private void validateTokenResponse(AzureTokenResponse tokenResponse) {
        if (tokenResponse == null) {
            throw new OAuthTokenException("OAuth response body is empty");
        }
        if (!StringUtils.hasText(tokenResponse.getAccessToken())) {
            throw new OAuthTokenException("OAuth access_token is empty");
        }
        if (tokenResponse.getExpiresIn() == null || tokenResponse.getExpiresIn() <= 0) {
            throw new OAuthTokenException(
                    "OAuth expires_in is invalid: " + tokenResponse.getExpiresIn());
        }
    }

    private boolean isExpiringWithinBuffer() {
        if (expirationTime == null) {
            return true;
        }
        Instant refreshThreshold = expirationTime.minus(REFRESH_BUFFER);
        return Instant.now().isAfter(refreshThreshold) || Instant.now().equals(refreshThreshold);
    }
}
