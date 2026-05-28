package com.olnatura.dynamics.exception;

import jakarta.servlet.http.HttpServletRequest;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestClientException;

import java.time.Instant;

@Slf4j
@RestControllerAdvice
public class GlobalExceptionHandler {

    @ExceptionHandler(IllegalArgumentException.class)
    public ResponseEntity<ErrorResponse> handleIllegalArgument(
            IllegalArgumentException ex,
            HttpServletRequest request) {
        log.warn("Validación de negocio: {}", ex.getMessage());
        return buildResponse(HttpStatus.BAD_REQUEST, "Validation Error", ex.getMessage(), request);
    }

    @ExceptionHandler(MethodArgumentNotValidException.class)
    public ResponseEntity<ErrorResponse> handleValidation(
            MethodArgumentNotValidException ex,
            HttpServletRequest request) {
        String message = ex.getBindingResult().getFieldErrors().stream()
                .map(err -> err.getField() + ": " + err.getDefaultMessage())
                .reduce((a, b) -> a + "; " + b)
                .orElse("Validación fallida");
        log.warn("Validation error: {}", message);
        return buildResponse(HttpStatus.BAD_REQUEST, "Validation Error", message, request);
    }

    @ExceptionHandler(OAuthTokenException.class)
    public ResponseEntity<ErrorResponse> handleOAuthToken(
            OAuthTokenException ex,
            HttpServletRequest request) {
        log.error("OAuth token error: {}", ex.getMessage(), ex);
        return buildResponse(HttpStatus.BAD_GATEWAY, "OAuth Error", ex.getMessage(), request);
    }

    @ExceptionHandler(DynamicsApiException.class)
    public ResponseEntity<ErrorResponse> handleDynamicsApi(
            DynamicsApiException ex,
            HttpServletRequest request) {
        log.error("Dynamics API error: {}", ex.getMessage());
        HttpStatus status = resolveHttpStatus(ex.getHttpStatus());
        return buildResponse(status, "Dynamics API Error", ex.getMessage(), request);
    }

    @ExceptionHandler(HttpStatusCodeException.class)
    public ResponseEntity<ErrorResponse> handleHttpStatus(
            HttpStatusCodeException ex,
            HttpServletRequest request) {
        log.error("HTTP error {}: {}", ex.getStatusCode(), ex.getResponseBodyAsString());
        String message = "HTTP " + ex.getStatusCode().value() + ": " + ex.getResponseBodyAsString();
        HttpStatus status = HttpStatus.valueOf(ex.getStatusCode().value());
        return buildResponse(status, "HTTP Error", message, request);
    }

    @ExceptionHandler(RestClientException.class)
    public ResponseEntity<ErrorResponse> handleRestClient(
            RestClientException ex,
            HttpServletRequest request) {
        log.error("Rest client error: {}", ex.getMessage(), ex);
        return buildResponse(
                HttpStatus.SERVICE_UNAVAILABLE,
                "Connection Error",
                ex.getMessage(),
                request);
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<ErrorResponse> handleGeneric(
            Exception ex,
            HttpServletRequest request) {
        log.error("Unexpected error: {}", ex.getMessage(), ex);
        String message = ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName();
        return buildResponse(
                HttpStatus.INTERNAL_SERVER_ERROR,
                "Internal Error",
                message,
                request);
    }

    private HttpStatus resolveHttpStatus(int httpStatus) {
        try {
            return HttpStatus.valueOf(httpStatus);
        } catch (IllegalArgumentException ex) {
            return HttpStatus.BAD_GATEWAY;
        }
    }

    private ResponseEntity<ErrorResponse> buildResponse(
            HttpStatus status,
            String error,
            String message,
            HttpServletRequest request) {
        ErrorResponse body = ErrorResponse.builder()
                .timestamp(Instant.now())
                .status(status.value())
                .error(error)
                .message(message)
                .path(request.getRequestURI())
                .build();
        return ResponseEntity.status(status).body(body);
    }
}
