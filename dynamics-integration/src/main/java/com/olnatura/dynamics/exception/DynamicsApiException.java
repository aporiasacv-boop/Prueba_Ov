package com.olnatura.dynamics.exception;

import lombok.Getter;
import org.springframework.web.client.HttpStatusCodeException;

@Getter
public class DynamicsApiException extends RuntimeException {

    private final int httpStatus;
    private final String responseBody;

    public DynamicsApiException(String message) {
        super(message);
        this.httpStatus = 502;
        this.responseBody = "";
    }

    public DynamicsApiException(String message, int httpStatus, String responseBody) {
        super(buildReadableMessage(message, httpStatus, responseBody));
        this.httpStatus = httpStatus;
        this.responseBody = responseBody != null ? responseBody : "";
    }

    public DynamicsApiException(String message, Throwable cause) {
        super(message, cause);
        this.httpStatus = 502;
        this.responseBody = "";
    }

    public static DynamicsApiException fromHttp(String url, HttpStatusCodeException ex) {
        String responseBody = ex.getResponseBodyAsString();
        String summary = "Dynamics HTTP " + ex.getStatusCode().value() + " en " + url;
        return new DynamicsApiException(summary, ex.getStatusCode().value(), responseBody, ex);
    }

    private static String buildReadableMessage(String message, int httpStatus, String responseBody) {
        StringBuilder sb = new StringBuilder(message);
        sb.append(" [HTTP ").append(httpStatus).append("]");
        if (responseBody != null && !responseBody.isBlank()) {
            sb.append(" ").append(responseBody);
        }
        return sb.toString();
    }

    private DynamicsApiException(String message, int httpStatus, String responseBody, Throwable cause) {
        super(buildReadableMessage(message, httpStatus, responseBody), cause);
        this.httpStatus = httpStatus;
        this.responseBody = responseBody != null ? responseBody : "";
    }
}
