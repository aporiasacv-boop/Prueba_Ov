package com.olnatura.dynamics.dto;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class DynamicsHealthResponse {

    boolean success;
    String message;

    public static DynamicsHealthResponse ok(String message) {
        return DynamicsHealthResponse.builder().success(true).message(message).build();
    }

    public static DynamicsHealthResponse fail(String message) {
        return DynamicsHealthResponse.builder().success(false).message(message).build();
    }
}
