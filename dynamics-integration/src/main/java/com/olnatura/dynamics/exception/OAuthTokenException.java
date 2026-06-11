package com.olnatura.dynamics.exception;

public class OAuthTokenException extends RuntimeException {

    public OAuthTokenException(String message) {
        super(message);
    }

    public OAuthTokenException(String message, Throwable cause) {
        super(message, cause);
    }
}
