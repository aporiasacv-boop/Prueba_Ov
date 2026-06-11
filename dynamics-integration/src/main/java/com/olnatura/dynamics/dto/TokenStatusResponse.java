package com.olnatura.dynamics.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TokenStatusResponse {

    private boolean hasToken;
    private String expiresAt;
    private long secondsRemaining;
}
