package com.olnatura.dynamics.controller;

import com.olnatura.dynamics.dto.TokenStatusResponse;
import com.olnatura.dynamics.service.OAuthTokenService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/token")
@RequiredArgsConstructor
public class TokenController {

    private final OAuthTokenService oAuthTokenService;

    @GetMapping("/status")
    public TokenStatusResponse status() {
        return oAuthTokenService.getTokenStatus();
    }

    @PostMapping("/refresh")
    public TokenStatusResponse refresh() {
        oAuthTokenService.forceRefresh();
        return oAuthTokenService.getTokenStatus();
    }
}
