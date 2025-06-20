package com.sterling.automation.domain;

import java.math.BigDecimal;

import lombok.Builder;

@Builder
public record DistributionAccount(
    String id,
    String name,
    BigDecimal balance
) {}
