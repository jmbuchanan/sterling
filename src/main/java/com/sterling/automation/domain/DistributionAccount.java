package com.sterling.automation.domain;

import lombok.Builder;

@Builder
public record DistributionAccount(
    String id,
    String name,
    double actuals,
    double budget,
    boolean existsInBudget,
    boolean existsInActuals
) {}
