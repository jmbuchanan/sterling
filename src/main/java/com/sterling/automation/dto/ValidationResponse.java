package com.sterling.automation.dto;

import org.apache.poi.ss.usermodel.Workbook;

import lombok.Builder;

@Builder
public record ValidationResponse (
    Workbook input,
    Workbook output,
    boolean isValid,
    int columnIndexOfActuals,
    String lotName
){}

    
