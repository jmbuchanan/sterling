package com.sterling.automation;

import com.sterling.automation.service.ExcelService;

public class App {

    public static void main( String[] args ) {

        ExcelService excelService = new ExcelService();

        excelService.process("budget.xlsx", "actuals-multiple.xlsx");

    }
}
