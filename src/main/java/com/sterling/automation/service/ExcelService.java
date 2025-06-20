package com.sterling.automation.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelService {

    private static final int LOT_NAME_SHEET = 0;
    private static final int LOT_NAME_ROW = 0;
    private static final int LOT_NAME_COLUMN = 3;

    public void process(final String budget, final String actuals) {

        System.out.println("Starting process");

        Workbook budgetWorkbook = loadWorkbook(budget);
        Workbook actualsWorkbook = loadWorkbook(actuals);

        String lotName = budgetWorkbook
                            .getSheetAt(LOT_NAME_SHEET)
                            .getRow(LOT_NAME_ROW)
                            .getCell(LOT_NAME_COLUMN)
                            .getStringCellValue();

        System.out.println(lotName);
    }

    private Workbook loadWorkbook(final String filePath) {
        try (InputStream file = new FileInputStream(filePath)) {
                
            return WorkbookFactory.create(file);

        } catch (IOException ioException) {
            System.err.println(ioException.getMessage());
            return null;
        }
    }
}
