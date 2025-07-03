package com.sterling.automation.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;
import java.util.regex.Pattern;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.sterling.automation.dto.ValidationResponse;

@Slf4j
public class ValidationService {

    private static final Pattern  ACTUALS_ACCOUNT_NAME_REGEX = Pattern.compile("^\\s*\\d{6}\\s.*");
    private static final int      BUDGET_LOT_NAME_ROW = 0;
    private static final int      BUDGET_LOT_NAME_COLUMN = 3;
    private static final String   HEADER_ROW_INDICATOR = "Distribution account";

    public ValidationResponse isValidInputOutput(final Scanner scanner, final String actuals, final String budget) {

        boolean isValid = false;

        Workbook budgetWorkbook = loadWorkbook(budget);
        String lotName = getLotName(budgetWorkbook);
        boolean isValidName = isValidName(lotName);

        if (!isValidName) {
            System.out.println("");
            System.out.println("Did you select the right budget file?");
            System.out.println("This script assumes the lot name is in cell D1 of the budget file.");
            System.out.print("Press ENTER to try again.");
            scanner.nextLine();
            System.out.println("");
            return ValidationResponse.builder().isValid(false).build();
        }

        Workbook actualsWorkbook = loadWorkbook(actuals);

        //boolean doActualsContainRequestedLot = doActualsContainRequestedLot(lotName, actualsWorkbook);


        return ValidationResponse.builder()
                    .lotName(lotName)
                    .input(actualsWorkbook)
                    .output(budgetWorkbook)
                    .isValid(true) //TODO:update
                    .columnIndexOfActuals(1) //TODO:update
                    .build();
    }

    private Workbook loadWorkbook(final String filePath) {

        File file = new File(filePath);

        log.info("filepath={} and File.exists() = {}", filePath, file.exists());

        try (InputStream inputStream = new FileInputStream(file)) {

            return WorkbookFactory.create(inputStream);

        } catch (IOException ioException) {
            log.error(ioException.getMessage());
            return null;
        }
    }

    private String getLotName(final Workbook workbook) {

        Cell cell = workbook.getSheetAt(0)
                    .getRow(BUDGET_LOT_NAME_ROW)
                    .getCell(BUDGET_LOT_NAME_COLUMN);

        if (cell == null) {
            return "";
        }

        return workbook.getSheetAt(0)
            .getRow(BUDGET_LOT_NAME_ROW)
            .getCell(BUDGET_LOT_NAME_COLUMN)
            .getStringCellValue();
    }

    private boolean isValidName(final String lotName) {
        return ACTUALS_ACCOUNT_NAME_REGEX.matcher(lotName).matches();
    }

    private boolean doActualsContainRequestedLot(final String lotName, final Workbook actuals) {

        boolean result = false;

        int headerRow = -1;

        for (Row row : actuals.getSheetAt(0)) {
            if (HEADER_ROW_INDICATOR.equals(row.getCell(0).getStringCellValue())) {
                headerRow = row.getRowNum();
            }
        }

        if (headerRow < 0) {
            System.err.println("No header found in actuals");
        }

        Row actualsLotNameRow = actuals.getSheetAt(0).getRow(headerRow);

        int columnOfActuals = -1; 

        for (Cell cell: actualsLotNameRow) {
            if (lotName.equals(cell.getStringCellValue())) {
                columnOfActuals = cell.getColumnIndex();
            }
        }

        if (columnOfActuals < 0) {
            System.err.println("No matching account found in actuals");
        }


        return result;
    }
}
