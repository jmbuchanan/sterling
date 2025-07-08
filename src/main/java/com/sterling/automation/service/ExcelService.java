package com.sterling.automation.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Pattern;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import com.sterling.automation.domain.DistributionAccount;
import com.sterling.automation.dto.ValidationResponse;

@Slf4j
public class ExcelService {

    //matches: "   123456 EXAMPLE ACCOUNT  "
    private static final Pattern ACTUALS_ACCOUNT_NAME_REGEX = Pattern.compile("^\\s*\\d{6}\\s.*");
    private static final Pattern BUDGET_ACCOUNT_NAME_REGEX = Pattern.compile("^\\d{6}$");

    public void process(final ValidationResponse response) {

        Sheet actualsSheet = response.input().getSheetAt(0);
        Sheet budgetSheet = response.output().getSheetAt(0);

        List<DistributionAccount> actualsAccounts = 
            getAccountsFromActuals(
                actualsSheet, 
                response.lotName(), 
                response.columnIndexOfActuals());

        ImmutablePair<Integer, List<DistributionAccount>> budgetInfo = addBudgetInfo(budgetSheet, actualsAccounts);
        int firstRowOfBudget = budgetInfo.getLeft();
        List<DistributionAccount> consolidatedAccounts = budgetInfo.getRight();

        Workbook budgetWb = response.output();

        CellStyle firstColumnCellStyle = budgetSheet.getRow(firstRowOfBudget).getCell(0).getCellStyle();

        CellStyle missingFromActualsCellStyle = budgetWb.createCellStyle();
        missingFromActualsCellStyle.cloneStyleFrom(firstColumnCellStyle);
        missingFromActualsCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        missingFromActualsCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());

        CellStyle missingFromBudgetCellStyle = budgetWb.createCellStyle();
        missingFromBudgetCellStyle.cloneStyleFrom(firstColumnCellStyle);
        missingFromBudgetCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        missingFromBudgetCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

        for (int i = 0; i < consolidatedAccounts.size(); i++) {

            int rowIndex = firstRowOfBudget + i;
            boolean shiftAfterInsert = false;

            Row row = budgetSheet.getRow(rowIndex);

            if (row == null) {
                row = budgetSheet.createRow(rowIndex);
            }

            if (row.getCell(0) == null) {
                shiftAfterInsert = true;
            } else {

                String firstColumnValue = null;


                if (row.getCell(0).getCellType().equals(CellType.NUMERIC)) {
                    firstColumnValue = String.format("%.0f", row.getCell(0).getNumericCellValue());
                }

                if (row.getCell(0).getCellType().equals(CellType.STRING)) {
                    firstColumnValue = String.valueOf(row.getCell(0).getStringCellValue());
                }

                if (!BUDGET_ACCOUNT_NAME_REGEX.matcher(firstColumnValue).matches()) {
                    budgetSheet.shiftRows(rowIndex, budgetSheet.getLastRowNum(), 1);
                    row = budgetSheet.createRow(rowIndex);
                }
            }


            DistributionAccount account = consolidatedAccounts.get(i);

            row.createCell(0).setCellValue(Float.parseFloat(account.id()));
            row.createCell(1).setCellValue(account.name());
            row.createCell(2).setCellValue(account.budget());
            row.createCell(3).setCellValue(account.actuals());
            row.createCell(4).setCellFormula(String.format("$C%d-$D%d", rowIndex + 1, rowIndex + 1));

            if (account.existsInBudget() && !account.existsInActuals()) {
                row.getCell(0).setCellStyle(missingFromActualsCellStyle);
            }

            if (account.existsInActuals() && !account.existsInBudget()) {
                row.getCell(0).setCellStyle(missingFromBudgetCellStyle);
            }

            if (shiftAfterInsert) {
                budgetSheet.shiftRows(rowIndex + 1, budgetSheet.getLastRowNum(), 1);
            }
        }

        int i = 0;
        boolean keepGoing = true;
        int lastRowOfBudget = firstRowOfBudget + consolidatedAccounts.size();
        int costsRow = 0;

        do {

            Row row = budgetSheet.getRow(lastRowOfBudget + i);
            i++;

            if (row != null &&
                row.getCell(0) != null && 
                row.getCell(0).getCellType() != null &&
                row.getCell(0).getCellType().equals(CellType.STRING)) {

                String firstColumnValue = String.valueOf(row.getCell(0).getStringCellValue());

                if (firstColumnValue.trim().equals("Total Costs")) {
                    costsRow = row.getRowNum();
                    row.getCell(2).setCellFormula(String.format("SUM(C%d:C%d)", firstRowOfBudget + 1, lastRowOfBudget));
                    row.getCell(3).setCellFormula(String.format("SUM(D%d:D%d)", firstRowOfBudget + 1, lastRowOfBudget));
                }

                if (firstColumnValue.trim().equals("Profit")) {
                    row.getCell(2).setCellFormula(String.format("SUM(C%d:C%d) - C%d", lastRowOfBudget + 1, costsRow, costsRow + 1));
                    row.getCell(3).setCellFormula(String.format("SUM(D%d:D%d) - D%d", lastRowOfBudget + 1, costsRow, costsRow + 1));
                    keepGoing = false;
                }
            }
        } while (keepGoing && i < 20);

        if (i >= 20) {
            log.error("Couldn't find profit row");
        }

        XSSFFormulaEvaluator.evaluateAllFormulaCells(budgetWb);

        File budgetFile = new File(response.budgetPath());


        try (OutputStream outputStream = new FileOutputStream(budgetFile)) {
            budgetWb.write(outputStream);
            budgetWb.close();
            System.out.println("Done.");
        } catch (Exception e) {
            log.error("Couldn't save budget file for some reason.");
            log.error(e.getMessage());
        }
    }

    private List<DistributionAccount> getAccountsFromActuals(final Sheet actualsSheet, final String lotName, final int columnOfActuals) {

        List<DistributionAccount> actualsAccounts = new ArrayList<>();

        for (Row row: actualsSheet) {

            if (row.getCell(0) == null) {
                continue;
            }

            String firstColumnValue = row.getCell(0).getStringCellValue();

            if (ACTUALS_ACCOUNT_NAME_REGEX.matcher(firstColumnValue).matches()) {

                int delimiterIndex = firstColumnValue.indexOf(" ");

                String id = firstColumnValue.substring(0, delimiterIndex);
                String name = firstColumnValue.substring(delimiterIndex + 1);
                double actuals = row.getCell(columnOfActuals).getNumericCellValue();

                DistributionAccount account = 
                    DistributionAccount.builder()
                        .id(id)
                        .name(name)
                        .actuals(actuals)
                        .existsInActuals(true)
                        .build();

                actualsAccounts.add(account);
            }
        }

        return actualsAccounts;
    }

    private ImmutablePair<Integer, List<DistributionAccount>> addBudgetInfo(final Sheet budget, final List<DistributionAccount> actualsAccounts) {

        int firstAccountRowIndex = -1;
        List<DistributionAccount> results = new ArrayList<>();

        boolean firstFound = false;

        for (Row row: budget) {

            if (row.getCell(0) == null) {
                break;
            }

            if (row.getCell(0).getCellType().equals(CellType.BLANK)) {
                continue;
            }

            String firstColumnValue = null;

            if (row.getCell(0).getCellType().equals(CellType.NUMERIC)) {
                firstColumnValue = String.format("%.0f", row.getCell(0).getNumericCellValue());
            }

            if (row.getCell(0).getCellType().equals(CellType.STRING)) {
                firstColumnValue = String.valueOf(row.getCell(0).getStringCellValue());
            }

            if (BUDGET_ACCOUNT_NAME_REGEX.matcher(firstColumnValue).matches()) {

                if (!firstFound) {
                    firstAccountRowIndex = row.getRowNum();
                    firstFound = true;
                }

                String budgetAccountId = firstColumnValue;

                Optional<DistributionAccount> actualsAccount =
                    actualsAccounts.stream()
                        .filter(account -> account.id().equals(budgetAccountId))
                        .findAny();

                DistributionAccount consolidatedAccount = null;

                if (actualsAccount.isPresent()) {

                    consolidatedAccount = 
                        DistributionAccount.builder()
                            .id(actualsAccount.get().id())
                            .name(actualsAccount.get().name())
                            .actuals(actualsAccount.get().actuals())
                            .existsInActuals(true)
                            .budget(row.getCell(2).getNumericCellValue())
                            .existsInBudget(true)
                            .build();
                } else {

                    consolidatedAccount = 
                        DistributionAccount.builder()
                            .id(firstColumnValue)
                            .name(row.getCell(1).getStringCellValue())
                            .actuals(0)
                            .existsInActuals(false)
                            .budget(row.getCell(2).getNumericCellValue())
                            .existsInBudget(true)
                            .build();
                }

                results.add(consolidatedAccount);

            }
        }

        for (DistributionAccount actualsAccount: actualsAccounts) {

            boolean isAccountMissingFromBudget = 
                results.stream()
                    .filter(budgetAccount -> budgetAccount.id().equals(actualsAccount.id()))
                    .findAny().isEmpty();

            if (isAccountMissingFromBudget) {

                DistributionAccount missingAccount =
                    DistributionAccount.builder()
                        .id(actualsAccount.id())
                        .name(actualsAccount.name())
                        .actuals(actualsAccount.actuals())
                        .existsInActuals(true)
                        .budget(0)
                        .existsInBudget(false)
                        .build();

                results.add(missingAccount);
            }
        }

        results.sort(Comparator.comparing(DistributionAccount::id));

        return ImmutablePair.of(firstAccountRowIndex, results);
    }
}
