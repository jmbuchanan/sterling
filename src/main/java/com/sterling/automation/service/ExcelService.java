package com.sterling.automation.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.sterling.automation.domain.DistributionAccount;

public class ExcelService {

    private static final String HEADER_ROW_INDICATOR = "Distribution account";

    private static final int BUDGET_LOT_NAME_ROW = 0;
    private static final int BUDGET_LOT_NAME_COLUMN = 3;

    //matches: "   123456 EXAMPLE ACCOUNT  "
    private static final Pattern ACTUALS_ACCOUNT_NAME_REGEX = Pattern.compile("^\\s*\\d{6}\\s.*");
    private static final Pattern BUDGET_ACCOUNT_NAME_REGEX = Pattern.compile("^\\d{6}$");

    public void process(final String budget, final String actuals) {

        Sheet budgetSheet = loadWorkbook(budget).getSheetAt(0);

        String lotName = 
            budgetSheet
                .getRow(BUDGET_LOT_NAME_ROW)
                .getCell(BUDGET_LOT_NAME_COLUMN)
                .getStringCellValue();

        Sheet actualsSheet = loadWorkbook(actuals).getSheetAt(0);

        List<DistributionAccount> actualsAccounts = getAccountsFromActuals(actualsSheet, lotName);

        List<DistributionAccount> consolidatedAccounts = addBudgetInfo(budgetSheet, actualsAccounts);

        Workbook wb = new HSSFWorkbook();
        wb.createSheet();

        Sheet outputSheet = wb.getSheetAt(0);

        try  (OutputStream fileOut = new FileOutputStream("output.xlsx")) {

            CellStyle missingFromActualsCellStyle = wb.createCellStyle();
            missingFromActualsCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            missingFromActualsCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());

            CellStyle missingFromBudgetCellStyle = wb.createCellStyle();
            missingFromBudgetCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            missingFromBudgetCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

            for (int i = 0; i < consolidatedAccounts.size(); i++) {

                Row row = outputSheet.createRow(i);
                DistributionAccount account = consolidatedAccounts.get(i);

                row.createCell(0).setCellValue(Float.valueOf(account.id()));
                row.createCell(1).setCellValue(account.name());
                row.createCell(2).setCellValue(account.budget());
                row.createCell(3).setCellValue(account.actuals());
                row.createCell(4).setCellFormula(String.format("$C%d-$D%d", i + 1, i + 1));
                row.createCell(5).setCellValue(String.format("Found in actuals: %b", account.existsInActuals()));
                row.createCell(6).setCellValue(String.format("Found in budget: %b", account.existsInBudget()));

                if (account.existsInBudget() && !account.existsInActuals()) {
                    row.getCell(0).setCellStyle(missingFromActualsCellStyle);
                }

                if (account.existsInActuals() && !account.existsInBudget()) {
                    row.getCell(0).setCellStyle(missingFromBudgetCellStyle);
                }
            }

            wb.write(fileOut);
            wb.close();

        } catch (IOException e) {
            System.err.println(e.getMessage());
        }
    }

    private List<DistributionAccount> getAccountsFromActuals(final Sheet actualsSheet, final String lotName) {

        int headerRow = -1;

        for (Row row : actualsSheet) {
            if (HEADER_ROW_INDICATOR.equals(row.getCell(0).getStringCellValue())) {
                headerRow = row.getRowNum();
            }
        }

        if (headerRow < 0) {
            System.err.println("No header found in actuals");
        }

        Row actualsLotNameRow = actualsSheet.getRow(headerRow);

        int columnOfActuals = -1; 

        for (Cell cell: actualsLotNameRow) {
            if (lotName.equals(cell.getStringCellValue())) {
                columnOfActuals = cell.getColumnIndex();
            }
        }

        if (columnOfActuals < 0) {
            System.err.println("No matching account found in actuals");
        }

        List<DistributionAccount> actualsAccounts = new ArrayList<>();

        for (Row row: actualsSheet) {

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

    private List<DistributionAccount> addBudgetInfo(final Sheet budget, final List<DistributionAccount> actualsAccounts) {

        List<DistributionAccount> results = new ArrayList<>();

        for (Row row: budget) {

            if (row.getCell(0) == null) {
                System.out.println("Rows first cell is null");
                break;
            }

            if (row.getCell(0).getCellType().equals(CellType.BLANK)) {
                System.out.println("Rows first cell is blank");
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

        results.sort((a1, a2) -> a1.id().compareTo(a2.id()));

        return results;
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
