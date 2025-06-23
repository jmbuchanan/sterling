package com.sterling.automation;

import java.util.Scanner;

import org.apache.commons.lang3.tuple.ImmutablePair;

import com.sterling.automation.dto.ValidationResponse;
import com.sterling.automation.service.ExcelService;
import com.sterling.automation.service.InputService;
import com.sterling.automation.service.ValidationService;

public class App {

    public static void main( String[] args ) {

        System.out.println("-----------------------");
        System.out.println("| Variance Automation |");
        System.out.println("-----------------------");
        System.out.println("");

        InputService inputService = new InputService();
        ValidationService validationService = new ValidationService();
        ExcelService excelService = new ExcelService();

        boolean isInputCorrect = false;

        ValidationResponse response = null;
        Scanner scanner = new Scanner(System.in);
        do {
            ImmutablePair<String, String> files = inputService.getFileBasedOnInput(scanner);
            response = validationService.isValidInputOutput(scanner, files.getLeft(), files.getRight());
        } while (!isInputCorrect);

        excelService.process(response);

        scanner.close();
    }
}
