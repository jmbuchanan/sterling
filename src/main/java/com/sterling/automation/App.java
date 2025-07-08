package com.sterling.automation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Scanner;

import org.apache.commons.lang3.tuple.ImmutablePair;

import com.sterling.automation.dto.ValidationResponse;
import com.sterling.automation.service.ExcelService;
import com.sterling.automation.service.InputService;
import com.sterling.automation.service.ValidationService;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class App {

    public static void main( String[] args ) {

        System.out.println("-----------------------");
        System.out.println("| Variance Automation |");
        System.out.println("-----------------------");
        System.out.println("");

        InputService inputService = new InputService();
        ValidationService validationService = new ValidationService();
        ExcelService excelService = new ExcelService();

        ValidationResponse response = null;
        Scanner scanner = new Scanner(System.in);
        do {
            ImmutablePair<String, String> files = inputService.getFileBasedOnInput(scanner);
            response = validationService.isValidInputOutput(scanner, files.getLeft(), files.getRight());
            if (response.isValid()) {
                //backup budget file just in case
                backup(files.getRight());
            }
        } while (!response.isValid());

        excelService.process(response);

        scanner.close();
    }

    private static void backup(final String budget) {
        try {
            Files.copy(Paths.get(budget), Paths.get(budget + ".bak"), StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            log.warn("Couldn't create a backup for some reason");
            log.warn(e.getMessage());
        }
    }
}
