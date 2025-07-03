package com.sterling.automation.service;

import java.io.File;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.commons.lang3.tuple.ImmutablePair;

public class InputService {

    public ImmutablePair<String, String> getFileBasedOnInput(final Scanner scanner) {

        List<String> filesInDir = getFilesInCurrentDirectory();

        System.out.println("Choices:");
        System.out.println("");

        for (int i = 0; i < filesInDir.size(); i++) {
            System.out.println(String.format("%d - %s", i + 1, filesInDir.get(i)));
        }

        System.out.println("");
        System.out.print("Enter the number of the QuickBooks export: ");
        int quickBooksExportIndex = scanner.nextInt();
        //handle newline
        scanner.nextLine();

        System.out.print("Enter the number of the Budget: ");
        int budgetIndex = scanner.nextInt();
        //handle newline
        scanner.nextLine();

        return ImmutablePair.of(
            filesInDir.get(quickBooksExportIndex - 1), 
            filesInDir.get(budgetIndex - 1)
        );
    }

    private List<String> getFilesInCurrentDirectory() {

        return Stream.of(new File(".").listFiles())
          .filter(file -> !file.isDirectory())
          .filter(file -> {
            return file.getName().endsWith(".xlsx") ||
            file.getName().endsWith(".xls");
            })
          .map(File::getName)
          .collect(Collectors.toList());
    }
}
