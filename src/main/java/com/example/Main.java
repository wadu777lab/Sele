package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.ArrayList;

/**
 * Main class for reading and processing Excel files using Apache POI.
 * This application reads data from the first column of the first sheet in an Excel file
 * and prints the items to the console. If the file doesn't exist, it creates a sample one.
 */
public class Main {
    /**
     * Main method - entry point of the application.
     * @param args command line arguments, first argument can be the Excel file path
     */
    public static void main(String[] args) {
        System.out.println("Starting application.");

        // Determine the Excel file path from arguments or use default
        String excelFilePath = args.length > 0 ? args[0] : "src/main/resources/data.xlsx";
        System.out.println("Excel file path: " + excelFilePath);

        // Retrieve all items from the first column of the first sheet
        List<String> items = getAllItemsFromFirstColumn(excelFilePath);
        System.out.println("Items from first column:");
        for (String item : items) {
            System.out.println(item);
        }
    }

    /**
     * Creates a sample Excel file with a single cell containing a URL.
     * @param filePath the path where the Excel file should be created
     */
    private static void createSampleExcel(String filePath) {
        try {
            // Ensure the parent directory exists
            Files.createDirectories(Paths.get(filePath).getParent());
            // Create a new Excel workbook
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Sheet1");
                Row row = sheet.createRow(0);
                Cell cell = row.createCell(0);
                cell.setCellValue("https://www.google.com");
                // Write the workbook to the file
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                }
            }
        } catch (IOException e) {
            System.err.println("Error creating sample Excel file: " + e.getMessage());
        }
    }

    /**
     * Reads all items from the first column of the first sheet in the Excel file.
     * If the file doesn't exist, creates a sample file first.
     * @param filePath the path to the Excel file
     * @return a list of strings from the first column
     */
    private static List<String> getAllItemsFromFirstColumn(String filePath) {
        List<String> items = new ArrayList<>();
        // Check if file exists, create sample if not
        if (!Files.exists(Paths.get(filePath))) {
            createSampleExcel(filePath);
        }
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Access the first sheet
            if (sheet != null) {
                // Iterate through all rows
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        Cell cell = row.getCell(0); // Get the first column cell
                        if (cell != null) {
                            // Convert cell value to string and add to list
                            items.add(getCellValueAsString(cell));
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        }
        return items;
    }

    /**
     * Converts a cell's value to a string based on its type.
     * @param cell the Excel cell to convert
     * @return the string representation of the cell's value
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        // Handle different cell types
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Check if it's a date
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}