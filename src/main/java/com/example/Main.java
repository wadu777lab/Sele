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

public class Main {
    public static void main(String[] args) {
        System.out.println("Starting application...");

        String excelFilePath = args.length > 0 ? args[0] : "src/main/resources/Data.xlsx";
        System.out.println("Excel file path: " + excelFilePath);

        // Get all items from first column of first sheet
        List<String> items = getAllItemsFromFirstColumn(excelFilePath);
        System.out.println("Items from first column:");
        for (int i = 0; i < items.size(); i++) {
            System.out.println((i + 1) + ". " + items.get(i));
        }
    }

    private static void createSampleExcel(String filePath) {
        try {
            Files.createDirectories(Paths.get(filePath).getParent());
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Sheet1");
                Row row = sheet.createRow(0);
                Cell cell = row.createCell(0);
                cell.setCellValue("https://www.google.com");
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                }
            }
        } catch (IOException e) {
            System.err.println("Error creating sample Excel file: " + e.getMessage());
        }
    }

    private static List<String> getAllItemsFromFirstColumn(String filePath) {
        List<String> items = new ArrayList<>();
        if (!Files.exists(Paths.get(filePath))) {
            createSampleExcel(filePath);
        }
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // First sheet
            if (sheet != null) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        Cell cell = row.getCell(0); // First column
                        if (cell != null) {
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

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }
}