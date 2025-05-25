package org.example;

import org.apache.poi.ss.usermodel.*; //added dependency in pom.xml from https://mvnrepository.com/artifact/org.apache.poi/poi/5.2.3
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // added dependency from https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/3.9
/* I had an error "ERROR StatusLogger Log4j2 could not find a logging implementation" and found the solution to add 3 .jar files(dependencies) to Project Structure.
Found solution on: https://stackoverflow.com/questions/47881821/error-statuslogger-log4j2-could-not-find-a-logging-implementation */

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class WorkingWithExcel {

    private static final int HEADER_ROW_INDEX = 0;
    private static int lastRowIndex = HEADER_ROW_INDEX;
    private static final int NUMBER_OF_COLUMNS = 8;
    private static Sheet sheet;
    private static FileOutputStream fileOut;
    private static final Workbook workbook = new XSSFWorkbook();
    private static final MathCalculations mathCalculations = new MathCalculations();

    public WorkingWithExcel() {
        try {
            sheet = workbook.getSheet("Warehouse");
            if (sheet == null) {
                sheet = workbook.createSheet("Warehouse");
                fillInHeaderRow(sheet);
            }
            fileOut = new FileOutputStream("employees.xlsx"); //creating xlsx file

            //Creating a header row
            fillInHeaderRow(sheet);

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }


    }

    public boolean fillInHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
        headerRow.createCell(0).setCellValue("Date purchased");
        headerRow.createCell(1).setCellValue("The title");
        headerRow.createCell(2).setCellValue("Quantity");
        headerRow.createCell(3).setCellValue("Sizes");
        headerRow.createCell(4).setCellValue("Bought price");
        headerRow.createCell(5).setCellValue("Sold Price");
        headerRow.createCell(6).setCellValue("Condition");
        headerRow.createCell(7).setCellValue("Is sold");
        headerRow.createCell(8).setCellValue("Profit");
        return true;
    }

    public boolean addItemToTheMainSheet(Item item) {
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);

        row.createCell(0).setCellValue(item.getDateBought());
        row.createCell(1).setCellValue(item.getTitle());
        row.createCell(2).setCellValue(item.getQuantity());
        row.createCell(3).setCellValue(item.getSizesText());
        row.createCell(4).setCellValue(item.getPriceBought());
        row.createCell(5).setCellValue(item.getPriceSold());
        row.createCell(6).setCellValue(item.getCondition() != null ? item.getCondition().getConditionsText() : "");
        row.createCell(7).setCellValue(item.isSold() ? "Yes" : "No");
        row.createCell(8).setCellValue(item.getProfitInPercents());

        return true;
    }

    public boolean addItem(Item item, Sheet sheet) {
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);

        row.createCell(0).setCellValue(item.getDateBought());
        row.createCell(1).setCellValue(item.getTitle());
        row.createCell(2).setCellValue(item.getQuantity());
        row.createCell(3).setCellValue(item.getSizesText());
        row.createCell(4).setCellValue(item.getPriceBought());
        row.createCell(5).setCellValue(item.getPriceSold());
        row.createCell(6).setCellValue(item.getCondition() != null ? item.getCondition().getConditionsText() : "");
        row.createCell(7).setCellValue(item.isSold() ? "Yes" : "No");
        row.createCell(8).setCellValue(item.getProfitInPercents());

        return true;
    }

    public void saveAndClose() {
        try {
            for (Sheet currentSheet : workbook) { //auto sizing all sheets
                for (int i = 0; i <= NUMBER_OF_COLUMNS; i++) {
                    currentSheet.autoSizeColumn(i);
                }
            }

            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("Failed to save and close Excel file", e);
        }
    }

    public Sheet createSheet(String title) {
        return workbook.createSheet(title);
    }
}
