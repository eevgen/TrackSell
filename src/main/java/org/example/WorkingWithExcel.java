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
    private static final int NUMBER_OF_COLUMNS = 6;
    private static Sheet sheet;
    private static FileOutputStream fileOut;
    private static final Workbook workbook = new XSSFWorkbook();

    public WorkingWithExcel() {
        try {
            sheet = workbook.createSheet("Warehouse");
            fileOut = new FileOutputStream("employees.xlsx"); //creating xlsx file

            //Creating a header row
            Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
            headerRow.createCell(0).setCellValue("Date purchased");
            headerRow.createCell(1).setCellValue("The title");
            headerRow.createCell(2).setCellValue("Bought price");
            headerRow.createCell(3).setCellValue("Sold Price");
            headerRow.createCell(4).setCellValue("Condition");
            headerRow.createCell(5).setCellValue("Is sold");

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }


    }

    public boolean addItem(Item item) {
        Row headerRow = sheet.createRow(lastRowIndex + 1);
        headerRow.createCell(0).setCellValue(item.getDateBought());
        headerRow.createCell(1).setCellValue(item.getTitle());
        headerRow.createCell(2).setCellValue(item.getPriceBought());
        if(item.isSold()){
            headerRow.createCell(3).setCellValue(item.getPriceSold());
        }
        headerRow.createCell(4).setCellValue(item.getCondition().getConditionsText());
        headerRow.createCell(5).setCellValue(item.isSold());

        saveAndClose();
        return true;
    }

    public void saveAndClose() {
        try {
            for (int i = 0; i <= NUMBER_OF_COLUMNS; i++) { //auto size all columns for better look
                sheet.autoSizeColumn(i);
            }

            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("Failed to save and close Excel file", e);
        }
    }
}
