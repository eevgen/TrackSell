package org.example;

import org.apache.poi.ss.usermodel.*; //added dependency in pom.xml from https://mvnrepository.com/artifact/org.apache.poi/poi/5.2.3
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // added dependency from https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/3.9
/* I had an error "ERROR StatusLogger Log4j2 could not find a logging implementation" and found the solution to add 2 .jar files(dependencies) to Project Structure.
They are located in the file JarDependencies in this project.
Found solution on: https://stackoverflow.com/questions/47881821/error-statuslogger-log4j2-could-not-find-a-logging-implementation */

import java.io.*;

public class WorkingWithExcel {

    private static final int HEADER_ROW_INDEX = 0;
    private static final int NUMBER_OF_COLUMNS = 8;
    private static Sheet sheet;
    private static Sheet revenuesSheet;
    private static Workbook workbook = new XSSFWorkbook();

    private static final String FILE_NAME = "trades.xlsx";
    private static final String[] MAIN_SHEET_HEADERS = {
            "Date purchased", "The title", "Quantity", "Sizes",
            "Bought price($)", "Sold Price($)", "Condition", "Is sold", "Profit(%)"
    };
    private static final String MAIN_SHEET_TITLE = "Main";
    private static final String REVENUES_SHEET_TITLE = "Revenues";
    private static final String[] REVENUES_SHEET_HEADERS = {
            "The title", "Size", "Bought price($)", "Sold Price($)", "Profit($)", "Profit(%)"
    };
    private static final WorkingWithCells workingWithCells = new WorkingWithCells();

    public WorkingWithExcel() {
        try {
            File file = new File(FILE_NAME);

            if (file.exists() && file.length() > 0) { // if file is already created
                FileInputStream fileInputStream = new FileInputStream(file);
                workbook = new XSSFWorkbook(fileInputStream);
                fileInputStream.close();
            } else {
                workbook = new XSSFWorkbook();
            }

            sheet = createSheet(MAIN_SHEET_TITLE);
            revenuesSheet = createSheet(REVENUES_SHEET_TITLE);
        } catch (IOException e) {
            throw new RuntimeException("Error in creating/opening file.", e);
        }
    }

    public void fillInTotalData() {
        int firstDataRow = 1; // Skip header row (row 0)
        int lastRow = revenuesSheet.getLastRowNum();
        int targetRow = lastRow + 1;


        Row totalRow = revenuesSheet.createRow(targetRow); // create the total row at the end

        Cell labelCell = totalRow.createCell(0); // column A
        CellStyle headerStyle = workingWithCells.getHeaderCellStyle(workbook, revenuesSheet);
        labelCell.setCellValue("Total");

        int[] columnsToSum = {2, 3, 4}; // columns C (index 2), D (index 3), and E (index 4)

        for (int colIndex : columnsToSum) {
            Cell totalCell = totalRow.createCell(colIndex);

            char colLetter = (char) ('A' + colIndex); // excel columns are A=1, so we convert index to letter

            totalCell.setCellFormula(String.format("SUM(%c%d:%c%d)",
                    colLetter, firstDataRow + 1, colLetter, lastRow + 1));
        }
        for (int i = 0; i < REVENUES_SHEET_HEADERS.length; i++) { // style all cells in raw
            Cell cell = totalRow.getCell(i);
            if(cell == null) {
                cell = totalRow.createCell(i);
            }
            cell.setCellStyle(headerStyle);
        }
        workingWithCells.centerAlignRow(totalRow);
    }

    public void fillInHeaderRow(Sheet sheet, String[] headers) { // filling in the header row (the first one) and styling it
        Row headerRow = sheet.createRow(HEADER_ROW_INDEX);

        CellStyle headerStyle = workingWithCells.getHeaderCellStyle(workbook, sheet);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

    }

    public boolean addItemToTheMainSheet(Item item) { // adding item to the main sheet
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);

        if(fillInRaw(item, row)) {
            if (item.getQuantity() > 1) {
                workingWithCells.createHyperLink(row.getCell(3), item.getSizesSheet(), workbook);
            } else {
                if (item.getPriceSold() > 0) {
                    int newRevenuesRowIndex = revenuesSheet.getPhysicalNumberOfRows() - 1;
                    Row revenuesRow = revenuesSheet.createRow(newRevenuesRowIndex);
                    fillInRevenuesRaw(item, revenuesRow);
                    workingWithCells.centerAlignRow(revenuesRow);
                }
            }
            workingWithCells.centerAlignRow(row);
            return true;
        }
        return false;
    }

    public boolean addItem(Item item, Sheet sheet) { // adding item to the sheet which is entered
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);

        if(fillInRaw(item, row)) {
            workingWithCells.centerAlignRow(row);
            if(item.getPriceSold() > 0) {
                int newRevenuesRowIndex = revenuesSheet.getPhysicalNumberOfRows();
                Row revenuesRow = revenuesSheet.createRow(newRevenuesRowIndex);
                fillInRevenuesRaw(item, revenuesRow);
                workingWithCells.centerAlignRow(revenuesRow);
            }
            return true;
        }
        return false;
    }

    public void clearLastRowInRevenuesSheet() { //clearing total profit raw for adding multiply sizes
        int lastRowNum = revenuesSheet.getLastRowNum();
        Row lastRow = revenuesSheet.getRow(lastRowNum);

        if (lastRow != null) {
            revenuesSheet.removeRow(lastRow); // Remove the row entirely
        }
    }



    public boolean fillInRaw(Item item, Row row) { // filling in raw except adding hyperlink to sizes
        if (item == null || row == null) {
            return false;
        }

        try {
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
        } catch (Exception e) {
            return false;
        }
    }

    public void fillInRevenuesRaw(Item item, Row row) { // filling in raw in revenues sheet
        row.createCell(0).setCellValue(item.getTitle());
        row.createCell(1).setCellValue(item.getFirstSize());
        row.createCell(2).setCellValue(item.getPriceBought());
        row.createCell(3).setCellValue(item.getPriceSold());
        row.createCell(4).setCellValue(item.getNetRevenue());
        row.createCell(5).setCellValue(item.getProfitInPercents());
    }

    public void save() { // saving file
        try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
            for (Sheet currentSheet : workbook) {
                for (int i = 0; i <= NUMBER_OF_COLUMNS; i++) {
                    currentSheet.autoSizeColumn(i);
                }
            }
            workbook.write(fileOut);
        } catch (IOException e) {
            throw new RuntimeException("Error saving file", e);
        }
    }
    public void close() { //closing file
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("Error closing workbook", e);
        }
    }

    public Sheet createSheet(String title) { //creating sheet with entered title
        Sheet sheet = workbook.getSheet(title);
        if(sheet == null) {
            Sheet newSheet = workbook.createSheet(title);
            if(title.equals(REVENUES_SHEET_TITLE)) {
                fillInHeaderRow(newSheet, REVENUES_SHEET_HEADERS);
            } else {
                fillInHeaderRow(newSheet, MAIN_SHEET_HEADERS);
            }
            return newSheet;
        } else {
            return sheet;
        }
    }
    public boolean isTheFirstItemInRevenuesSheet() { // checking if the sheet has only header row for not clearing it while adding items with multiply sizes
        return revenuesSheet.getPhysicalNumberOfRows() == 1;
    }
}
