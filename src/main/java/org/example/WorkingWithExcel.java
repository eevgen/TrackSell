package org.example;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*; //added dependency in pom.xml from https://mvnrepository.com/artifact/org.apache.poi/poi/5.2.3
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // added dependency from https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/3.9
/* I had an error "ERROR StatusLogger Log4j2 could not find a logging implementation" and found the solution to add 3 .jar files(dependencies) to Project Structure.
Found solution on: https://stackoverflow.com/questions/47881821/error-statuslogger-log4j2-could-not-find-a-logging-implementation */

import java.io.*;
import java.util.Scanner;

public class WorkingWithExcel {

    private static final int HEADER_ROW_INDEX = 0;
    private static final int NUMBER_OF_COLUMNS = 8;
    private static Sheet sheet;
    private static FileOutputStream fileOut;
    private static Workbook workbook = new XSSFWorkbook();
    private CellStyle centeredStyle;
    private static final String FILE_NAME = "trades.xlsx";
    private static final MathCalculations mathCalculations = new MathCalculations();

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

            centeredStyle = workbook.createCellStyle();
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);
            centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            sheet = createSheet("Main");
        } catch (IOException e) {
            throw new RuntimeException("Error in creating/opening file.", e);
        }
    }

    public boolean fillInHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(HEADER_ROW_INDEX);

        CellStyle headerStyle = workbook.createCellStyle(); // Create a style with gray background
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = sheet.getWorkbook().createFont(); // make the font bold
        font.setBold(true);
        headerStyle.setFont(font);

        String[] headers = {
                "Date purchased", "The title", "Quantity", "Sizes",
                "Bought price($)", "Sold Price($)", "Condition", "Is sold", "Profit(%)"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        return true;
    }

    public boolean addItemToTheMainSheet(Item item) {
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);



        if(fillInRaw(item, row)) {
            if (item.getQuantity() > 1) {
                createHyperLink(row.getCell(3), item.getSizesSheet());
            }
            centerAlignRow(row);
            return true;
        }
        return false;
    }

    public boolean addItem(Item item, Sheet sheet) {
        int newRowIndex = sheet.getPhysicalNumberOfRows();
        Row row = sheet.createRow(newRowIndex);

        if(fillInRaw(item, row)) {
            centerAlignRow(row);
            return true;
        }
        return false;
    }

    public void centerAlignRow(Row row) {
        if (row == null) {
            return;
        }

        Workbook workbook = row.getSheet().getWorkbook();

        for (Cell cell : row) {
            CellStyle originalStyle = cell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            newStyle.setAlignment(centeredStyle.getAlignment());
            newStyle.setVerticalAlignment(centeredStyle.getVerticalAlignment());

            cell.setCellStyle(newStyle);
        }
    }


    public boolean fillInRaw(Item item, Row row) {
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

    public void save() {
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
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("Error closing workbook", e);
        }
    }

    public Sheet createSheet(String title) {
        Sheet sheet = workbook.getSheet(title);
        if(sheet == null) {
            Sheet newSheet = workbook.createSheet(title);
            fillInHeaderRow(newSheet);
            return newSheet;
        } else {
            return sheet;
        }
    }
    private boolean createHyperLink(Cell cell, Sheet targetSheet) {
        CreationHelper createHelper = workbook.getCreationHelper(); // got from https://stackoverflow.com/questions/57300034/how-to-use-apache-poi-to-create-excel-hyper-link-that-links-to-long-url
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT); // HyperlinkType.DOCUMENT is only for internal links
        hyperlink.setAddress(String.format("'%s'!%s", targetSheet.getSheetName(), cell.getAddress()));
        cell.setHyperlink(hyperlink);

        CellStyle linkStyle = workbook.createCellStyle(); // got all styling methods from https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/CellStyle.html
        Font linkFont = workbook.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(linkFont);
        cell.setCellStyle(linkStyle); // apply all changes

        return true;
    }

}
