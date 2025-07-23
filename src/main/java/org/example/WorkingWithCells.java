package org.example;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;

public class WorkingWithCells {

    public void createHyperLink(Cell cell, Sheet targetSheet, Workbook workbook) { //creating hyperlink which directs to the target sheet
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
    }

    public void centerAlignRow(Row row) { // center the text in the cell
        if (row == null) return;

        Workbook workbook = row.getSheet().getWorkbook();

        for (Cell cell : row) {
            CellStyle original = cell.getCellStyle();
            CellStyle aligned = workbook.createCellStyle();
            aligned.cloneStyleFrom(original);
            aligned.setAlignment(HorizontalAlignment.CENTER);
            aligned.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(aligned);
        }
    }

    public CellStyle getHeaderCellStyle(Workbook workbook, Sheet sheet) { // returns header cell style
        CellStyle headerStyle = workbook.createCellStyle(); // Create a style with gray background
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = sheet.getWorkbook().createFont(); // make the font bold
        font.setBold(true);
        headerStyle.setFont(font);

        return headerStyle;
    }

}
