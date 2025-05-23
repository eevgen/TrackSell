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

    public WorkingWithExcel() {
        try {
            Workbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Warehouse");

            //Creating a header row
            Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
            headerRow.createCell(0).setCellValue("Date purchased");
            headerRow.createCell(1).setCellValue("The title");
            headerRow.createCell(2).setCellValue("Costed");
            headerRow.createCell(3).setCellValue("Status");

            Scanner scanner = new Scanner(System.in);
            Scanner scanner1 = new Scanner(System.in);
            System.out.print("Enter date the item was purchased: ");
            String date = scanner.nextLine();
            System.out.print("Enter date the items title: ");
            String title = scanner.nextLine();
            System.out.print("Enter date the items price: ");
            double cost = scanner1.nextDouble();
            System.out.print("Enter the status: ");
            String status = scanner.nextLine();

            Row firstItem = sheet.createRow(HEADER_ROW_INDEX + 1);
            firstItem.createCell(0).setCellValue(date);
            firstItem.createCell(1).setCellValue(title);
            firstItem.createCell(2).setCellValue(cost);
            firstItem.createCell(3).setCellValue(status);


            FileOutputStream fileOut = new FileOutputStream("employees.xlsx"); //creating xlsx file
            workbook.write(fileOut);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
}
