package org.example;


/*
 * Project Description:
 *
 * This project is designed to assist resellers in managing their inventory and tracking financial performance using Excel sheets.
 * It simplifies the process of recording and organizing item details such as:
 *  - Purchase date
 *  - Item title
 *  - Available sizes
 *  - Purchase price
 *  - Sale status (sold or not)
 *  - Sale price (if sold)
 *  - Item condition
 *
 * The application automatically calculates:
 *  - Net profit for each item
 *  - Profit percentage
 *
 * Features:
 *  - Automatically generates a dedicated Excel sheet for each item, especially when multiple sizes are involved.
 *  - Creates a summary sheet with an overview of revenue, net income, and profit margins.
 *  - Helps resellers keep track of their stock status and financial metrics in a clear, organized way.
 *
 * This tool aims to save time and reduce manual effort in managing small resale businesses.
 */
// Project Description text was formatted by https://chatgpt.com


public class Main {
    private static final MainMenu mainMenu = new MainMenu();

    public static void main(String[] args) {

        mainMenu.launch();
    }
}