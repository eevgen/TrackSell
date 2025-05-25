package org.example;

import java.util.Scanner;

public class Main {
    private static final WorkingWithItems workingWithItems = new WorkingWithItems();
    private static final Size size = new Size();

    public static void main(String[] args) {
        size.getSizeFromText("XL, 2");
        workingWithItems.createAndAddItem();
    }
}