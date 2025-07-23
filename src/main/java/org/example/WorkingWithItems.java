package org.example;

import java.util.ArrayList;
import java.util.Scanner;

public class WorkingWithItems {

    private static final WorkingWithExcel workingWithExcel = new WorkingWithExcel();
    private static final Scanner scanner = new Scanner(System.in);
    private static final Size size = new Size();
    private static final MainMenu mainMenu = new MainMenu();


    public boolean createAndAddItem() { // creating and adding item
        ArrayList<Item> items = inputItemData(scanner);
        return addItemsData(items);
    }

    public ArrayList<Item> inputItemData(Scanner scanner) { // getting data from user
        String dateBought;
        while (true) {
            System.out.print("Enter the date you bought this item: ");
            dateBought = scanner.nextLine();
            if (!dateBought.trim().isEmpty()) {
                break;
            } else {
                System.out.println("Date cannot be empty. Please enter a date.");
            }
        }

        String title;
        while (true) {
            System.out.print("Enter the title: ");
            title = scanner.nextLine();
            if (!title.trim().isEmpty()) {
                break;
            } else {
                System.out.println("Title cannot be empty. Please enter a valid title.");
            }
        }

        boolean validInput;
        boolean allPhrasesAreEntered = false;
        ArrayList<Size> sizes = new ArrayList<>();
        System.out.println("To finish entering sizes enter '-'.");
        while(!allPhrasesAreEntered) {
            validInput = false;
            while (!validInput) {
                System.out.print("Enter input ([size], [number]): ");
                String enteredTextSize = scanner.nextLine();
                if(enteredTextSize.equalsIgnoreCase("-")) {
                    allPhrasesAreEntered = true;
                    validInput = true;
                } else {
                    try {
                        Size enteredSize = size.getSizeFromText(enteredTextSize);
                        sizes.add(enteredSize);
                        validInput = true;
                    } catch (Exception e) {
                        System.out.println("Invalid input: " + e.getMessage() +
                                "Please try again.\n");
                    }
                }
            }
        }

        ArrayList<Size> expandedSizes = new ArrayList<>(); // unpacking sizes by quantity
        for (Size s : sizes) {
            for (int i = 0; i < s.getQuantity(); i++) {
                expandedSizes.add(new Size(1, s.getSize()));
            }
        }

        ArrayList<Item> itemsBySizes = new ArrayList<>();
        for (Size size : expandedSizes) {
            double priceBought;
            while (true) {
                System.out.print("Enter the bought price for " + size.getSize() + ": ");
                String priceInput = scanner.nextLine();
                try {
                    priceBought = Double.parseDouble(priceInput);
                    if (priceBought >= 0) {
                        break;
                    } else {
                        System.out.println("Price cannot be negative. Please enter a valid price.");
                    }
                } catch (NumberFormatException e) {
                    System.out.println("Invalid number format. Please enter a valid price.");
                }
            }

            boolean isSold;
            while (true) {
                System.out.print("Is it sold (Y : N)? ");
                String isSoldText = scanner.nextLine();
                if (isSoldText.equalsIgnoreCase("Y")) {
                    isSold = true;
                    break;
                } else if (isSoldText.equalsIgnoreCase("N")) {
                    isSold = false;
                    break;
                } else {
                    System.out.println("Invalid input. Please enter Y or N.");
                }
            }

            double priceSold = 0;
            if (isSold) {
                while (true) {
                    System.out.print("Enter the sold price for " + size.getSize() + ": ");
                    String priceSoldInput = scanner.nextLine();
                    try {
                        priceSold = Double.parseDouble(priceSoldInput);
                        if (priceSold >= 0) {
                            break;
                        } else {
                            System.out.println("Price cannot be negative. Please enter a valid sold price.");
                        }
                    } catch (NumberFormatException e) {
                        System.out.println("Invalid number format. Please enter a valid sold price.");
                    }
                }
            }

            Condition condition;
            while (true) {
                System.out.print(Condition.getAllConditions() + "\nEnter the condition from the list for " + size.getSize() + ": ");
                String conditionInput = scanner.nextLine();
                try {
                    int conditionNumber = Integer.parseInt(conditionInput);
                    condition = Condition.getConditionFromTheList(conditionNumber);
                    if (condition != null) {
                        break;
                    } else {
                        System.out.println("Invalid condition number. Please select a valid condition from the list.");
                    }
                } catch (NumberFormatException e) {
                    System.out.println("Invalid input. Please enter a number corresponding to the condition.");
                }
            }

            if (isSold) {
                itemsBySizes.add(new Item(dateBought, title, size, priceBought, priceSold, condition, true));
            } else {
                itemsBySizes.add(new Item(dateBought, title, size, priceBought, condition, false));
            }
        }



        return itemsBySizes;
    }

    public Item getTheMainItem(ArrayList<Item> items) { // getting the item needs to be added to the main sheet if it has multiply sizes
        if(items.size() == 1) {
            return items.get(0);
        } else if(items.size() > 1) {
            String dateBought = items.get(0).getDateBought();
            String title = items.get(0).getTitle();
            double summedPriceBought = 0;
            double summedPriceSold = 0;
            boolean allItemsSold = true;
            ArrayList<Size> sizes = new ArrayList<>();
            for(Item item : items) {
                sizes.addAll(item.getSizes());
                summedPriceBought += item.getPriceBought();
                if (item.isSold()) {
                    summedPriceSold += item.getPriceSold();
                } else {
                    allItemsSold = false;
                }
            }
            return new Item(dateBought, title, sizes, summedPriceBought, summedPriceSold, allItemsSold);
        }
        return null;
    }

    private boolean addItemsData(ArrayList<Item> items) { //adding all information about items
        if (items == null || items.isEmpty()) {
            return false; // nothing to add
        }
        if(items.size() == 1) {
            boolean isAdded = workingWithExcel.addItemToTheMainSheet(items.get(0));
            workingWithExcel.save();
            mainMenu.launch();
            return isAdded;
        }
        Item theMainItem = getTheMainItem(items);
        boolean mainAdded = workingWithExcel.addItemToTheMainSheet(theMainItem);
        if (!mainAdded) { // securing the program and saving it
            workingWithExcel.save();
            return false;
        }
        int i = 0; // to get first added item
        for (Item item : items) {
            if (i == 0 && !workingWithExcel.isTheFirstItemInRevenuesSheet()) {
                workingWithExcel.clearLastRowInRevenuesSheet();
            }
            i++;
            boolean itemAdded = theMainItem.addItemInSheet(item);
            if (!itemAdded) { // securing the program and saving it
                workingWithExcel.save();
                return false;
            }
        }
        workingWithExcel.fillInTotalData();
        workingWithExcel.save();
        mainMenu.launch();
        return true;
    }
}
