package org.example;

import java.util.ArrayList;
import java.util.InputMismatchException;
import java.util.Scanner;

public class MainMenu {

        private static final ArrayList<String> functions = new ArrayList<>();
        private static final Scanner scanner = new Scanner(System.in);
        private static final WorkingWithItems workingWithItems = new WorkingWithItems();
        private static final WorkingWithCMD workingWithCMD = new WorkingWithCMD();
        private static final WorkingWithExcel workingWithExcel = new WorkingWithExcel();

        public MainMenu() {
            if (functions.isEmpty()) {
                functions.add("Add an item to the sheet");
                functions.add("Save and close file");
            }
        }

        public void launch() {

            workingWithCMD.clearConsole();

            printAllFunctions();

            completeUsersTask();

        }

        public void printAllFunctions() {
            int i = 1;
            for(String function : functions) {
                System.out.printf("%d. %s%n", i, function);
                i++;
            }
        }

        public int getUsersChoice() { // get the number user entered (chose an option)
            boolean isOK = false;
            int choiceNumber = 0;

            while (!isOK) {
                try {
                    System.out.print("\nEnter your choice: ");
                    choiceNumber = scanner.nextInt();

                    if (choiceNumber < 1 || choiceNumber > functions.size()) {
                        System.out.printf("Please enter a number between 1 and %d.%n", functions.size());
                        continue;
                    }

                    isOK = true;

                } catch (InputMismatchException e) {
                    System.out.println("Invalid input. Please enter a valid number.");
                    scanner.nextLine(); // clear invalid input from buffer
                }
            }

            return choiceNumber;
        }

        public void completeUsersTask() { // starting all tasks
            switch (getUsersChoice()) {
                case 1 -> {
                    workingWithCMD.clearConsole();
                    System.out.println(workingWithItems.createAndAddItem() ?
                            "Item has been successfully added." : "There was an error in adding item.");
                }
                case 2 -> {
                    workingWithExcel.save();
                    workingWithExcel.close();
                    System.exit(0); // termination
                }
            }
        }

}
