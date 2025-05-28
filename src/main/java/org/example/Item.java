package org.example;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

public class Item {


    private final String dateBought;
    private final String title;
    private final ArrayList<Size> sizes;
    private final double priceBought;
    private double priceSold;
    private Condition condition;
    private final boolean isSold;
    private double profitInPercents;
    private final int quantity;
    private Sheet sizesSheet;
    private double netRevenue;

    private static final MathCalculations mathCalculations = new MathCalculations();
    private static final WorkingWithExcel workingWithExcel = new WorkingWithExcel();


    public Item(String dateBought, String title, ArrayList<Size> sizes, double priceBought, double priceSold, boolean isSold) { //for main items
        this.dateBought = dateBought;
        this.title = title;
        this.sizes = sizes;
        quantity = sizes.size();
        if(quantity > 1) {
            createSizesSheet();
        }
        this.priceBought = priceBought;
        this.priceSold = priceSold;
        this.isSold = isSold;
        profitInPercents = mathCalculations.getTheProfitPercentage(priceBought, priceSold, 2);
        netRevenue = mathCalculations.getNetRevenue(priceBought, priceSold);
    }

    public Item(String dateBought, String title, Size size, double priceBought, double priceSold, Condition condition, boolean isSold) { // for sold items
        this.dateBought = dateBought;
        this.title = title;
        ArrayList<Size> sizes = new ArrayList<>();
        sizes.add(size);
        this.sizes = sizes;
        quantity = sizes.size();
        this.priceBought = priceBought;
        this.priceSold = priceSold;
        this.condition = condition;
        this.isSold = isSold;
        profitInPercents = mathCalculations.getTheProfitPercentage(priceBought, priceSold, 2);
        netRevenue = mathCalculations.getNetRevenue(priceBought, priceSold);
    }

    public Item(String dateBought, String title, Size size, double priceBought, Condition condition, boolean isSold) { // for items which are in stock
        this.dateBought = dateBought;
        this.title = title;
        ArrayList<Size> sizes = new ArrayList<>();
        sizes.add(size);
        this.sizes = sizes;
        quantity = sizes.size();
        this.priceBought = priceBought;
        this.condition = condition;
        this.isSold = isSold;
    }


    public String getSizesText() { // returns sizes text and if there is multiply sizes, returns it in format [size]x[quantity]
        if (sizes == null || sizes.isEmpty()) {
            return "";
        }

        ArrayList<Size> groupedSizes = new ArrayList<>();
        for (Size s : sizes) {
            boolean found = false;
            for (Size g : groupedSizes) { // check if the size already exists in the grouped list
                if (g.getSize().equalsIgnoreCase(s.getSize())) {
                    int newQuantity = g.getQuantity() + s.getQuantity();
                    groupedSizes.remove(g); // remove old entry
                    groupedSizes.add(new Size(newQuantity, g.getSize())); // add new entry with updated quantity
                    found = true;
                    break;
                }
            }
            if (!found) {
                groupedSizes.add(new Size(s.getQuantity(), s.getSize())); // if the size wasn't found in the grouped list, just add it
            }
        }

        StringBuilder finalText = new StringBuilder();
        for (Size s : groupedSizes) {
            if (finalText.length() > 0) {
                finalText.append(", ");
            }
            finalText.append(s.getQuantity()).append("x").append(s.getSize());
        }

        return finalText.toString();
    }

    public boolean createSizesSheet() {
        sizesSheet = workingWithExcel.createSheet(title);
        return sizesSheet != null;
    }

    public boolean addItemInSheet(Item item) { //adding item in sheet
        if(sizesSheet == null) {
            if (!createSizesSheet()) { // secure in case it won't be created
                return false;
            }
        }
        return workingWithExcel.addItem(item, sizesSheet);
    }

    public Sheet getSizesSheet() {
        return sizesSheet;
    }

    public String getDateBought() {
        return dateBought;
    }

    public String getTitle() {
        return title;
    }

    public double getPriceBought() {
        return priceBought;
    }

    public double getPriceSold() {
        return priceSold;
    }

    public Condition getCondition() {
        return condition;
    }

    public boolean isSold() {
        return isSold;
    }
    public String getFirstSize() {
        return sizes.get(0).getSize();
    }

    public ArrayList<Size> getSizes() {
        return sizes;
    }

    public double getProfitInPercents() {
        return profitInPercents;
    }

    public int getQuantity() {
        return quantity;
    }

    public double getNetRevenue() {
        return netRevenue;
    }
}