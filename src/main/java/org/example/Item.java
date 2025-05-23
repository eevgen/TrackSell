package org.example;

public class Item {

    private String dateBought;
    private String title;
    private double priceBought;
    private double priceSold;
    private Condition condition;
    private boolean isSold;

    public Item(String dateBought, String title, double priceBought, double priceSold, Condition condition, boolean isSold) {
        this.dateBought = dateBought;
        this.title = title;
        this.priceBought = priceBought;
        this.priceSold = priceSold;
        this.condition = condition;
        this.isSold = isSold;
    }

    public Item(String dateBought, String title, double priceBought, Condition condition, boolean isSold) {
        this.dateBought = dateBought;
        this.title = title;
        this.priceBought = priceBought;
        this.condition = condition;
        this.isSold = isSold;
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
}