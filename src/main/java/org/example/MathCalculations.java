package org.example;

public class MathCalculations {

    public double getTheProfitPercentage(double boughtPrice, double soldPrice, int roundValue) { // returns profit in percents
        return roundToDecimalPlaces((soldPrice/boughtPrice - 1) * 100, roundValue);
    }

    public double roundToDecimalPlaces(double value, int places) { // rounds number to decimal places
        if (places < 0) throw new IllegalArgumentException("Decimal places can't be negative.");
        double scale = Math.pow(10, places);

        /*
        Step 1: Multiply the value by the scale to shift decimal point
        Step 2: Use Math.round() to round the shifted number
        Step 3: Divide back by the scale to return to original range
         */
        return Math.round(value * scale) / scale;
    }

    public double getNetRevenue(double boughtPrice, double soldPrice) { // returning net revenue
        return soldPrice - boughtPrice;
    }
}
