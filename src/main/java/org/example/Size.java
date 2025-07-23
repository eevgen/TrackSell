package org.example;


public class Size {

    private int quantity;
    private String size;

    public Size(int quantity, String size) {
        this.quantity = quantity;
        this.size = size;
    }
    public Size() {

    }

    public Size getSizeFromText(String text) {
        // format: [size], [quantity]
        // Split the input string by comma
        String[] parts = text.split(",");

        if (parts.length != 2) {
            throw new IllegalArgumentException("Invalid format. Use 'number, size'");
        }

        size = parts[0].trim(); //removes all spaces

        if (size.isEmpty()) {
            throw new IllegalArgumentException("Size cannot be empty.");
        }

        quantity = Integer.parseInt(parts[1].trim());
        return new Size(quantity, size);
    }

    @Override
    public String toString() {
        return "Size: " + size +
                "\nQuantity: " + quantity;
    }

    public int getQuantity() {
        return quantity;
    }

    public String getSize() {
        return size;
    }
}
