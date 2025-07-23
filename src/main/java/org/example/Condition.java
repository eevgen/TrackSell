package org.example;

public enum Condition {


    NEW_WITH_TAG("New With Tag"),
    NEW_WITHOUT_TAG("New Without Tag"),
    VERY_GOOD("Very Good"),
    GOOD("Good"),
    BAD("Bad");

    final String conditionsText;

    Condition(String conditionsText) {
        this.conditionsText = conditionsText;
    }

    public String getConditionsText() {
        return conditionsText;
    }

    public static String getAllConditions() {
        return "1. " + NEW_WITH_TAG.getConditionsText() +
                "\n2. " + NEW_WITHOUT_TAG.getConditionsText() +
                "\n3. " + VERY_GOOD.getConditionsText() +
                "\n4. " + GOOD.getConditionsText() +
                "\n5. " + BAD.getConditionsText();

    }

    public static Condition getConditionFromTheList(int number) {
        return switch (number) { // recommended format by IDEA
            case 1 -> NEW_WITH_TAG;
            case 2 -> NEW_WITHOUT_TAG;
            case 3 -> VERY_GOOD;
            case 4 -> GOOD;
            case 5 -> BAD;
            default -> null;
        };
    }
}